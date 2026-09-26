import { PDFDocument, PDFDict, PDFName, PDFString, PDFHexString, PDFStream } from 'pdf-lib';

/** Fail closed for detached/ambiguous widgets. This adapter never repairs field trees. */
export function verifyField(doc: PDFDocument, name: string, expected: string | boolean, requireAppearance = true): boolean {
  try {
    const fields = doc.getForm().getFields().filter(f => f.getName() === name);
    if (fields.length !== 1) return false;
    const field = fields[0]!;
    const actual = typeof expected === 'boolean' ? doc.getForm().getCheckBox(name).isChecked() : doc.getForm().getTextField(name).getText() ?? '';
    if (actual !== expected) return false;
    const canonical = new Set(field.acroField.getWidgets().map(w => w.dict));
    if (!canonical.size) return false;
    const found = new Set<PDFDict>();
    for (const page of doc.getPages()) for (const ref of page.node.Annots()?.asArray() ?? []) {
      const widget = doc.context.lookup(ref);
      if (!(widget instanceof PDFDict) || widget.get(PDFName.of('Subtype'))?.toString() !== '/Widget') continue;
      const chain: PDFDict[] = [];
      let current: PDFDict | undefined = widget;
      while (current) {
        if (chain.includes(current) || chain.length > 100) return false;
        chain.push(current);
        const parent: unknown = doc.context.lookup(current.get(PDFName.of('Parent')));
        current = parent instanceof PDFDict ? parent : undefined;
      }
      const qualified = [...chain].reverse().flatMap(d => {
        const t = d.lookup(PDFName.of('T'));
        return t instanceof PDFString || t instanceof PDFHexString ? [t.decodeText()] : [];
      }).join('.');
      if (qualified !== name && !canonical.has(widget)) continue;
      if (qualified !== name || !canonical.has(widget) || !chain.includes(field.acroField.dict) || found.has(widget)) return false;
      found.add(widget);
      const value = chain.map(d => d.lookup(PDFName.of('V'))).find(v => v !== undefined);
      const effective = typeof expected === 'boolean'
        ? value instanceof PDFName && value.decodeText() !== 'Off'
        : value instanceof PDFString || value instanceof PDFHexString ? value.decodeText() : '';
      if (effective !== expected) return false;
      if (requireAppearance) {
        const ap = widget.lookup(PDFName.of('AP'));
        if (!(ap instanceof PDFDict)) return false;
        const normal = ap.lookup(PDFName.of('N'));
        const state = widget.get(PDFName.of('AS'));
        const appearance = normal instanceof PDFDict && state instanceof PDFName ? normal.lookup(state) : normal;
        if (!(appearance instanceof PDFStream) || appearance.getContents().length === 0) return false;
        if (typeof expected === 'boolean' && (!(state instanceof PDFName) || (state.decodeText() !== 'Off') !== expected)) return false;
      }
    }
    return found.size === canonical.size;
  } catch { return false; }
}
