import { createHash } from 'node:crypto';
import { describe, expect, it } from 'vitest';
import { createPdfOfficeModule } from '../src/index';

function minimalPdf(text: string | string[] = 'Hello PDF', pageCount = 1) {
  const objects = ['<< /Type /Catalog /Pages 2 0 R >>', '', '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'];
  const pageIds = [];
  for (let page = 0; page < pageCount; page++) {
    const items = Array.isArray(text) ? text : [text];
    const stream = `BT /F1 12 Tf 72 720 Td ${items.map((item, index) => `${index ? '50 0 Td ' : ''}(${item}) Tj`).join(' ')} ET`;
    const contentId = objects.length + 1, pageId = contentId + 1;
    objects.push(`<< /Length ${Buffer.byteLength(stream)} >>\nstream\n${stream}\nendstream`);
    objects.push(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 3 0 R >> >> /Contents ${contentId} 0 R >>`);
    pageIds.push(`${pageId} 0 R`);
  }
  objects[1] = `<< /Type /Pages /Kids [${pageIds.join(' ')}] /Count ${pageCount} >>`;
  let body = '%PDF-1.4\n';
  const offsets = [0];
  for (let i = 0; i < objects.length; i++) { offsets.push(Buffer.byteLength(body)); body += `${i + 1} 0 obj\n${objects[i]}\nendobj\n`; }
  const xref = Buffer.byteLength(body);
  body += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  for (const offset of offsets.slice(1)) body += `${String(offset).padStart(10, '0')} 00000 n \n`;
  body += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return new Uint8Array(Buffer.from(body, 'latin1'));
}

function fixture(bytes = minimalPdf(), label = 'native.pdf') {
  const digest = createHash('sha256').update(bytes).digest('hex');
  const ref = { id: `sha256:${digest}`, uri: `file:///managed/${digest}/${label}`, sha256: digest, sizeBytes: bytes.length,
    mediaType: 'application/pdf', label };
  const files = new Map([[ref.uri, bytes]]);
  return { ref, files, module: createPdfOfficeModule({ artifactStore: { async read(source) {
    const value = files.get(source.uri); if (!value) throw new Error('not found'); return value;
  } } }) };
}

describe('pdf-office', () => {
  it('inspects and extracts native text and page geometry without claiming OCR or reading order', async () => {
    const { ref, module } = fixture();
    const inspected = await module.handlers.inspect({ requestId: 'inspect-native', operation: 'inspect', artifactRef: ref }) as any;
    expect(inspected.result).toMatchObject({ format: 'pdf', mediaType: 'application/pdf', pageCount: 1, coverage: 'partial' });
    const parsed = await module.handlers.execute({ requestId: 'extract-native', operation: 'execute', artifactRef: ref }) as any;
    expect(parsed.result).toMatchObject({ source: { id: ref.id, sha256: ref.sha256 }, coordinates: 'pdf-user-space-points',
      coverage: { text: 'native-text-only', readingOrder: 'not-inferred', tables: 'not-inferred', ocr: 'not-performed' }, pageCount: 1 });
    expect(parsed.result.pages[0]).toMatchObject({ pageNumber: 1, widthPt: 612, heightPt: 792 });
    expect(parsed.result.pages[0].text).toContain('Hello PDF');
    await module.dispose();
  });

  it('verifies expected native text and page count', async () => {
    const { ref, module } = fixture();
    const verified = await module.handlers.verify({ requestId: 'verify-native', operation: 'verify', artifactRef: ref,
      payload: { expectedPageCount: 1, textIncludes: ['Hello PDF'] } }) as any;
    expect(verified.result).toMatchObject({ ok: true, summary: { total: 3, passed: 3, failed: 0, skipped: 0 } });
    await module.dispose();
  });

  it('rejects non-PDFs, malformed PDF data, stale hashes, and text budget overflow', async () => {
    const { ref: nonPdf, module: nonPdfModule } = fixture(new Uint8Array(Buffer.from('not a pdf')), 'bad.docx');
    await expect(nonPdfModule.handlers.inspect({ requestId: 'non-pdf', operation: 'inspect', artifactRef: nonPdf }))
      .rejects.toMatchObject({ code: 'FORMAT_MISMATCH' });
    const malformed = fixture(new Uint8Array(Buffer.from('%PDF-1.7\nnot really a PDF')), 'malformed.pdf');
    await expect(malformed.module.handlers.execute({ requestId: 'malformed', operation: 'execute', artifactRef: malformed.ref }))
      .rejects.toMatchObject({ code: 'PARSE_FAILED' });
    const stale = fixture();
    await expect(stale.module.handlers.inspect({ requestId: 'stale', operation: 'inspect', artifactRef: { ...stale.ref, sha256: '0'.repeat(64) } }))
      .rejects.toMatchObject({ code: 'ARTIFACT_INTEGRITY_FAILED' });
    const bounded = fixture();
    const boundedModule = createPdfOfficeModule({ artifactStore: { read: async () => bounded.files.get(bounded.ref.uri)! }, limits: { maxTextChars: 3 } });
    await expect(boundedModule.handlers.execute({ requestId: 'budget', operation: 'execute', artifactRef: bounded.ref }))
      .rejects.toMatchObject({ code: 'LIMIT_EXCEEDED' });
    const overInput = fixture();
    const inputLimitModule = createPdfOfficeModule({ artifactStore: { read: async () => overInput.files.get(overInput.ref.uri)! },
      limits: { maxInputBytes: overInput.ref.sizeBytes! - 1 } });
    await expect(inputLimitModule.handlers.execute({ requestId: 'input-budget', operation: 'execute', artifactRef: overInput.ref }))
      .rejects.toMatchObject({ code: 'LIMIT_EXCEEDED' });
    const overPages = fixture(minimalPdf('Hello PDF', 2));
    const pageLimitModule = createPdfOfficeModule({ artifactStore: { read: async () => overPages.files.get(overPages.ref.uri)! }, limits: { maxPages: 1 } });
    await expect(pageLimitModule.handlers.execute({ requestId: 'page-budget', operation: 'execute', artifactRef: overPages.ref }))
      .rejects.toMatchObject({ code: 'LIMIT_EXCEEDED' });
    const tooManyItems = fixture(minimalPdf(['first', 'second']));
    const itemLimitModule = createPdfOfficeModule({ artifactStore: { read: async () => tooManyItems.files.get(tooManyItems.ref.uri)! }, limits: { maxTextItems: 1 } });
    await expect(itemLimitModule.handlers.execute({ requestId: 'node-budget', operation: 'execute', artifactRef: tooManyItems.ref }))
      .rejects.toMatchObject({ code: 'LIMIT_EXCEEDED' });
    const outputLimitModule = createPdfOfficeModule({ artifactStore: { read: async () => overInput.files.get(overInput.ref.uri)! }, limits: { maxOutputJsonBytes: 1 } });
    await expect(outputLimitModule.handlers.execute({ requestId: 'output-budget', operation: 'execute', artifactRef: overInput.ref }))
      .rejects.toMatchObject({ code: 'LIMIT_EXCEEDED' });
    const encryptedBytes = Buffer.from('JVBERi0xLjMKJeLjz9MKMSAwIG9iago8PAovUHJvZHVjZXIgPDI0ZjdiNzEyZGY+Cj4+CmVuZG9iagoyIDAgb2JqCjw8Ci9UeXBlIC9QYWdlcwovQ291bnQgMQovS2lkcyBbIDQgMCBSIF0KPj4KZW5kb2JqCjMgMCBvYmoKPDwKL1R5cGUgL0NhdGFsb2cKL1BhZ2VzIDIgMCBSCj4+CmVuZG9iago0IDAgb2JqCjw8Ci9UeXBlIC9QYWdlCi9SZXNvdXJjZXMgPDwKPj4KL01lZGlhQm94IFsgMC4wIDAuMCA2MTIgNzkyIF0KL1BhcmVudCAyIDAgUgo+PgplbmRvYmoKNSAwIG9iago8PAovViAyCi9SIDMKL0xlbmd0aCAxMjgKL1AgNDI5NDk2NzI5MgovRmlsdGVyIC9TdGFuZGFyZAovTyA8MGU1MjI5MjVhM2U0ZTg3NGMzY2ZhY2JlZjUxMWE3M2FjNGVjMmJkODY1ZGNkM2Q0NjI3NjE0OTE3YWJmZDdlND4KL1UgPDAxODBmY2VkMTZhNjA0MjJmNDJjNDhhNTMzZjMzYjRlMjhiZjRlNWU0ZTc1OGE0MTY0MDA0ZTU2ZmZmYTAxMDg+Cj4+CmVuZG9iagp4cmVmCjAgNgowMDAwMDAwMDAwIDY1NTM1IGYgCjAwMDAwMDAwMTUgMDAwMDAgbiAKMDAwMDAwMDA1OSAwMDAwMCBuIAowMDAwMDAwMTE4IDAwMDAwIG4gCjAwMDAwMDAxNjcgMDAwMDAgbiAKMDAwMDAwMDI2MSAwMDAwMCBuIAp0cmFpbGVyCjw8Ci9TaXplIDYKL1Jvb3QgMyAwIFIKL0luZm8gMSAwIFIKL0lEIFsgPDM1NjEzMTMyNjIzNzY0MzczODM1NjEzNjY0MzUzNTM3MzUzNjM5NjI2MjM3MzA2NDMyMzQzMjMyNjEzNzMwMzk+IDwzNTYxMzEzMjYyMzc2NDM3MzgzNTYxMzY2NDM1MzUzNzM1MzYzOTYyNjIzNzMwNjQzMjM0MzIzMjYxMzczMDM5PiBdCi9FbmNyeXB0IDUgMCBSCj4+CnN0YXJ0eHJlZgo0NzYKJSVFT0YK', 'base64');
    const encrypted = fixture(new Uint8Array(encryptedBytes));
    await expect(encrypted.module.handlers.execute({ requestId: 'password-required', operation: 'execute', artifactRef: encrypted.ref }))
      .rejects.toMatchObject({ code: 'PASSWORD_REQUIRED' });
  });
});
