# DSH Office 运行时 0.9.0（Windows x64）

`dsh-toolchain` 接收的第一份工具链交接包：**打包好的运行时**，与 `@deepseek-ai/dsh-docx@0.9.0` 配对。

## 资产

| 资产 | 字节 | sha256 |
| --- | --- | --- |
| `deepseek-ai-dsh-docx-runtime-0.9.0.tgz` | 572,864,213 | `5e128b92c9e8a391dddf2e519d63f4c1497b1c7b5538ce095798e8a186df2b36` |

（这个文件与 `sdaniasdsd/dsh-office` 的 v0.9.0 release 里那份是**同一份字节**，digest 相同，装哪边都一样。
文件名沿用了 npm 包名 `@deepseek-ai/dsh-docx-runtime`。）

## 里面是什么

解包后 `<包根>/runtime/win32-x64/`：

| 组件 | 版本（二进制自报） | 体积 | 文件数 |
| --- | --- | --- | --- |
| Python（embeddable） | `Python 3.13.15` | 50,013,265 B | 665 |
| LibreOffice | `26.8.0.3 bce0998afefdbc355585ca324285661a2170ba77` | 1,577,413,569 B | 19,456 |
| Poppler | `pdftoppm version 26.09.0` | 126,544,679 B | 561 |

Python 侧带了 `lxml 6.1.0`、`python-pptx 1.0.2`、`Pillow 12.3.0`、`XlsxWriter 3.2.9`、`typing_extensions 4.16.0`
——也就是 DSH office 插件真正 import 的那几个。**不含** `oletools` / `docling` / `rdocx`（按需另装，缺了插件会明确报不可用）。

包根还有 `runtime.json`（schema `dsh-office-runtime/v1`），插件靠它定位；`README.md` 与 `package.json` 里声明了 `os: win32`、`cpu: x64`。

## 验收回执（本机对这份资产实跑）

```
✓ binary:python              Python 3.13.15
✓ binary:libreoffice         LibreOffice 26.8.0.3 bce0998afefdbc355585ca324285661a2170ba77
✓ binary:poppler             pdftoppm version 26.09.0
✓ python:lxml                6.1.0
✓ python:pptx                1.0.2
✓ python:PIL                 12.3.0
✓ python:xlsxwriter          3.2.9
✓ python:typing_extensions   ok
✓ manifest:schema            dsh-office-runtime/v1
✓ manifest:platform          win32-x64
✓ manifest:components        三个组件都标 present
结论：全部通过（退出码 0）
```

命令：`node scripts/verify-runtime.mjs <解包后的 runtime/win32-x64>`（本仓库里就有这份脚本）。

## 怎么用

1. **给 DSH office 插件用**：把 tgz 与核心包一起装进 profile（两个 `file:` 依赖），插件会自己按兄弟包找到它；
   实测 pnpm 隔离布局下也成立（`runtime source = runtime-package`、无缺失项）。
2. **复用宿主已有的 LibreOffice/Python**：不用装这个包，设 `DSH_OFFICE_RUNTIME_ROOT` 或插件配置 `runtimeRoot` 指过去即可。
3. **只要二进制**：解包后把 `python/`、`libreoffice/`、`poppler/` 任意搬走都行，路径由上面两种方式指定。

## 契约与后续

- 接口只有三样东西：`python/python.exe`、`libreoffice/program/soffice.com`、`poppler/poppler-26.09.0/Library/bin/pdftoppm.exe`，
  外加可选的 `runtime.json`。详见仓库里的 `docs/CONTRACT.md`。
- 想自己重建这份运行时（换版本、换来源）：`pwsh -File scripts/fetch-runtime.ps1` → `node scripts/verify-runtime.mjs runtime/win32-x64`。
- 换 poppler 版本要注意：目录名 `poppler-26.09.0` 目前被插件侧的默认路径写死，需要同时改插件侧或显式给 `pdftoppmPath`。
