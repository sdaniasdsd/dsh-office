# 第一批母版

`report`（通用报告）、`technical`（技术文档）、`chinese-long`（中文长文）是本模块原创的代码式样式母版，定义在 `src/engine/presets.ts`。没有直接打包网络上的公司标识、校徽或授权不清的 Word 文件，也不宣称这些母版来自第三方开源项目。

运行 `npm run fixtures:build`，会在 `fixtures/_generated/<唯一目录>/` 生成三份可重复填充的 DOCX 母版和三份填好的样例。母版槽位为 `{{organization}}`、`{{title}}`、`{{summary}}`。

也可把自己的 DOCX 母版交给 `fillTemplate`。目前只填充正文、表格、页眉和页脚中普通段落的单行文字，支持槽位被 Word 拆成多个文字片段；保留首片段的样式。循环行、条件段落、替换图片、跨段落槽位、内容控件绑定和字段内部槽位不支持，不能把此能力当成完整模板语言。

字体未嵌入：报告/技术文档使用 Calibri、Microsoft YaHei，中文长文使用 Times New Roman、SimSun。接收端字体不齐全时需要由部署方提供合法字体或调整样式母版。最终版面必须由 `docx-render` 渲染后检查。
