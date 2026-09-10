# Study Docs Skills

Two Codex skills for durable course-document ingestion and source-grounded retrieval.

这是一套面向大学课程资料的 Codex Skill：先把 PDF、Office 文档和图片解析成经过验证的 Markdown、资源文件与检索分块，再从持久缓存中回答问题。重复使用同一份资料时，不会无故重新解析。

## 包含内容

- `study-document-ingest`：计算内容哈希，调用外部 MinerU Skill，验证 Markdown、图片和检索分块，提交内容寻址缓存与目录，并可安全发布人类可读副本到 Obsidian。
- `study-source-reader`：只读取已摄取的目录、分块、Markdown 和保存的视觉证据；它不会调用 MinerU，也不会自动重新摄取。

两者共用默认状态目录 `~/.codex/study-docs`。可用环境变量 `STUDY_DOCS_STATE_DIR` 改到其他位置。运行状态、缓存、日志、课程原文件、标准化 PDF 和 Obsidian 笔记均不在本仓库中。

## 重要边界

- `quality` 是默认模式，会通过外部 MinerU Skill 使用云端解析。不要把机密或受监管文件交给该模式。
- `fast` 可能在本地能力不可用时转向云端。
- `private` 只接受 PDF，要求安装 `pymupdf4llm`，并禁止云端回退。
- Obsidian 发布只写入 `<vault>/课程资料/<Course>/<Document>/source.md` 和 `assets/`；`<vault>/笔记/` 始终保留给用户笔记。
- `MINERU_TOKEN` 只从进程环境继承；不要把它写入配置、命令参数或仓库文件。无令牌时，外部 MinerU Skill 可使用其免令牌 Agent API；令牌用于较大的 Standard API 任务。
- 验证失败不会被报告为缓存命中、目录更新或发布成功。

## 安装

要求 Windows PowerShell、Git 和 Python 3.10 或更高版本。质量模式的 PPTX 还需要本机 Microsoft PowerPoint，或可执行的 LibreOffice headless 后端。

```powershell
git clone https://github.com/liuliuconqueror-spec/Study-Docs-Skills.git
Set-Location .\Study-Docs-Skills
.\install.ps1
```

安装脚本优先使用 `$env:CODEX_HOME\skills`（若已设置 `CODEX_HOME`），否则使用当前 OpenAI 文档中的个人 Skill 目录 `$HOME\.agents\skills`。也可显式指定：

```powershell
.\install.ps1 -SkillsRoot '<skills-root>'
```

Codex 通常会自动发现变化；若新 Skill 没有出现，请重启 Codex。目录约定参见 [OpenAI 官方 Build skills 文档](https://developers.openai.com/codex/build-skills)。

完整依赖、MinerU 安装、配置和首次运行步骤见 [QUICKSTART.md](QUICKSTART.md)。

## 配置结构

实现支持的配置字段只有以下 7 个；不要添加未实现的键：

```json
{
  "schema_version": 1,
  "vault_path": null,
  "default_mode": "quality",
  "default_course_fallback": "_Inbox",
  "source_retention": "reference",
  "publish_enabled": true,
  "publication_root": "课程资料"
}
```

- `vault_path`：已存在且由用户明确选择的 Obsidian vault，或 `null`。
- `default_mode`：`quality`、`fast` 或 `private`。
- `default_course_fallback`：无法可靠推断课程时使用的安全分类。
- `source_retention`：`reference` 仅记录路径和哈希；`copy` 还按内容哈希保存一份原文件。
- `publish_enabled`：选择 vault 后是否允许发布。
- `publication_root`：必须是 `课程资料`；`笔记` 和其他根名称会被拒绝。

示例文件位于 [`study-document-ingest/config.example.json`](study-document-ingest/config.example.json)，字段细节见 [`config-schema.md`](study-document-ingest/references/config-schema.md)。首次非只读运行会在状态目录缺失时创建默认配置和空目录。

## MinerU 依赖与打包决定

本仓库不复制或捆绑 MinerU 代码。摄取器依赖单独安装的 [Nebutra/MinerU-Skill](https://github.com/Nebutra/MinerU-Skill)，并使用其中的 `scripts/mineru.py` 与 `scripts/chunking.py`；当前离线测试基线为 `v3.3.1`。脚本会依次检查同级 Skill 目录、`$HOME/.agents/skills`、`$CODEX_HOME/skills` 和旧版 `$HOME/.codex/skills`，也支持用 `STUDY_DOCS_MINERU_SCRIPT` 与 `STUDY_DOCS_MINERU_CHUNKER` 显式指定。

依赖保持外置有三个原因：

1. MinerU-Skill 是独立维护的第三方包装器，当前标注为 MIT；直接依赖便于更新和保留其完整许可证。
2. 上游 [OpenDataLab MinerU](https://github.com/opendatalab/MinerU) 使用基于 Apache 2.0、带商业门槛与在线服务标识义务的 [MinerU Open Source License](https://github.com/opendatalab/MinerU/blob/master/LICENSE.md)。使用者应按自己的场景核对当前条款。
3. 本项目不需要复制解析引擎、模型权重或云端客户端代码即可工作。

## 目录

```text
Study-Docs-Skills/
├── README.md
├── QUICKSTART.md
├── install.ps1
├── uninstall.ps1
├── .gitignore
├── study-document-ingest/
│   ├── SKILL.md
│   ├── config.example.json
│   ├── references/
│   ├── scripts/
│   └── tests/
└── study-source-reader/
    ├── SKILL.md
    ├── references/
    ├── scripts/
    └── tests/
```

## 开发与测试

离线单元测试不会上传课程文件：

```powershell
python -m unittest discover -s .\study-document-ingest\tests -p 'test_*.py'
python -m unittest discover -s .\study-source-reader\tests -p 'test_*.py'
```

真实云端解析应只用可公开的测试资料，并在测试前确认当前 MinerU 服务条款和数据处理要求。

## 项目许可证

本仓库当前没有项目级 `LICENSE` 文件。公开可见不等于自动授予修改或再分发权；如果希望接受外部贡献或允许正式再分发，应由仓库所有者另行选择并添加许可证。外部 MinerU 依赖继续遵循各自许可证。

## 卸载

```powershell
.\uninstall.ps1
```

卸载脚本只删除这两个已安装 Skill。它故意保留 `~/.codex/study-docs` 中的目录、缓存和配置，避免误删学习资料。

