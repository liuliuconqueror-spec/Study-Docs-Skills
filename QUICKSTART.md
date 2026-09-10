# Quickstart

下面的步骤适用于 Windows PowerShell。所有占位符都需要替换成你自己的路径；不要把令牌粘贴到仓库文件中。

## 1. 准备 Python

需要 Python 3.10 或更高版本：

```powershell
python --version
python -m pip install pypdf
```

`pypdf` 用于 PDF 页数预检、超限拆分判断和质量模式 PPTX 导出验证。若要使用完全本地的 PDF `private` 模式，再安装：

```powershell
python -m pip install pymupdf4llm
```

## 2. 安装这两个 Skill

在仓库根目录运行：

```powershell
.\install.ps1
```

默认安装位置为 `$env:CODEX_HOME\skills`（如果 `CODEX_HOME` 已设置），否则为 `$HOME\.agents\skills`。使用自定义目录时，后续命令沿用同一个值：

```powershell
$skillsRoot = '<skills-root>'
.\install.ps1 -SkillsRoot $skillsRoot
```

未自定义时，可这样得到相同的默认值：

```powershell
$skillsRoot = if ($env:CODEX_HOME) {
  Join-Path $env:CODEX_HOME 'skills'
} else {
  Join-Path $HOME '.agents\skills'
}
```

## 3. 单独安装 MinerU Skill

本仓库不捆绑 MinerU。当前测试基线是 `Nebutra/MinerU-Skill` 的 `v3.3.1`。可使用其维护者提供的安装方式安装当前版本：

```powershell
npx skills add Nebutra/MinerU-Skill
```

需要固定到已测试版本时，把仓库克隆成与本项目同级、目录名为 `mineru` 的 Skill：

```powershell
git clone --depth 1 --branch v3.3.1 https://github.com/Nebutra/MinerU-Skill.git `
  (Join-Path $skillsRoot 'mineru')
```

摄取器需要以下两个文件存在：

```text
<skills-root>/mineru/scripts/mineru.py
<skills-root>/mineru/scripts/chunking.py
```

如果 MinerU 安装在别处，可为当前 PowerShell 会话指定：

```powershell
$env:STUDY_DOCS_MINERU_SCRIPT = '<path-to-mineru.py>'
$env:STUDY_DOCS_MINERU_CHUNKER = '<path-to-chunking.py>'
```

不要把可用令牌写进任何 JSON 或脚本。确实需要 Standard API 时，只在环境中设置 `MINERU_TOKEN`，并按 MinerU 官方方式管理它。

## 4. 运行环境检查

```powershell
$ingest = Join-Path $skillsRoot 'study-document-ingest\scripts\ingest.py'
$reader = Join-Path $skillsRoot 'study-source-reader\scripts\search.py'
python $ingest --doctor --json
```

环境检查会报告 Python、MinerU Skill、chunker、`pypdf`、令牌是否被检测到、PPTX 标准化后端、状态目录和可选 vault 的状态，但不会打印令牌值。

## 5. 可选：配置 Obsidian

不配置 vault 也能建立缓存和目录。需要默认发布到 Obsidian 时，先初始化状态目录，再复制示例配置：

```powershell
$stateDir = if ($env:STUDY_DOCS_STATE_DIR) {
  $env:STUDY_DOCS_STATE_DIR
} else {
  Join-Path $HOME '.codex\study-docs'
}

python $ingest --doctor --json
Copy-Item (Join-Path $skillsRoot 'study-document-ingest\config.example.json') `
  (Join-Path $stateDir 'config.json') -Force
notepad (Join-Path $stateDir 'config.json')
```

配置文件只支持：

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

把 `vault_path` 改为已存在且由你明确选择的 vault，或继续保留 `null`。也可以不改配置，在单次摄取命令中使用 `--vault '<vault-path>'`。

## 6. 首次摄取

先用只读计划确认分类和目标：

```powershell
python $ingest '<path-to-course-file>' `
  --course '<course-name>' `
  --title '<document-title>' `
  --dry-run --json
```

确认后运行质量模式；不需要 Obsidian 发布时加 `--no-publish`：

```powershell
python $ingest '<path-to-course-file>' `
  --course '<course-name>' `
  --title '<document-title>' `
  --mode quality --no-publish --json
```

`quality` 会把文件提交给 MinerU 云端。机密 PDF 应改用本地模式：

```powershell
python $ingest '<path-to-file.pdf>' `
  --course '<course-name>' `
  --mode private --no-publish --json
```

不要对 Office 或图片使用 `private`；该模式会在调用任何云端前拒绝这些格式。

## 7. 检索已摄取资料

```powershell
python $reader '<question-or-keyword>' `
  --course '<course-name>' `
  --title '<document-title>' `
  --json
```

返回 `AMBIGUOUS` 时补充课程或标题；返回 `NOT_FOUND` 时，先确认资料是否已成功摄取。检索器不会调用 MinerU，也不会把失败结果当作缓存。
