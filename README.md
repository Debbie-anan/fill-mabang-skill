# fill-mabang-skill

Claude Code skill：自动读取 Excel 订单数据并填写马帮ERP新建订单表单。

## 功能

- 读取桌面 `马帮信息.xlsx` 中的订单数据（支持多语言字段标签）
- 使用 Playwright 自动打开浏览器并导航到马帮ERP
- 在 iframe 中自动填写所有可填字段
- 支持持久化登录会话（首次登录后无需重复登录）

## 安装

### 1. 安装依赖

```bash
pip install openpyxl playwright
playwright install chromium
```

### 2. 安装 Claude Code Skill

将此仓库克隆到 Claude Code skills 目录：

```bash
git clone https://github.com/Debbie-anan/fill-mabang-skill.git ~/.claude/skills/fill-mabang
```

或手动将 `SKILL.md` 和 `scripts/` 目录复制到 `~/.claude/skills/fill-mabang/`。

### 3. 准备 Excel 文件

在桌面放置 `马帮信息.xlsx`，格式为 `标签：值`（支持中英文冒号），例如：

| A列 |
|-----|
| 收件人：John Smith |
| 电话：+1-555-1234 |
| 邮箱：john@example.com |
| Country：United States |
| 地址：123 Main St |
| 城市：New York |
| 省：NY |
| 邮编：10001 |
| 交易号：US+123-456-789 |
| 内部单号：1234567 |

## 使用

在 Claude Code 中输入：

```
/fill-mabang
```

或直接说"填马帮"。

## 配置（可选）

通过环境变量自定义路径：

| 环境变量 | 默认值 | 说明 |
|---------|--------|------|
| `MABANG_EXCEL_PATH` | `~/Desktop/马帮信息.xlsx` | Excel 文件路径 |
| `MABANG_SESSION_DIR` | `~/mabang_session` | 浏览器会话目录 |
| `MABANG_CDP_PORT` | `9222` | Chrome DevTools 端口 |

## 支持的字段

| 字段 | 支持的标签语言 |
|------|--------------|
| 收件人姓名 | 中/英/德/法/西/意/葡/荷/俄/波/捷/匈/土/日/韩/瑞典/丹麦/芬兰 |
| 电话 | 中/英/德/法/西/意/葡/荷/俄/波/芬兰/日/韩 |
| 邮箱 | 中/英/德/法/西/俄/日/韩 |
| 国家 | 中/英/意/法/西/葡/德/荷/波/土/捷/匈/俄/日/韩/芬兰 |
| 省/州 | 中/英/德/西/葡/法/俄/波/日 |
| 城市 | 中/英/德/法/西/意/葡/荷/俄/波/土/匈/捷/日/韩/瑞典 |
| 邮编 | 中/英/德/法/西/意/俄/波/土/瑞典/芬兰/荷/日/韩/葡 |
| 地址 | 中/英/德/法/西/意/葡/荷/俄/波/土/瑞典/芬兰/日/韩 |
| 交易号 | 中文 |
| 内部单号 | 中文 |
| 客户ID | 中文 |

## 注意事项

- 仅支持 Windows 系统
- "所属国家"下拉框为自定义组件，需手动选择
- 首次运行时需在弹出的浏览器中手动登录马帮ERP
- 订单号保持系统默认值，不会自动填写
