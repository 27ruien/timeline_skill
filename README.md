# Timeline Workbench v17

Kivisense 项目排期工作台。v17 版本提供一个可本地运行的网页工具，用于把项目任务、负责人、开始/结束日期快速生成 Kivisense 风格 Excel timeline / Gantt 排期表。

这个仓库只保留 v17 工具本体、生成脚本、样式资产和输入说明。

## v17 功能

- 本地 Web 工作台，默认地址 `http://127.0.0.1:8765`
- 中文 / English 双语界面
- 项目标题、阶段、任务表格、负责人、开始日期、结束日期、工作日自动计算
- 阶段是核心字段，主编辑区按真实阶段动态生成 Tabs，并可拖动调整阶段顺序
- 可选展示字段：`状态`
- 内置 AR、3D、数字化项目排期模板
- 支持按阶段快捷添加任务
- 支持拖拽排序、上移、下移、复制、删除任务
- 支持排期预览抽屉，包含甘特图视图和表格视图
- 支持导入标准导出的 `.xlsx` 排期文件继续编辑
- 支持导出 Kivisense 风格 Excel，包含 logo、负责人勾选、状态列、日期轴、甘特色块和结束星标
- 支持 `BASE_PATH` 子路径部署，例如 `/tool/timeline`
- 支持命令行用 JSON 直接生成 `.xlsx`

## 文件结构

```text
.
├── README.md
├── SKILL.md
├── local_app.py
├── requirements.txt
├── start.command
├── assets/
│   ├── gantt-end-star.png
│   └── kivisense-logo.png
├── references/
│   └── input-schema.md
└── scripts/
    └── build_timeline.py
```

## 快速开始

克隆仓库：

```bash
git clone https://github.com/27ruien/gridtimeline.git
cd gridtimeline
```

macOS 可以直接双击：

```text
start.command
```

首次启动会自动创建 `.venv` 并安装依赖。启动后会自动打开：

```text
http://127.0.0.1:8765
```

也可以用终端启动：

```bash
python3 -m venv .venv
. .venv/bin/activate
python3 -m pip install -r requirements.txt
python3 local_app.py
```

后台运行、不自动打开浏览器：

```bash
python3 local_app.py --no-browser
```

指定端口：

```bash
python3 local_app.py --port 8888
```

## 网页使用

1. 填写项目标题。
2. 选择界面语言：`中文` 或 `English`。
3. 使用阶段 Tabs 切换当前阶段；可以在表格上方重命名当前阶段。
4. 选择是否展示 `状态` 字段。
5. 使用内置模板，或点击 `+ 新增任务` / 阶段快捷按钮添加任务。
6. 为每行填写任务、负责人、开始日期和结束日期。
7. 打开 `排期预览` 检查甘特图和表格。
8. 点击 `导出 Excel` 下载排期文件。

负责人选项：

```text
Kivisense
Brand
Brand & Kivisense
```

导出后会自动映射为 Excel 里的 `Kivisense` / `Brands` 勾选列。

## Excel 输出规则

- 工作表名称为 `Timeline`
- 顶部包含 Kivisense logo 和项目标题
- 支持中文或英文表头，跟随网页语言导出
- 阶段始终导出为 `Model` / `工作内容` 列，并按阶段合并单元格
- 阶段、任务、负责人、状态等单元格内容保留页面当前值，不做中英文反向翻译
- `include_model` 仅作为旧输入兼容字段保留，当前导出总是包含阶段列
- 开启 `状态` 后会导出状态列
- 日期轴按中国工作日生成
- 任务色块覆盖开始日期到结束日期之间的工作日
- 每个任务最后一个甘特单元格带结束星标
- 默认状态为未完成；只有显式选择完成才会在状态列标记完成

## 导入排期

点击 `导入排期`，选择由本工具导出的 `.xlsx` 文件或包含彩色任务条的视觉甘特图。系统会尝试读取：

- 项目标题
- 阶段 / Model
- 任务名称
- Kivisense / Brands 负责人勾选
- 状态
- 甘特日期范围
- 任务条的背景填充色（即使单元格没有值）及不连续分段

如果源文件只包含月份和日数字、没有明确年份，可以在导入按钮旁填写年份；否则系统会使用文件名或工作簿创建时间作为候选年份，并在写入前显示预览要求确认。日期以 `YYYY-MM-DD` 字符串保存，时间轴中缺失的周末或节假日不会被自动补齐。

导入后可以继续在网页中编辑，并再次导出 Excel。

## 命令行生成

除了网页，也可以使用脚本从 JSON 生成 Excel：

```bash
python3 scripts/build_timeline.py input.json output.xlsx
```

示例 `input.json`：

```json
{
  "project_name": "AR Campaign",
  "include_model": true,
  "include_status": true,
  "language": "zh",
  "tasks": [
    {
      "model": "需求",
      "name": "需求梳理与范围确认",
      "owners": ["Kivisense", "Brands"],
      "start": "2026-06-10",
      "end": "2026-06-12",
      "status": "incomplete"
    },
    {
      "model": "内容制作",
      "name": "高视效 3D 模型制作",
      "owners": ["Kivisense"],
      "start": "2026-06-15",
      "workdays": 5,
      "status": "done"
    }
  ]
}
```

更多输入规则见：

```text
references/input-schema.md
```

## 子路径部署

如果需要部署到 Nginx 子路径，例如：

```text
https://example.com/tool/timeline/
```

启动时设置 `BASE_PATH`：

```bash
BASE_PATH=/tool/timeline python3 local_app.py --no-browser
```

Nginx 示例：

```nginx
location = /tool/timeline {
    return 301 /tool/timeline/;
}

location /tool/timeline/ {
    proxy_pass http://127.0.0.1:8765/;
    proxy_http_version 1.1;
    proxy_set_header Host $host;
    proxy_set_header X-Real-IP $remote_addr;
    proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
    proxy_set_header X-Forwarded-Proto $scheme;
}
```

本地验证：

```bash
BASE_PATH=/tool/timeline python3 local_app.py --no-browser
curl -I http://127.0.0.1:8765/tool/timeline/
curl -I http://127.0.0.1:8765/tool/timeline/assets/kivisense-logo.png
```

## 依赖

```text
openpyxl>=3.1.0
pillow>=10.0.0
```

`start.command` 会在首次启动时自动安装依赖；手动运行时请先执行 `python3 -m pip install -r requirements.txt`。

## 自检

运行内置自检：

```bash
python3 local_app.py --self-test
```

自检会在临时目录生成一个测试 Excel，用来确认 Excel 生成链路可用。
