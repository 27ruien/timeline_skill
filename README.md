# Timeline Maker

Kivisense 内部项目 timeline 自动生成工具。可以作为 Codex skill 使用，也可以在本地打开一个网页工具使用。用于把简单事项清单转换成 Excel 甘特排期表，自动生成日期轴、责任方勾选、未完成状态、甘特色块、结束星标和 Kivisense logo。

## 方式一：本地网页使用

适合不使用 Codex 的同事。

### 下载

在本机终端执行：

```bash
git clone https://github.com/27ruien/timeline_skill.git
cd timeline_skill
```

或直接在 GitHub 页面点击 `Code` -> `Download ZIP`，下载后解压。

### 启动

macOS 可以双击：

```text
start.command
```

第一次双击时，它会自动创建本地运行环境并安装依赖。这个过程可能需要几十秒。

也可以在终端执行：

```bash
python3 local_app.py
```

如果用终端启动，需要先安装依赖：

```bash
python3 -m pip install -r requirements.txt
```

启动后会自动打开本地网页：

```text
http://127.0.0.1:8765
```

在网页里填写项目标题、选择是否需要 `Model` / `Status`，然后在表格里逐行填写事项，点击底部悬浮操作条里的 `生成`。

页面支持：

- 点击底部悬浮操作条里的 `+ 新增` 增加一行，不需要手写序号。
- 每个字段都有固定输入位置，不需要手写逗号。
- 相关方用下拉框选择。
- 勾选 `Status` 后，事项清单会显示独立状态列，默认 `未完成`。
- 拖拽每行左侧手柄，可以调整事项顺序。
- 在 `日期范围` 里选择开始日期和结束日期，系统会反算工作日天数。
- 也可以先填工作日天数，系统会自动展示结束日期。
- 生成甘特图时，以日期范围的开始日期和结束日期为准。
- 工作日按中国放假调休规则计算：排除 2026 年中国法定节假日，并计入调休补班日。

### Model / Status 勾选项

- 勾选 `Model`：输出表会增加 `Model` 列，并按 Model 自动归组、合并单元格。
- 不勾选 `Model`：输出表从 `Description` 开始。
- 勾选 `Status`：页面事项清单和输出表都包含 `Status` 列。
- 不勾选 `Status`：输出表不包含 `Status` 列。
- 同一个 Model 会按第一次出现的 Model 顺序归组；同一个 Model 内的事项会按开始日期排序。

如果勾选了 `Model`，每行格式变为：

```text
Model, 事项名称, 相关方, 开始日期, 工作日天数或结束日期
```

示例：

```text
需求, Project requirement, Kivisense, 2026-06-01, 5天
设计, Creative Proposal, Kivisense, brand, 2026-06-08, 2026-06-22
需求, Scope addendum, brand, 2026-06-18, 4天
```

输出时，第 1 行和第 14 行都会被归到 `需求` 下面，`Model` 单元格会合并；`Description` 仍然保留两行不同事项。

## 方式二：Codex Skill 使用

适合已经在使用 Codex 的同事。

在本机终端执行：

```bash
git clone https://github.com/27ruien/timeline_skill.git ~/.codex/skills/timeline-maker
```

安装后重启 Codex。重启后可以用：

```text
$timeline-maker
```

来触发这个 skill。

## 部署到 Nginx 子路径

服务支持通过 `BASE_PATH` 部署到子路径，例如外部访问：

```text
https://gridworks.cn/tool/timeline
https://gridworks.cn/tool/timeline/
```

启动服务时设置：

```bash
BASE_PATH=/tool/timeline python3 local_app.py --no-browser
```

本地开发不设置 `BASE_PATH`，仍然访问：

```text
http://127.0.0.1:8765/
```

### systemd 示例

将路径按服务器实际目录调整：

```ini
[Unit]
Description=Timeline Maker
After=network.target

[Service]
Type=simple
User=www-data
Group=www-data
WorkingDirectory=/opt/timeline_skill
Environment=BASE_PATH=/tool/timeline
ExecStart=/usr/bin/python3 /opt/timeline_skill/local_app.py --no-browser
Restart=always
RestartSec=3

[Install]
WantedBy=multi-user.target
```

启用：

```bash
sudo systemctl daemon-reload
sudo systemctl enable --now timeline-maker
sudo systemctl status timeline-maker
```

### Nginx 示例

推荐保留尾斜杠的 `proxy_pass`，让 Nginx 把 `/tool/timeline/` 转发到后端 `/`，同时应用仍会在页面里生成带 `/tool/timeline` 的资源和接口路径。

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

如果你的 Nginx 使用不带尾斜杠的 `proxy_pass http://127.0.0.1:8765;`，后端也能识别 `/tool/timeline/...` 并剥掉前缀后处理。

### 测试

本地测试：

```bash
python3 local_app.py --no-browser
curl -I http://127.0.0.1:8765/
curl -I http://127.0.0.1:8765/assets/kivisense-logo.png
```

子路径测试：

```bash
BASE_PATH=/tool/timeline python3 local_app.py --no-browser
curl -I http://127.0.0.1:8765/tool/timeline
curl -I http://127.0.0.1:8765/tool/timeline/
curl -I http://127.0.0.1:8765/tool/timeline/assets/kivisense-logo.png
```

浏览器 Network 面板里，图片和生成接口应请求 `/tool/timeline/assets/...`、`/tool/timeline/generate`，不应请求根路径 `/assets` 或 `/generate`。

## 基础用法

直接在 Codex 里输入：

```text
用 $timeline-maker 做一个 timeline：
项目名：AR Campaign
1. Project requirement, Kivisense, 2026-06-01, 5天
2. Creative Proposal, Kivisense, brand, 2026-06-08, 10天
3. Launch online, Kivisense, brand, 2026-06-22, 1天
```

Codex 会输出一个 `.xlsx` 文件。

## 输入格式

表格中每条事项至少填写：

```text
事项名称, 相关方, 开始日期, 工作日天数或结束日期
```

示例：

```text
UI Design, Kivisense, 2026-06-01, 5天
Asset Review, brand, 2026-06-03, 2026-06-05
Development & Integration, Kivisense, brand, 2026-06-08, 10天
```

如果开启 `Model`，每条事项给 5 个信息：

```text
Model, 事项名称, 相关方, 开始日期, 工作日天数或结束日期
```

## 相关方规则

```text
Kivisense          只在 Kivisense 列打勾
brand / Brands    只在 Brands 列打勾
Kivisense, brand  两列都打勾
```

如果没有写责任方，则两列都不打勾。

## 日期和状态规则

- `5天`、`5 days`、`5 workdays` 都会按 5 个工作日处理。
- 工作日默认采用中国工作日：跳过周六、周日和 2026 年中国法定节假日，并计入调休补班日。
- 开始日期如果不是工作日，会顺延到下一个中国工作日。
- `Status` 默认未完成，所以默认留空。
- 只有明确写 `完成`、`已完成`、`done`、`complete` 或 `√`，才会在 `Status` 列打勾。

## 样式规则

- A 列是 `Description`，没有隐藏前置列。
- B/C/D 是 `Kivisense`、`Brands`、`Status`。
- E 列开始是工作日甘特图。
- 顶部包含 Kivisense logo 和项目标题。
- 每条甘特图末尾自动放一个星标。
- 甘特色块接近正方形。
- 禁止黄色色块。
- 任意相邻 4 行内，甘特色块颜色不重复。

## 推荐内部话术

可以直接复制下面这段发给团队：

```text
我们现在有一个 timeline 自动生成工具，可以把事项清单直接生成 Kivisense 风格 Excel 甘特排期表。

不用 Codex 的同事：
1. 打开 https://github.com/27ruien/timeline_skill
2. 下载仓库或执行 git clone
3. 进入 timeline_skill 文件夹
4. 双击 start.command
5. 第一次启动会自动安装依赖，等浏览器打开
6. 在打开的本地网页里选择是否需要 Model / Status
7. 点击底部 `+ 新增`，逐行填写内容；需要调整顺序时拖拽行
8. 点击底部 `生成`

使用 Codex 的同事：
git clone https://github.com/27ruien/timeline_skill.git ~/.codex/skills/timeline-maker

安装后重启 Codex，然后这样使用：
用 $timeline-maker 做一个 timeline：
项目名：xxx
1. 事项名称, Kivisense, 2026-06-01, 5天
2. 事项名称, Kivisense, brand, 2026-06-08, 10天
3. 事项名称, brand, 2026-06-20, 3天

规则：
- Kivisense / brand 会自动在对应相关方列打勾
- 可以在同一个日期范围组件里选择开始日期和结束日期，也可以填工作日天数自动算结束日期
- 甘特图按日期范围生成，会跳过中国法定节假日和非调休周末
- Status 默认未完成，所以不用写
- 甘特图末尾会自动放星标
- 如果勾选 Model，每行第一个字段写工作内容/模块，例如：需求, 事项名称, Kivisense, 2026-06-01, 5天
```

## 更新

如果是本地网页版本：

```bash
cd timeline_skill
git pull
```

如果是 Codex skill 版本：

```bash
cd ~/.codex/skills/timeline-maker
git pull
```

然后重启 Codex。
