# Timeline Maker

Kivisense 内部项目 timeline 自动生成 skill。用于把简单事项清单转换成 Excel 甘特排期表，自动生成日期轴、责任方勾选、未完成状态、甘特色块、结束星标和 Kivisense logo。

## 安装

在本机终端执行：

```bash
git clone https://github.com/27ruien/timeline_skill.git ~/.codex/skills/timeline-maker
```

安装后重启 Codex。重启后可以用：

```text
$timeline-maker
```

来触发这个 skill。

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

每条事项至少给 4 个信息：

```text
事项名称, 责任方, 开始日期, 工作日天数
```

示例：

```text
1. UI Design, Kivisense, 2026-06-01, 5天
2. Asset Review, brand, 2026-06-03, 3天
3. Development & Integration, Kivisense, brand, 2026-06-08, 10天
```

## 责任方规则

```text
Kivisense          只在 Kivisense 列打勾
brand / Brands    只在 Brands 列打勾
Kivisense, brand  两列都打勾
```

如果没有写责任方，则两列都不打勾。

## 日期和状态规则

- `5天`、`5 days`、`5 workdays` 都会按 5 个工作日处理。
- 工作日默认跳过周六、周日。
- 开始日期如果落在周末，会顺延到下一个周一。
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

首次使用请先安装：
git clone https://github.com/27ruien/timeline_skill.git ~/.codex/skills/timeline-maker

安装后重启 Codex，然后这样使用：
用 $timeline-maker 做一个 timeline：
项目名：xxx
1. 事项名称, Kivisense, 2026-06-01, 5天
2. 事项名称, Kivisense, brand, 2026-06-08, 10天
3. 事项名称, brand, 2026-06-20, 3天

规则：
- Kivisense / brand 会自动在对应责任方列打勾
- 天数按工作日计算，会跳过周末
- Status 默认未完成，所以不用写
- 甘特图末尾会自动放星标
```

## 更新

如果已经安装过，后续更新执行：

```bash
cd ~/.codex/skills/timeline-maker
git pull
```

然后重启 Codex。

