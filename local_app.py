#!/usr/bin/env python3
"""Local web app for generating timeline Excel files."""

from __future__ import annotations

import argparse
import json
import re
import sys
import tempfile
import webbrowser
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from io import BytesIO
from pathlib import Path
from urllib.parse import urlparse

from scripts.build_timeline import build_workbook


HOST = "127.0.0.1"
DEFAULT_PORT = 8765


INDEX_HTML = """<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Timeline Maker</title>
  <style>
    :root {
      --ink: #1f2933;
      --muted: #64748b;
      --line: #cbd5e1;
      --soft: #f8fafc;
      --accent: #00b050;
      --accent-dark: #04733a;
      --danger: #b42318;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Microsoft YaHei", sans-serif;
      color: var(--ink);
      background: #ffffff;
    }
    header {
      display: grid;
      grid-template-columns: minmax(180px, 280px) 1fr auto;
      align-items: center;
      gap: 28px;
      padding: 28px 38px 18px;
      border-bottom: 1px solid var(--line);
    }
    header img { width: 220px; max-width: 100%; }
    h1 {
      margin: 0;
      font-size: 28px;
      line-height: 1.2;
      font-weight: 700;
      letter-spacing: 0;
    }
    main {
      display: grid;
      grid-template-columns: 340px minmax(0, 1fr);
      min-height: calc(100vh - 104px);
    }
    aside {
      border-right: 1px solid var(--line);
      padding: 24px 28px;
      background: var(--soft);
    }
    section {
      padding: 24px 32px 32px;
    }
    label {
      display: block;
      font-size: 13px;
      font-weight: 650;
      margin: 0 0 8px;
    }
    input, select, textarea {
      width: 100%;
      border: 1px solid var(--line);
      border-radius: 6px;
      color: var(--ink);
      font: inherit;
      outline: none;
      background: #fff;
    }
    input, select {
      height: 40px;
      padding: 8px 10px;
      margin-bottom: 18px;
    }
    input[type="checkbox"] {
      width: 16px;
      height: 16px;
      margin: 0;
      accent-color: var(--accent);
    }
    textarea {
      min-height: 440px;
      resize: vertical;
      padding: 12px;
      line-height: 1.55;
      tab-size: 2;
    }
    .task-table-wrap {
      overflow-x: auto;
      border: 1px solid var(--line);
      border-radius: 6px;
    }
    table {
      width: 100%;
      min-width: 920px;
      border-collapse: collapse;
      background: #fff;
    }
    th, td {
      border-bottom: 1px solid var(--line);
      padding: 8px;
      vertical-align: middle;
    }
    th {
      font-size: 12px;
      text-align: left;
      background: var(--soft);
      color: var(--muted);
      font-weight: 700;
    }
    td input, td select {
      margin-bottom: 0;
      height: 34px;
      border-radius: 4px;
    }
    .col-index { width: 44px; text-align: center; color: var(--muted); }
    .col-model { width: 130px; }
    .col-task { min-width: 260px; }
    .col-owner { width: 190px; }
    .col-date { width: 150px; }
    .col-days { width: 110px; }
    .col-action { width: 72px; text-align: center; }
    .icon-button {
      width: 34px;
      height: 34px;
      padding: 0;
      border-radius: 6px;
      background: #fee2e2;
      color: var(--danger);
      font-weight: 800;
    }
    .text-mode {
      margin-top: 18px;
      display: none;
    }
    .text-mode.open {
      display: block;
    }
    .text-mode textarea {
      min-height: 180px;
    }
    .hint {
      margin: 0 0 18px;
      color: var(--muted);
      font-size: 13px;
      line-height: 1.55;
    }
    .options {
      display: grid;
      gap: 10px;
      margin: 0 0 18px;
    }
    .option {
      display: flex;
      align-items: center;
      gap: 8px;
      font-size: 13px;
      font-weight: 650;
    }
    .actions {
      display: flex;
      align-items: center;
      gap: 12px;
      margin-top: 16px;
    }
    button {
      border: 0;
      border-radius: 6px;
      padding: 10px 16px;
      font: inherit;
      font-weight: 700;
      cursor: pointer;
    }
    .primary {
      background: var(--accent);
      color: white;
    }
    .primary:hover { background: var(--accent-dark); }
    .secondary {
      color: var(--ink);
      background: #e2e8f0;
    }
    .status {
      min-height: 20px;
      color: var(--muted);
      font-size: 13px;
    }
    .status.error { color: var(--danger); }
    code {
      font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace;
      font-size: 12px;
    }
    @media (max-width: 820px) {
      header { grid-template-columns: 1fr; }
      main { grid-template-columns: 1fr; }
      aside { border-right: 0; border-bottom: 1px solid var(--line); }
      textarea { min-height: 360px; }
    }
  </style>
</head>
<body>
  <header>
    <img src="/assets/kivisense-logo.png" alt="Kivisense">
    <h1>Timeline Maker</h1>
    <div class="status" id="status"></div>
  </header>
  <main>
    <aside>
      <label for="projectName">项目标题</label>
      <input id="projectName" maxlength="20" value="AR Campaign">
      <div class="options">
        <label class="option"><input id="includeModel" type="checkbox"> Model</label>
        <label class="option"><input id="includeStatus" type="checkbox" checked> Status</label>
      </div>
      <p class="hint" id="formatHint">每行格式：<br><code>事项名称, 责任方, 开始日期, 工作日天数</code></p>
      <p class="hint">责任方可写 <code>Kivisense</code>、<code>brand</code>，或两者都写。</p>
      <button class="secondary" id="addTaskButton" type="button">新增事项</button>
      <br><br>
      <button class="secondary" id="sampleButton" type="button">填入示例</button>
    </aside>
    <section>
      <label>事项清单</label>
      <div class="task-table-wrap">
        <table>
          <thead>
            <tr>
              <th class="col-index">#</th>
              <th class="col-model" data-model-col>Model</th>
              <th class="col-task">事项名称</th>
              <th class="col-owner">责任方</th>
              <th class="col-date">开始日期</th>
              <th class="col-days">工作日天数</th>
              <th class="col-action"></th>
            </tr>
          </thead>
          <tbody id="taskRows"></tbody>
        </table>
      </div>
      <div class="actions">
        <button class="secondary" id="toggleTextButton" type="button">粘贴文本</button>
      </div>
      <div class="text-mode" id="textMode">
        <label for="tasks">粘贴事项清单</label>
        <textarea id="tasks" spellcheck="false"></textarea>
      </div>
      <div class="actions">
        <button class="primary" id="generateButton" type="button">生成 Excel</button>
        <div class="status" id="inlineStatus"></div>
      </div>
    </section>
  </main>
  <script>
    const standardSample = `1. Project requirement, Kivisense, 2026-06-01, 5天
2. Creative Proposal, Kivisense, brand, 2026-06-08, 10天
3. Development & Integration, Kivisense, 2026-06-15, 8天
4. Brand Asset Review, brand, 2026-06-18, 4天
5. Launch online, Kivisense, brand, 2026-06-30, 1天`;
    const modelSample = `1. 需求, Project requirement, Kivisense, 2026-06-01, 5天
2. 设计, Creative Proposal, Kivisense, brand, 2026-06-08, 10天
3. 开发, Development & Integration, Kivisense, 2026-06-15, 8天
4. 需求, Scope addendum, brand, 2026-06-18, 4天
5. 上线, Launch online, Kivisense, brand, 2026-06-30, 1天`;

    const statusEl = document.getElementById("status");
    const inlineStatusEl = document.getElementById("inlineStatus");
    const tasksEl = document.getElementById("tasks");
    const taskRowsEl = document.getElementById("taskRows");
    const projectEl = document.getElementById("projectName");
    const includeModelEl = document.getElementById("includeModel");
    const includeStatusEl = document.getElementById("includeStatus");
    const formatHintEl = document.getElementById("formatHint");
    const textModeEl = document.getElementById("textMode");
    const toggleTextButton = document.getElementById("toggleTextButton");
    const button = document.getElementById("generateButton");

    const standardRows = [
      { model: "", name: "Project requirement", owner: "Kivisense", start: "2026-06-01", workdays: 5 },
      { model: "", name: "Creative Proposal", owner: "Both", start: "2026-06-08", workdays: 10 },
      { model: "", name: "Development & Integration", owner: "Kivisense", start: "2026-06-15", workdays: 8 },
      { model: "", name: "Brand Asset Review", owner: "Brands", start: "2026-06-18", workdays: 4 },
      { model: "", name: "Launch online", owner: "Both", start: "2026-06-30", workdays: 1 }
    ];
    const modelRows = [
      { model: "需求", name: "Project requirement", owner: "Kivisense", start: "2026-06-01", workdays: 5 },
      { model: "设计", name: "Creative Proposal", owner: "Both", start: "2026-06-08", workdays: 10 },
      { model: "开发", name: "Development & Integration", owner: "Kivisense", start: "2026-06-15", workdays: 8 },
      { model: "需求", name: "Scope addendum", owner: "Brands", start: "2026-06-18", workdays: 4 },
      { model: "上线", name: "Launch online", owner: "Both", start: "2026-06-30", workdays: 1 }
    ];

    function currentSample() {
      return includeModelEl.checked ? modelSample : standardSample;
    }

    function updateFormatHint() {
      formatHintEl.innerHTML = includeModelEl.checked
        ? '每行格式：<br><code>Model, 事项名称, 责任方, 开始日期, 工作日天数</code>'
        : '每行格式：<br><code>事项名称, 责任方, 开始日期, 工作日天数</code>';
      document.querySelectorAll("[data-model-col]").forEach((el) => {
        el.style.display = includeModelEl.checked ? "" : "none";
      });
    }

    tasksEl.value = currentSample();
    updateFormatHint();

    function ownerToList(owner) {
      if (owner === "Both") return ["Kivisense", "Brands"];
      if (owner === "Kivisense") return ["Kivisense"];
      if (owner === "Brands") return ["Brands"];
      return [];
    }

    function addRow(task = {}) {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td class="col-index"></td>
        <td class="col-model" data-model-col><input data-field="model" placeholder="需求"></td>
        <td class="col-task"><input data-field="name" placeholder="事项名称"></td>
        <td class="col-owner">
          <select data-field="owner">
            <option value="">未指定</option>
            <option value="Kivisense">Kivisense</option>
            <option value="Brands">Brands</option>
            <option value="Both">Kivisense + brand</option>
          </select>
        </td>
        <td class="col-date"><input data-field="start" type="date"></td>
        <td class="col-days"><input data-field="workdays" type="number" min="1" step="1" placeholder="5"></td>
        <td class="col-action"><button class="icon-button" type="button" title="删除">×</button></td>
      `;
      taskRowsEl.appendChild(tr);
      tr.querySelector('[data-field="model"]').value = task.model || "";
      tr.querySelector('[data-field="name"]').value = task.name || "";
      tr.querySelector('[data-field="owner"]').value = task.owner || "";
      tr.querySelector('[data-field="start"]').value = task.start || "";
      tr.querySelector('[data-field="workdays"]').value = task.workdays || "";
      tr.querySelector("button").addEventListener("click", () => {
        if (taskRowsEl.children.length > 1) {
          tr.remove();
          renumberRows();
        }
      });
      updateFormatHint();
      renumberRows();
    }

    function renumberRows() {
      [...taskRowsEl.children].forEach((row, index) => {
        row.querySelector(".col-index").textContent = index + 1;
      });
    }

    function fillRows(rows) {
      taskRowsEl.innerHTML = "";
      rows.forEach(addRow);
    }

    function collectRows() {
      const rows = [...taskRowsEl.children].map((row, index) => {
        const model = row.querySelector('[data-field="model"]').value.trim();
        const name = row.querySelector('[data-field="name"]').value.trim();
        const owner = row.querySelector('[data-field="owner"]').value;
        const start = row.querySelector('[data-field="start"]').value;
        const workdays = Number(row.querySelector('[data-field="workdays"]').value);
        if (!name && !start && !workdays && !model) return null;
        if (!name) throw new Error(`第 ${index + 1} 行缺少事项名称`);
        if (includeModelEl.checked && !model) throw new Error(`第 ${index + 1} 行缺少 Model`);
        if (!start) throw new Error(`第 ${index + 1} 行缺少开始日期`);
        if (!Number.isInteger(workdays) || workdays < 1) throw new Error(`第 ${index + 1} 行工作日天数无效`);
        return { model, name, owners: ownerToList(owner), start, workdays };
      }).filter(Boolean);
      if (!rows.length) throw new Error("请至少填写一条事项");
      return rows;
    }

    fillRows(standardRows);

    document.getElementById("sampleButton").addEventListener("click", () => {
      tasksEl.value = currentSample();
      fillRows(includeModelEl.checked ? modelRows : standardRows);
      inlineStatusEl.textContent = "";
      statusEl.textContent = "";
      statusEl.className = "status";
    });
    includeModelEl.addEventListener("change", () => {
      updateFormatHint();
    });
    document.getElementById("addTaskButton").addEventListener("click", () => addRow());
    toggleTextButton.addEventListener("click", () => {
      textModeEl.classList.toggle("open");
      toggleTextButton.textContent = textModeEl.classList.contains("open") ? "收起文本" : "粘贴文本";
    });

    async function generate() {
      button.disabled = true;
      inlineStatusEl.textContent = "生成中...";
      statusEl.textContent = "";
      statusEl.className = "status";
      try {
        const tableTasks = collectRows();
        const response = await fetch("/generate", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            project_name: projectEl.value,
            raw_tasks: tasksEl.value,
            tasks: tableTasks,
            include_model: includeModelEl.checked,
            include_status: includeStatusEl.checked
          })
        });
        if (!response.ok) {
          const problem = await response.json().catch(() => ({ error: "生成失败" }));
          throw new Error(problem.error || "生成失败");
        }
        const blob = await response.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        const safeName = (projectEl.value || "timeline").replace(/[\\\\/:*?"<>|\\s]+/g, "_");
        a.href = url;
        a.download = `${safeName}_timeline.xlsx`;
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);
        inlineStatusEl.textContent = "已生成";
      } catch (error) {
        inlineStatusEl.textContent = "";
        statusEl.textContent = error.message;
        statusEl.className = "status error";
      } finally {
        button.disabled = false;
      }
    }

    button.addEventListener("click", generate);
  </script>
</body>
</html>
"""


DATE_RE = re.compile(r"\b\d{4}[-/]\d{1,2}[-/]\d{1,2}\b")
DURATION_RE = re.compile(r"\b\d+\s*(?:天|day|days|workday|workdays|个工作日)\b", re.IGNORECASE)
OWNER_RE = re.compile(r"kivisense|kv|brand|brands|client|我方|客户|品牌方", re.IGNORECASE)


def parse_raw_tasks(raw_text: str, include_model: bool = False) -> list[dict]:
    tasks: list[dict] = []
    for line_number, raw_line in enumerate(raw_text.splitlines(), start=1):
        line = raw_line.strip()
        if not line:
            continue
        line = re.sub(r"^\s*\d+[\.\)、)]\s*", "", line)
        date_match = DATE_RE.search(line)
        duration_match = DURATION_RE.search(line)
        if not date_match or not duration_match:
            raise ValueError(f"第 {line_number} 行缺少开始日期或工作日天数")

        start = date_match.group(0)
        workdays = re.search(r"\d+", duration_match.group(0)).group(0)
        owners = []
        if re.search(r"\b(kivisense|kv)\b|我方", line, re.IGNORECASE):
            owners.append("Kivisense")
        if re.search(r"\b(brand|brands|client)\b|客户|品牌方", line, re.IGNORECASE):
            owners.append("Brands")

        name_part = line[: min(date_match.start(), duration_match.start())]
        pieces = [piece.strip() for piece in re.split(r"[,，/、]+", name_part) if piece.strip()]
        filtered = [piece for piece in pieces if not OWNER_RE.fullmatch(piece)]
        if include_model:
            if len(filtered) < 2:
                raise ValueError(f"第 {line_number} 行开启 Model 后，需要写：Model, 事项名称, 责任方, 开始日期, 工作日天数")
            model = filtered[0]
            name = filtered[1]
        else:
            model = ""
            name = filtered[0] if filtered else pieces[0] if pieces else ""
        if not name:
            raise ValueError(f"第 {line_number} 行缺少事项名称")

        tasks.append({"model": model, "name": name, "owners": owners, "start": start, "workdays": int(workdays)})

    if not tasks:
        raise ValueError("请至少输入一条事项")
    return tasks


class TimelineRequestHandler(BaseHTTPRequestHandler):
    server_version = "TimelineMaker/1.0"

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/":
            self.respond_text(INDEX_HTML, "text/html; charset=utf-8")
            return
        if parsed.path == "/assets/kivisense-logo.png":
            self.respond_file(Path(__file__).parent / "assets" / "kivisense-logo.png", "image/png")
            return
        self.send_error(404)

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path != "/generate":
            self.send_error(404)
            return
        try:
            length = int(self.headers.get("content-length", "0"))
            payload = json.loads(self.rfile.read(length).decode("utf-8"))
            include_model = bool(payload.get("include_model", False))
            include_status = bool(payload.get("include_status", True))
            structured_tasks = payload.get("tasks")
            tasks = structured_tasks if structured_tasks else parse_raw_tasks(payload.get("raw_tasks", ""), include_model=include_model)
            config = {
                "project_name": payload.get("project_name") or "Timeline",
                "tasks": tasks,
                "include_model": include_model,
                "include_status": include_status,
            }
            workbook = build_workbook(config)
            output = BytesIO()
            workbook.save(output)
            body = output.getvalue()
            filename = f"timeline_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        except Exception as exc:
            self.respond_json({"error": str(exc)}, status=400)

    def respond_text(self, text: str, content_type: str) -> None:
        body = text.encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def respond_json(self, payload: dict, status: int = 200) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def respond_file(self, path: Path, content_type: str) -> None:
        if not path.exists():
            self.send_error(404)
            return
        body = path.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, format: str, *args) -> None:
        return


def run_self_test() -> Path:
    config = {
        "project_name": "Local Test",
        "include_model": True,
        "include_status": False,
        "tasks": parse_raw_tasks(
            """1. 需求, Project requirement, Kivisense, 2026-06-01, 5天
2. 设计, Creative Proposal, Kivisense, brand, 2026-06-08, 10天
3. 需求, Scope addendum, brand, 2026-06-18, 4天""",
            include_model=True,
        ),
    }
    workbook = build_workbook(config)
    output_path = Path(tempfile.gettempdir()) / "timeline_maker_self_test.xlsx"
    workbook.save(output_path)
    return output_path


def main() -> int:
    parser = argparse.ArgumentParser(description="Run Timeline Maker locally.")
    parser.add_argument("--port", type=int, default=DEFAULT_PORT)
    parser.add_argument("--no-browser", action="store_true")
    parser.add_argument("--self-test", action="store_true")
    args = parser.parse_args()

    if args.self_test:
        output_path = run_self_test()
        print(output_path)
        return 0

    server = ThreadingHTTPServer((HOST, args.port), TimelineRequestHandler)
    url = f"http://{HOST}:{args.port}"
    print(f"Timeline Maker is running: {url}")
    if not args.no_browser:
        webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
