#!/usr/bin/env python3
"""Local web app for generating timeline Excel files."""

from __future__ import annotations

import argparse
import json
import os
import re
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


def normalize_base_path(value: str) -> str:
    base = value.strip().rstrip("/")
    if base and not base.startswith("/"):
        base = "/" + base
    return base


BASE_PATH = normalize_base_path(os.environ.get("BASE_PATH", ""))


def with_base(path: str) -> str:
    if not path.startswith("/"):
        path = "/" + path
    if not BASE_PATH:
        return path
    return BASE_PATH + path


def strip_base_path(path: str) -> str:
    if BASE_PATH and path.startswith(BASE_PATH + "/"):
        stripped = path[len(BASE_PATH) :]
        return stripped or "/"
    return path


def render_index_html() -> str:
    return (
        INDEX_HTML
        .replace("__BASE_PATH_JSON__", json.dumps(BASE_PATH))
        .replace("__LOGO_SRC__", with_base("/assets/kivisense-logo.png"))
    )


INDEX_HTML = """<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>项目排期工具</title>
  <style>
    :root {
      --ink: #111827;
      --muted: #667085;
      --line: #e5e7eb;
      --soft: #f8fafc;
      --panel: #ffffff;
      --accent: #00b050;
      --accent-dark: #04733a;
      --danger: #b42318;
      --shadow: 0 1px 2px rgba(16, 24, 40, 0.05), 0 12px 28px rgba(16, 24, 40, 0.06);
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: "Google Sans Text", -apple-system, BlinkMacSystemFont, "Segoe UI", "Microsoft YaHei", sans-serif;
      color: var(--ink);
      background: #f6f8fb;
      padding-bottom: 124px;
    }
    .topbar {
      height: 92px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 18px;
      padding: 0 28px;
      background: rgba(255, 255, 255, 0.94);
      border-bottom: 1px solid var(--line);
      backdrop-filter: blur(12px);
    }
    .topbar-left {
      min-width: 0;
      display: flex;
      align-items: center;
      gap: 18px;
      flex: 0 0 auto;
    }
    .brand-logo {
      width: auto;
      height: 58px;
      object-fit: contain;
      display: block;
    }
    main {
      padding: 22px 28px 30px;
    }
    .example-panel {
      width: fit-content;
      max-width: 100%;
      margin-left: auto;
      display: flex;
      align-items: center;
      gap: 10px;
      background: rgba(255, 255, 255, 0.78);
      border: 1px solid #e7edf5;
      border-radius: 12px;
      padding: 7px;
      box-shadow: 0 8px 24px rgba(15, 23, 42, 0.05);
    }
    .example-title {
      flex: 0 0 auto;
      height: 28px;
      display: inline-flex;
      align-items: center;
      border-radius: 8px;
      background: #f3f7fb;
      padding: 0 9px;
      color: #53657c;
      font-size: 12px;
      font-weight: 750;
    }
    .example-list {
      min-width: 0;
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
    }
    .example-item {
      display: inline-flex;
      align-items: center;
      gap: 7px;
      max-width: 100%;
      height: 28px;
      padding: 0 9px;
      border: 1px solid #edf1f5;
      border-radius: 8px;
      background: #fbfcfe;
      color: #334155;
      font-size: 12px;
      white-space: nowrap;
    }
    .example-model {
      color: #04733a;
      font-weight: 750;
    }
    .example-date {
      color: var(--muted);
      font-variant-numeric: tabular-nums;
    }
    .project-panel,
    .task-panel {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 10px;
      box-shadow: var(--shadow);
    }
    .project-panel {
      width: min(860px, 100%);
      margin-bottom: 16px;
      padding: 16px 20px;
    }
    .task-panel {
      padding: 18px;
    }
    .toolbar {
      display: grid;
      grid-template-columns: minmax(240px, 490px) auto 1fr;
      align-items: end;
      gap: 18px;
    }
    label {
      display: block;
      font-size: 12px;
      font-weight: 700;
      color: var(--muted);
      margin: 0 0 7px;
    }
    input, select {
      width: 100%;
      height: 36px;
      border: 1px solid var(--line);
      border-radius: 8px;
      color: var(--ink);
      background: #ffffff;
      font: inherit;
      padding: 7px 9px;
      outline: none;
    }
    select {
      appearance: none;
      padding-right: 34px;
      background-image: url("data:image/svg+xml,%3Csvg width='14' height='14' viewBox='0 0 14 14' fill='none' xmlns='http://www.w3.org/2000/svg'%3E%3Cpath d='M3.5 5.25L7 8.75L10.5 5.25' stroke='%23334155' stroke-width='1.5' stroke-linecap='round' stroke-linejoin='round'/%3E%3C/svg%3E");
      background-position: right 13px center;
      background-repeat: no-repeat;
      background-size: 14px 14px;
    }
    input:focus, select:focus {
      border-color: var(--accent);
      box-shadow: 0 0 0 3px rgba(0, 176, 80, 0.12);
    }
    .custom-select {
      position: relative;
    }
    .select-trigger {
      width: 100%;
      height: 34px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 10px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: #fff;
      color: var(--ink);
      font: inherit;
      font-weight: 400;
      text-align: left;
      padding: 0 11px;
      cursor: pointer;
    }
    .select-trigger:focus {
      border-color: var(--accent);
      box-shadow: 0 0 0 3px rgba(0, 176, 80, 0.12);
      outline: none;
    }
    .select-menu {
      position: absolute;
      top: calc(100% + 6px);
      left: 0;
      right: 0;
      z-index: 35;
      display: none;
      padding: 5px;
      border: 1px solid #d9e2ec;
      border-radius: 10px;
      background: #fff;
      box-shadow: 0 14px 34px rgba(15, 23, 42, 0.12);
    }
    .custom-select.open .select-menu {
      display: block;
    }
    .select-option {
      width: 100%;
      height: 32px;
      display: flex;
      align-items: center;
      gap: 8px;
      border: 0;
      border-radius: 7px;
      background: #fff;
      color: #172033;
      font: inherit;
      font-weight: 400;
      text-align: left;
      padding: 0 9px;
      cursor: pointer;
    }
    .select-option:hover {
      background: #f6f8fb;
    }
    .select-option.selected {
      color: #0f172a;
      background: #fff;
      font-weight: 520;
    }
    .select-option.selected::before {
      content: "✓";
      width: 14px;
      color: var(--accent);
      font-weight: 760;
    }
    .select-option:not(.selected)::before {
      content: "";
      width: 14px;
    }
    .select-chevron {
      flex: 0 0 auto;
      position: relative;
      width: 14px;
      height: 14px;
      color: #475569;
      font-size: 0;
    }
    .select-chevron::after {
      content: "";
      position: absolute;
      left: 3px;
      top: 4px;
      width: 6px;
      height: 6px;
      border-right: 1.5px solid currentColor;
      border-bottom: 1.5px solid currentColor;
      transform: rotate(45deg);
    }
    input[type="checkbox"] {
      width: 16px;
      height: 16px;
      padding: 0;
      accent-color: var(--accent);
    }
    .checks {
      display: flex;
      gap: 14px;
      align-items: center;
      padding-bottom: 8px;
    }
    .check {
      display: inline-flex;
      align-items: center;
      gap: 7px;
      font-size: 13px;
      font-weight: 700;
      color: var(--ink);
    }
    .check input {
      border-radius: 4px;
    }
    .primary, .secondary, .danger {
      border: 1px solid transparent;
      border-radius: 8px;
      height: 34px;
      padding: 0 13px;
      font: inherit;
      font-weight: 750;
      cursor: pointer;
      white-space: nowrap;
    }
    .primary {
      height: 36px;
      padding: 0 16px;
      color: #fff;
      background: var(--accent);
      box-shadow: none;
    }
    .primary:hover { background: var(--accent-dark); }
    .secondary {
      color: var(--ink);
      background: #fff;
      border: 1px solid var(--line);
    }
    .secondary:hover {
      background: #f9fafb;
      border-color: #cbd5e1;
    }
    .danger {
      width: 30px;
      height: 30px;
      padding: 0;
      color: #98a2b3;
      background: transparent;
      border-color: transparent;
      font-size: 18px;
      line-height: 1;
    }
    .danger:hover {
      color: var(--danger);
      background: #fff1f1;
      border-color: #ffd5d5;
    }
    .danger:disabled {
      color: #d4dbe5;
      background: transparent;
      border-color: transparent;
      cursor: not-allowed;
    }
    .section-head {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      margin-bottom: 8px;
    }
    .section-head label { margin-bottom: 0; }
    .table-wrap {
      overflow-x: auto;
      border: 1px solid var(--line);
      border-radius: 10px;
      background: #fff;
      box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.8);
    }
    table {
      width: 100%;
      min-width: 1120px;
      border-collapse: collapse;
      table-layout: fixed;
    }
    th, td {
      border-bottom: 1px solid var(--line);
      padding: 8px;
      vertical-align: middle;
    }
    td { height: 50px; }
    th {
      height: 36px;
      background: #f9fbfd;
      color: var(--muted);
      font-size: 12px;
      text-align: left;
      font-weight: 800;
      position: sticky;
      top: 0;
      z-index: 1;
    }
    tbody tr.dragging { opacity: 0.45; }
    tbody tr:hover { background: #f8fbff; }
    td input, td select { height: 34px; }
    .col-drag { width: 44px; text-align: center; }
    .col-index { width: 46px; text-align: center; color: var(--muted); }
    .col-model { width: 140px; }
    .col-task { width: 290px; }
    .col-stakeholder { width: 190px; }
    .col-status { width: 116px; }
    .col-range { width: 270px; }
    .col-days { width: 116px; }
    .col-action { width: 58px; text-align: center; }
    .range-field {
      display: grid;
      grid-template-columns: minmax(0, 1fr) 18px minmax(0, 1fr);
      align-items: center;
      height: 34px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: #fff;
      overflow: hidden;
    }
    .range-field:focus-within {
      border-color: var(--accent);
      box-shadow: 0 0 0 3px rgba(0, 176, 80, 0.12);
    }
    .range-field input {
      height: 32px;
      border: 0;
      border-radius: 0;
      padding: 6px 8px;
      color: #1f2937;
      background: transparent;
      font-variant-numeric: tabular-nums;
      letter-spacing: 0;
    }
    .range-field input:focus {
      box-shadow: none;
    }
    .range-sep {
      color: var(--muted);
      font-size: 12px;
      text-align: center;
    }
    .drag-handle {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      width: 28px;
      height: 28px;
      border-radius: 5px;
      color: var(--muted);
      cursor: grab;
      user-select: none;
    }
    .drag-handle:active { cursor: grabbing; }
    .floating-actions {
      position: fixed;
      left: 50%;
      bottom: 24px;
      z-index: 20;
      display: flex;
      align-items: center;
      gap: 10px;
      max-width: calc(100vw - 48px);
      padding: 6px;
      border: 1px solid rgba(203, 213, 225, 0.72);
      border-radius: 14px;
      background: rgba(255, 255, 255, 0.62);
      box-shadow: 0 10px 28px rgba(15, 23, 42, 0.12);
      backdrop-filter: blur(20px);
      transform: translateX(-50%);
    }
    .floating-actions::before {
      content: "";
      position: absolute;
      inset: 0;
      border-radius: inherit;
      pointer-events: none;
      box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.86);
    }
    .floating-actions .primary,
    .floating-actions .secondary {
      position: relative;
      height: 32px;
      padding: 0 16px;
      border-radius: 10px;
      font-size: 13px;
      font-weight: 560;
    }
    .quick-actions {
      position: relative;
      display: flex;
      align-items: center;
      gap: 6px;
      max-width: min(760px, calc(100vw - 250px));
      overflow-x: auto;
      scrollbar-width: none;
    }
    .quick-actions::-webkit-scrollbar {
      display: none;
    }
    .floating-actions .primary {
      min-width: 118px;
      background: rgba(0, 176, 80, 0.86);
      box-shadow: 0 6px 16px rgba(0, 176, 80, 0.12);
    }
    .floating-actions .primary:hover {
      background: rgba(0, 146, 67, 0.9);
    }
    .floating-actions .secondary {
      border-color: rgba(203, 213, 225, 0.72);
      background: rgba(255, 255, 255, 0.54);
      color: #172033;
    }
    .floating-actions .secondary:hover {
      background: rgba(255, 255, 255, 0.78);
    }
    .toast {
      position: fixed;
      top: 22px;
      left: 50%;
      z-index: 40;
      display: inline-flex;
      align-items: center;
      gap: 8px;
      min-width: 154px;
      height: 38px;
      padding: 0 14px;
      border: 1px solid rgba(226, 232, 240, 0.86);
      border-radius: 999px;
      background: rgba(255, 255, 255, 0.82);
      box-shadow: 0 10px 24px rgba(15, 23, 42, 0.1);
      color: #0f172a;
      font-size: 13px;
      font-weight: 560;
      pointer-events: none;
      opacity: 0;
      transform: translate(-50%, -10px);
      transition: opacity 180ms ease, transform 180ms ease;
      backdrop-filter: blur(18px);
    }
    .toast.show {
      opacity: 1;
      transform: translate(-50%, 0);
    }
    .toast-icon {
      width: 20px;
      height: 20px;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      border-radius: 999px;
      color: #fff;
      background: var(--accent);
      font-size: 13px;
      font-weight: 760;
      line-height: 1;
      box-shadow: 0 6px 14px rgba(0, 176, 80, 0.12);
    }
    .add-icon {
      display: inline-block;
      margin-right: 6px;
      font-size: 14px;
      line-height: 0;
      transform: translateY(0);
    }
    .example {
      display: none;
      margin-top: auto;
      max-width: 680px;
      color: var(--muted);
      font-size: 12px;
      line-height: 1.65;
    }
    .example pre {
      margin: 8px 0 0;
      padding: 12px;
      overflow: auto;
      border: 1px solid var(--line);
      border-radius: 6px;
      background: #fbfcfe;
      color: #334155;
      font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace;
      font-size: 12px;
    }
    .status {
      min-height: 18px;
      font-size: 13px;
      color: var(--muted);
    }
    .status.error { color: var(--danger); }
    @media (max-width: 860px) {
      .topbar { padding: 0 16px; }
      main { padding: 16px; }
      .brand-logo { height: 44px; }
      .example-panel { align-items: flex-start; flex-direction: column; }
      .example-item { height: auto; min-height: 30px; white-space: normal; }
      .toolbar { grid-template-columns: 1fr; align-items: stretch; }
      .checks { padding-bottom: 0; }
      .floating-actions {
        left: 16px;
        right: 16px;
        justify-content: flex-start;
        transform: none;
      }
      .quick-actions { max-width: calc(100vw - 190px); }
    }
  </style>
</head>
<body>
  <div class="workspace">
    <header class="topbar">
      <div class="topbar-left">
        <img class="brand-logo" src="__LOGO_SRC__" alt="Kivisense">
      </div>
      <section class="example-panel">
        <div class="example-title">Example</div>
        <div class="example-list">
          <div class="example-item">
            <span class="example-model">需求</span>
            <span>Project requirement</span>
            <span>Kivisense</span>
            <span class="example-date">2026-06-01 - 2026-06-05</span>
            <span>5 workdays</span>
          </div>
          <div class="example-item">
            <span class="example-model">设计</span>
            <span>Creative Proposal</span>
            <span>Kivisense + brand</span>
            <span class="example-date">2026-06-08 - 2026-06-22</span>
            <span>10 workdays</span>
          </div>
        </div>
      </section>
    </header>
    <main>
      <section class="project-panel">
        <div class="toolbar">
          <div>
            <label for="projectName">Project title</label>
            <input id="projectName" maxlength="20" value="AR Campaign" autocomplete="off">
          </div>
          <div class="checks">
            <label class="check"><input id="includeModel" type="checkbox"> Model</label>
            <label class="check"><input id="includeStatus" type="checkbox" checked> Status</label>
          </div>
          <div class="status" id="status"></div>
        </div>
      </section>

      <section class="task-panel">
        <div>
          <div class="section-head">
            <label>Task list</label>
          </div>
          <div class="table-wrap">
            <table>
            <thead>
              <tr>
                <th class="col-drag"></th>
                <th class="col-index">#</th>
                <th class="col-model" data-model-col>Model</th>
                <th class="col-task">Task</th>
                <th class="col-stakeholder">Stakeholder</th>
                <th class="col-status" data-status-col>Status</th>
                <th class="col-range">Date range</th>
                <th class="col-days">Workdays</th>
                <th class="col-action"></th>
              </tr>
            </thead>
            <tbody id="taskRows"></tbody>
            </table>
          </div>
        </div>
      </section>
    </main>
    <div class="floating-actions">
      <div class="quick-actions" aria-label="Quick add tasks">
        <button class="secondary" data-quick-add="需求" type="button"><span class="add-icon">+</span>需求</button>
        <button class="secondary" data-quick-add="方案" type="button"><span class="add-icon">+</span>方案</button>
        <button class="secondary" data-quick-add="设计" type="button"><span class="add-icon">+</span>设计</button>
        <button class="secondary" data-quick-add="内容物料" type="button"><span class="add-icon">+</span>内容物料</button>
        <button class="secondary" data-quick-add="内容制作" type="button"><span class="add-icon">+</span>内容制作</button>
        <button class="secondary" data-quick-add="程序开发" type="button"><span class="add-icon">+</span>程序开发</button>
        <button class="secondary" data-quick-add="UAT" type="button"><span class="add-icon">+</span>UAT</button>
        <button class="secondary" data-quick-add="上线" type="button"><span class="add-icon">+</span>上线</button>
      </div>
      <button class="primary" id="generateButton" type="button">Generate</button>
    </div>
    <div class="toast" id="successToast" role="status" aria-live="polite">
      <span class="toast-icon">✓</span>
      <span class="toast-text">操作成功</span>
    </div>
  </div>
  <script>
    window.BASE_PATH = __BASE_PATH_JSON__;
    const taskRowsEl = document.getElementById("taskRows");
    const projectEl = document.getElementById("projectName");
    const includeModelEl = document.getElementById("includeModel");
    const includeStatusEl = document.getElementById("includeStatus");
    const statusEl = document.getElementById("status");
    const generateButton = document.getElementById("generateButton");
    const successToast = document.getElementById("successToast");
    let toastTimer = null;
    const chinaPublicHolidays = new Set([
      "2026-01-01", "2026-01-02", "2026-01-03",
      "2026-02-15", "2026-02-16", "2026-02-17", "2026-02-18", "2026-02-19", "2026-02-20", "2026-02-21", "2026-02-22", "2026-02-23",
      "2026-04-04", "2026-04-05", "2026-04-06",
      "2026-05-01", "2026-05-02", "2026-05-03", "2026-05-04", "2026-05-05",
      "2026-06-19", "2026-06-20", "2026-06-21",
      "2026-09-25", "2026-09-26", "2026-09-27",
      "2026-10-01", "2026-10-02", "2026-10-03", "2026-10-04", "2026-10-05", "2026-10-06", "2026-10-07"
    ]);
    const chinaAdjustedWorkdays = new Set([
      "2026-01-04",
      "2026-02-14", "2026-02-28",
      "2026-05-09",
      "2026-09-20", "2026-10-10"
    ]);

    const standardRows = [
      { model: "", name: "Project requirement", stakeholder: "Kivisense", start: "2026-06-01", workdays: 5 },
      { model: "", name: "Creative Proposal", stakeholder: "Both", start: "2026-06-08", workdays: 10 },
      { model: "", name: "Development & Integration", stakeholder: "Kivisense", start: "2026-06-15", workdays: 8 },
      { model: "", name: "Brand Asset Review", stakeholder: "Brands", start: "2026-06-18", workdays: 4 },
      { model: "", name: "Launch online", stakeholder: "Both", start: "2026-06-30", workdays: 1 }
    ];
    const modelRows = [
      { model: "需求", name: "Project requirement", stakeholder: "Kivisense", start: "2026-06-01", workdays: 5 },
      { model: "设计", name: "Creative Proposal", stakeholder: "Both", start: "2026-06-08", workdays: 10 },
      { model: "开发", name: "Development & Integration", stakeholder: "Kivisense", start: "2026-06-15", workdays: 8 },
      { model: "需求", name: "Scope addendum", stakeholder: "Brands", start: "2026-06-18", workdays: 4 },
      { model: "上线", name: "Launch online", stakeholder: "Both", start: "2026-06-30", workdays: 1 }
    ];
    const quickTaskTemplates = {
      "需求": ["需求梳理", "需求确认"],
      "内容物料": ["工业模型&文件", "ID图", "交互高保"],
      "内容制作": [
        "高视效3D模型制作*1",
        "高视效3D材质渲染",
        "颜色/纹理/材质*3",
        "渲染引擎",
        "亮点功能",
        "客户反馈",
        "反馈修改",
        "内容确认",
        "内容交付"
      ],
      "程序开发": ["前后端开发", "多端适配", "数据埋点", "测试报告"],
      "UAT": ["UAT测试 & 反馈", "UAT修改", "UAT确认"]
    };

    function appPath(path) {
      const base = window.BASE_PATH || "";
      const normalized = path.startsWith("/") ? path : `/${path}`;
      return `${base}${normalized}`;
    }

    function parseLocalDate(value) {
      if (!value) return null;
      const [year, month, day] = value.split("-").map(Number);
      return new Date(year, month - 1, day);
    }

    function formatLocalDate(date) {
      const year = date.getFullYear();
      const month = String(date.getMonth() + 1).padStart(2, "0");
      const day = String(date.getDate()).padStart(2, "0");
      return `${year}-${month}-${day}`;
    }

    function isWorkday(date) {
      const key = formatLocalDate(date);
      if (chinaAdjustedWorkdays.has(key)) return true;
      if (chinaPublicHolidays.has(key)) return false;
      const day = date.getDay();
      return day !== 0 && day !== 6;
    }

    function nextWorkday(date) {
      const next = new Date(date);
      while (!isWorkday(next)) next.setDate(next.getDate() + 1);
      return next;
    }

    function addWorkdays(startValue, workdays) {
      let current = nextWorkday(parseLocalDate(startValue));
      let remaining = Number(workdays) - 1;
      while (remaining > 0) {
        current.setDate(current.getDate() + 1);
        if (isWorkday(current)) remaining -= 1;
      }
      return formatLocalDate(current);
    }

    function countWorkdays(startValue, endValue) {
      let current = nextWorkday(parseLocalDate(startValue));
      const end = nextWorkday(parseLocalDate(endValue));
      let count = 0;
      while (current <= end) {
        if (isWorkday(current)) count += 1;
        current.setDate(current.getDate() + 1);
      }
      return count;
    }

    function stakeholderToList(value) {
      if (value === "Both") return ["Kivisense", "Brands"];
      if (value === "Kivisense") return ["Kivisense"];
      if (value === "Brands") return ["Brands"];
      return [];
    }

    function closeCustomSelects(except = null) {
      document.querySelectorAll("[data-custom-select]").forEach((select) => {
        if (select !== except) select.classList.remove("open");
      });
    }

    function syncCustomSelect(select) {
      const input = select.querySelector("input[type='hidden']");
      const trigger = select.querySelector("[data-select-trigger]");
      const options = [...select.querySelectorAll("[data-select-option]")];
      const selected = options.find((option) => option.dataset.value === input.value) || options[0];
      input.value = selected.dataset.value || "";
      trigger.querySelector("[data-select-label]").textContent = selected.dataset.label || selected.textContent.trim();
      options.forEach((option) => {
        option.classList.toggle("selected", option === selected);
      });
    }

    function wireCustomSelect(select) {
      const input = select.querySelector("input[type='hidden']");
      const trigger = select.querySelector("[data-select-trigger]");
      const options = [...select.querySelectorAll("[data-select-option]")];
      trigger.addEventListener("click", (event) => {
        event.stopPropagation();
        const willOpen = !select.classList.contains("open");
        closeCustomSelects(select);
        select.classList.toggle("open", willOpen);
      });
      options.forEach((option) => {
        option.addEventListener("click", (event) => {
          event.stopPropagation();
          input.value = option.dataset.value || "";
          syncCustomSelect(select);
          select.classList.remove("open");
        });
      });
      syncCustomSelect(select);
    }

    function showSuccessToast(message = "操作成功") {
      successToast.querySelector(".toast-text").textContent = message;
      successToast.classList.add("show");
      clearTimeout(toastTimer);
      toastTimer = setTimeout(() => {
        successToast.classList.remove("show");
      }, 2000);
    }

    function syncModelVisibility() {
      document.querySelectorAll("[data-model-col]").forEach((el) => {
        el.style.display = includeModelEl.checked ? "" : "none";
      });
    }

    function syncStatusVisibility() {
      document.querySelectorAll("[data-status-col]").forEach((el) => {
        el.style.display = includeStatusEl.checked ? "" : "none";
      });
    }

    function renumberRows() {
      [...taskRowsEl.children].forEach((row, index) => {
        row.querySelector(".col-index").textContent = index + 1;
        const deleteButton = row.querySelector("[data-action='delete']");
        if (deleteButton) deleteButton.disabled = index === 0;
      });
    }

    function wireDateLogic(row) {
      const start = row.querySelector('[data-field="start"]');
      const days = row.querySelector('[data-field="workdays"]');
      const end = row.querySelector('[data-field="end"]');

      const updateEnd = () => {
        if (start.value && Number(days.value) > 0) {
          end.value = addWorkdays(start.value, Number(days.value));
        }
      };
      const updateDays = () => {
        if (start.value && end.value) {
          const count = countWorkdays(start.value, end.value);
          if (count > 0) days.value = count;
        }
      };

      start.addEventListener("change", updateEnd);
      days.addEventListener("input", updateEnd);
      end.addEventListener("change", updateDays);
      updateEnd();
    }

    function addRow(task = {}) {
      const row = document.createElement("tr");
      row.draggable = true;
      row.innerHTML = `
        <td class="col-drag"><span class="drag-handle" title="Drag to reorder">⋮⋮</span></td>
        <td class="col-index"></td>
        <td class="col-model" data-model-col><input data-field="model" placeholder="需求" autocomplete="off"></td>
        <td class="col-task"><input data-field="name" placeholder="Task name" autocomplete="off"></td>
        <td class="col-stakeholder">
          <div class="custom-select" data-custom-select>
            <input data-field="stakeholder" type="hidden">
            <button class="select-trigger" data-select-trigger type="button">
              <span data-select-label>Unassigned</span>
              <span class="select-chevron">⌄</span>
            </button>
            <div class="select-menu">
              <button class="select-option" data-select-option data-value="" data-label="Unassigned" type="button">Unassigned</button>
              <button class="select-option" data-select-option data-value="Kivisense" data-label="Kivisense" type="button">Kivisense</button>
              <button class="select-option" data-select-option data-value="Brands" data-label="brand" type="button">brand</button>
              <button class="select-option" data-select-option data-value="Both" data-label="Kivisense + brand" type="button">Kivisense + brand</button>
            </div>
          </div>
        </td>
        <td class="col-status" data-status-col>
          <div class="custom-select" data-custom-select>
            <input data-field="status" type="hidden">
            <button class="select-trigger" data-select-trigger type="button">
              <span data-select-label>Incomplete</span>
              <span class="select-chevron">⌄</span>
            </button>
            <div class="select-menu">
              <button class="select-option" data-select-option data-value="incomplete" data-label="Incomplete" type="button">Incomplete</button>
              <button class="select-option" data-select-option data-value="done" data-label="Done" type="button">Done</button>
            </div>
          </div>
        </td>
        <td class="col-range">
          <span class="range-field">
            <input data-field="start" type="date" autocomplete="off" aria-label="Start date">
            <span class="range-sep">至</span>
            <input data-field="end" type="date" autocomplete="off" aria-label="End date">
          </span>
        </td>
        <td class="col-days"><input data-field="workdays" type="number" min="0" step="1" placeholder="0" autocomplete="off"></td>
        <td class="col-action"><button class="danger" data-action="delete" type="button" title="Delete" aria-label="Delete">×</button></td>
      `;
      taskRowsEl.appendChild(row);
      row.querySelector('[data-field="model"]').value = task.model || "";
      row.querySelector('[data-field="name"]').value = task.name || "";
      row.querySelector('[data-field="stakeholder"]').value = task.stakeholder || "";
      row.querySelector('[data-field="status"]').value = task.status || "incomplete";
      row.querySelector('[data-field="start"]').value = task.start || "";
      row.querySelector('[data-field="workdays"]').value = Object.prototype.hasOwnProperty.call(task, "workdays") ? task.workdays : 0;
      if (task.end) row.querySelector('[data-field="end"]').value = task.end;

      row.querySelectorAll("[data-custom-select]").forEach(wireCustomSelect);
      wireDateLogic(row);
      row.querySelector("[data-action='delete']").addEventListener("click", () => {
        if (taskRowsEl.children.length > 1 && row !== taskRowsEl.firstElementChild) {
          row.remove();
          renumberRows();
        }
      });
      row.addEventListener("dragstart", () => row.classList.add("dragging"));
      row.addEventListener("dragend", () => {
        row.classList.remove("dragging");
        renumberRows();
      });
      row.addEventListener("dragover", (event) => {
        event.preventDefault();
        const dragging = taskRowsEl.querySelector(".dragging");
        if (!dragging || dragging === row) return;
        const rect = row.getBoundingClientRect();
        const after = event.clientY > rect.top + rect.height / 2;
        taskRowsEl.insertBefore(dragging, after ? row.nextSibling : row);
      });
      syncModelVisibility();
      syncStatusVisibility();
      renumberRows();
    }

    function fillRows(rows) {
      taskRowsEl.innerHTML = "";
      rows.forEach(addRow);
    }

    function collectRows() {
      const rows = [...taskRowsEl.children].map((row, index) => {
        const model = row.querySelector('[data-field="model"]').value.trim();
        const name = row.querySelector('[data-field="name"]').value.trim();
        const stakeholder = row.querySelector('[data-field="stakeholder"]').value;
        const status = row.querySelector('[data-field="status"]').value;
        const start = row.querySelector('[data-field="start"]').value;
        const workdays = Number(row.querySelector('[data-field="workdays"]').value);
        const end = row.querySelector('[data-field="end"]').value;
        if (!model && !name && !stakeholder && !start && !workdays && !end) return null;
        if (includeModelEl.checked && !model) throw new Error(`第 ${index + 1} 行缺少 Model`);
        if (!name) throw new Error(`第 ${index + 1} 行缺少事项名称`);
        if (!start) throw new Error(`第 ${index + 1} 行缺少开始日期`);
        if (!end && (!Number.isInteger(workdays) || workdays < 1)) {
          throw new Error(`第 ${index + 1} 行需要结束日期或工作日天数`);
        }
        return {
          model,
          name,
          owners: stakeholderToList(stakeholder),
          status: includeStatusEl.checked ? status : "incomplete",
          start,
          end,
          workdays: Number.isInteger(workdays) && workdays > 0 ? workdays : undefined
        };
      }).filter(Boolean);
      if (!rows.length) throw new Error("请至少填写一条事项");
      return rows;
    }

    includeModelEl.addEventListener("change", () => {
      syncModelVisibility();
    });
    includeStatusEl.addEventListener("change", syncStatusVisibility);
    document.querySelectorAll("[data-quick-add]").forEach((button) => {
      button.addEventListener("click", () => {
        const model = button.dataset.quickAdd || "";
        const names = quickTaskTemplates[model] || [""];
        names.forEach((name) => {
          addRow({ model, name, workdays: 0 });
        });
      });
    });
    document.addEventListener("click", () => closeCustomSelects());

    async function generate() {
      generateButton.disabled = true;
      statusEl.textContent = "Generating...";
      statusEl.className = "status";
      try {
        const tasks = collectRows();
        const response = await fetch(appPath("/generate"), {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            project_name: projectEl.value,
            tasks,
            include_model: includeModelEl.checked,
            include_status: includeStatusEl.checked
          })
        });
        if (!response.ok) {
          const problem = await response.json().catch(() => ({ error: "Generate failed" }));
          throw new Error(problem.error || "Generate failed");
        }
        const blob = await response.blob();
        const url = URL.createObjectURL(blob);
        const safeName = (projectEl.value || "timeline").replace(/[\\\\/:*?"<>|\\s]+/g, "_");
        const a = document.createElement("a");
        a.href = url;
        a.download = `${safeName}_timeline.xlsx`;
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);
        statusEl.textContent = "";
        showSuccessToast("操作成功");
      } catch (error) {
        statusEl.textContent = error.message;
        statusEl.className = "status error";
      } finally {
        generateButton.disabled = false;
      }
    }

    generateButton.addEventListener("click", generate);
    fillRows(standardRows);
    syncModelVisibility();
    syncStatusVisibility();
    window.addEventListener("pageshow", () => {
      syncModelVisibility();
      syncStatusVisibility();
    });
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
        dates = DATE_RE.findall(line)
        duration_match = DURATION_RE.search(line)
        if not dates:
            raise ValueError(f"第 {line_number} 行缺少开始日期")

        start = dates[0]
        end = dates[1] if len(dates) > 1 else ""
        workdays = re.search(r"\d+", duration_match.group(0)).group(0) if duration_match else ""
        owners = []
        if re.search(r"\b(kivisense|kv)\b|我方", line, re.IGNORECASE):
            owners.append("Kivisense")
        if re.search(r"\b(brand|brands|client)\b|客户|品牌方", line, re.IGNORECASE):
            owners.append("Brands")

        first_date = DATE_RE.search(line)
        name_part = line[: first_date.start()]
        pieces = [piece.strip() for piece in re.split(r"[,，/、]+", name_part) if piece.strip()]
        filtered = [piece for piece in pieces if not OWNER_RE.fullmatch(piece)]
        if include_model:
            if len(filtered) < 2:
                raise ValueError(f"第 {line_number} 行开启 Model 后，需要写：Model, 事项名称, 相关方, 开始日期, 结束日期或工作日")
            model = filtered[0]
            name = filtered[1]
        else:
            model = ""
            name = filtered[0] if filtered else pieces[0] if pieces else ""
        if not name:
            raise ValueError(f"第 {line_number} 行缺少事项名称")

        task = {"model": model, "name": name, "owners": owners, "start": start}
        if end:
            task["end"] = end
        if workdays:
            task["workdays"] = int(workdays)
        tasks.append(task)

    if not tasks:
        raise ValueError("请至少输入一条事项")
    return tasks


class TimelineRequestHandler(BaseHTTPRequestHandler):
    server_version = "TimelineMaker/1.0"

    def do_GET(self) -> None:
        self.handle_get_like_request(write_body=True)

    def do_HEAD(self) -> None:
        self.handle_get_like_request(write_body=False)

    def handle_get_like_request(self, write_body: bool) -> None:
        parsed = urlparse(self.path)
        if BASE_PATH and parsed.path == BASE_PATH:
            location = with_base("/") + (f"?{parsed.query}" if parsed.query else "")
            self.respond_redirect(location)
            return

        path = strip_base_path(parsed.path)
        if path == "/":
            self.respond_text(render_index_html(), "text/html; charset=utf-8", write_body=write_body)
            return
        if path == "/assets/kivisense-logo.png":
            self.respond_file(Path(__file__).parent / "assets" / "kivisense-logo.png", "image/png", write_body=write_body)
            return
        self.send_error(404)

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        path = strip_base_path(parsed.path)
        if path != "/generate":
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

    def respond_text(self, text: str, content_type: str, write_body: bool = True) -> None:
        body = text.encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        if write_body:
            self.wfile.write(body)

    def respond_redirect(self, location: str, status: int = 301) -> None:
        self.send_response(status)
        self.send_header("Location", location)
        self.end_headers()

    def respond_json(self, payload: dict, status: int = 200) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def respond_file(self, path: Path, content_type: str, write_body: bool = True) -> None:
        if not path.exists():
            self.send_error(404)
            return
        body = path.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        if write_body:
            self.wfile.write(body)

    def log_message(self, format: str, *args) -> None:
        return


def run_self_test() -> Path:
    config = {
        "project_name": "Local Test",
        "include_model": True,
        "include_status": False,
        "tasks": [
            {"model": "需求", "name": "Project requirement", "owners": ["Kivisense"], "start": "2026-06-01", "workdays": 5},
            {"model": "设计", "name": "Creative Proposal", "owners": ["Kivisense", "Brands"], "start": "2026-06-08", "end": "2026-06-22"},
            {"model": "需求", "name": "Scope addendum", "owners": ["Brands"], "start": "2026-06-18", "workdays": 4},
        ],
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
    url = f"http://{HOST}:{args.port}{with_base('/')}"
    print(f"Timeline Maker is running: {url}")
    if BASE_PATH:
        print(f"BASE_PATH is set to: {BASE_PATH}")
    if not args.no_browser:
        webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
