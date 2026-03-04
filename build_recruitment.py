#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
採用管理ダッシュボード — タブベースUI
CSVをパースし、4タブ構成のHTMLダッシュボードを生成する。
"""
import csv
import json
import os
from pathlib import Path

BASE = Path(__file__).parent

# ---------------------------------------------------------------------------
# ステータス分類
# ---------------------------------------------------------------------------
HIDDEN_STATUSES = {'不採用', '離脱', '採用基準✖', ''}
HIRED_STATUSES = {'仮採用', '採用', '契約中'}
# それ以外は ACTION

# ACTION内の優先度グループ
ACTION_HIGH = {'面談予定', '面談日程調節中', '外部連絡申請中'}
ACTION_MEDIUM = {'トライアル実施中', '課題提出済み', '課題確認済み', '採用基準〇'}
ACTION_MONITORING = {'ヒアリング中', '返信保留中'}

# 全ステータス色マップ
STATUS_COLORS = {
    '採用': '#16a34a',
    '仮採用': '#22c55e',
    '契約中': '#15803d',
    '面談予定': '#2563eb',
    '面談日程調節中': '#3b82f6',
    '外部連絡申請中': '#ca8a04',
    'トライアル実施中': '#8b5cf6',
    '課題提出済み': '#a855f7',
    '課題確認済み': '#7c3aed',
    '採用基準〇': '#059669',
    'ヒアリング中': '#0891b2',
    '返信保留中': '#ea580c',
    '不採用': '#94a3b8',
    '離脱': '#cbd5e1',
    '採用基準✖': '#9ca3af',
    '別案件流し': '#6b7280',
    '不採用→再連絡中': '#d97706',
    '返信なし': '#9ca3af',
    '未対応': '#ef4444',
    '未設定': '#d1d5db',
}

DEFAULT_COLOR = '#64748b'


def _esc(t):
    if not t:
        return ''
    return str(t).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')


def _attr_esc(t):
    if not t:
        return ''
    return str(t).replace('&', '&amp;').replace('"', '&quot;').replace('<', '&lt;').replace('>', '&gt;')


# ---------------------------------------------------------------------------
# Step 1: load_csv — 空行除外
# ---------------------------------------------------------------------------
def load_csv():
    csv_path = BASE / '採用管理.csv'
    if not csv_path.exists():
        print(f"Warning: {csv_path} not found")
        return []

    with open(csv_path, 'r', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        rows = []
        for row in reader:
            name = (row.get('お名前') or '').strip()
            status = (row.get('現状のステータス') or '').strip()
            role = (row.get('役割') or '').strip()
            # 名前・ステータス・役割がすべて空の行を除外
            if not name and not status and not role:
                continue
            rows.append({
                'name': name,
                'link': (row.get('Crowdworks返信リンク') or '').strip(),
                'role': role,
                'expectation': (row.get('期待度') or '').strip(),
                'status': status,
                'note': (row.get('備考') or '').strip(),
                'interview_date': (row.get('面談日') or '').strip(),
            })
    return rows


# ---------------------------------------------------------------------------
# Step 2: aggregate — 3カテゴリ分類
# ---------------------------------------------------------------------------
def aggregate(rows):
    total = len(rows)
    status_counts = {}
    role_counts = {}
    interview_list = []

    hidden_rows = []
    hired_rows = []
    action_rows = {'high': [], 'medium': [], 'monitoring': [], 'other': []}
    action_count = 0

    for r in rows:
        st = r['status'] or '未設定'
        status_counts[st] = status_counts.get(st, 0) + 1

        role = r['role'] or '未設定'
        if role not in role_counts:
            role_counts[role] = {'total': 0, 'statuses': {}}
        role_counts[role]['total'] += 1
        role_counts[role]['statuses'][st] = role_counts[role]['statuses'].get(st, 0) + 1

        if st == '面談予定':
            interview_list.append(r)

        # カテゴリ分類
        if st in HIDDEN_STATUSES or r['status'] == '':
            hidden_rows.append(r)
        elif st in HIRED_STATUSES:
            hired_rows.append(r)
        else:
            action_count += 1
            if st in ACTION_HIGH:
                action_rows['high'].append(r)
            elif st in ACTION_MEDIUM:
                action_rows['medium'].append(r)
            elif st in ACTION_MONITORING:
                action_rows['monitoring'].append(r)
            else:
                action_rows['other'].append(r)

    # アクティブ応募者 = 全体 - hidden
    active_count = total - len(hidden_rows)

    return {
        'total': total,
        'active_count': active_count,
        'status_counts': status_counts,
        'role_counts': role_counts,
        'interview_list': interview_list,
        'hidden_rows': hidden_rows,
        'hired_rows': hired_rows,
        'action_rows': action_rows,
        'action_count': action_count,
        'hired_count': len(hired_rows),
        'mendan_yotei': status_counts.get('面談予定', 0),
        'saiyo': status_counts.get('採用', 0),
        'kari_saiyo': status_counts.get('仮採用', 0),
        'keiyaku': status_counts.get('契約中', 0),
    }


# ---------------------------------------------------------------------------
# Step 3-5: render_html — タブベースUI + CSS + JS
# ---------------------------------------------------------------------------
def render_html(rows, agg):
    # --- タブ1: 概況 ---
    tab1 = _render_tab_overview(rows, agg)
    # --- タブ2: 対応が必要 ---
    tab2 = _render_tab_action(agg)
    # --- タブ3: 採用済み ---
    tab3 = _render_tab_hired(agg)
    # --- タブ4: 全応募者検索 (JSON埋め込み) ---
    tab4_html, search_json = _render_tab_search(rows)

    action_badge = f'<span class="tab-badge">{agg["action_count"]}</span>' if agg['action_count'] > 0 else ''

    return f'''<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>採用管理ダッシュボード</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
{_css()}
</style>
</head>
<body>
<div class="container">
  <header class="header">
    <h1>採用管理ダッシュボード</h1>
    <p>応募者のステータス・役割別状況を一覧表示</p>
  </header>

  <nav class="tab-bar">
    <button class="tab-btn active" data-tab="overview">概況</button>
    <button class="tab-btn" data-tab="action">対応が必要{action_badge}</button>
    <button class="tab-btn" data-tab="hired">採用済み</button>
    <button class="tab-btn" data-tab="search">全応募者検索</button>
  </nav>

  <div class="tab-content active" data-tab="overview">
    {tab1}
  </div>
  <div class="tab-content" data-tab="action">
    {tab2}
  </div>
  <div class="tab-content" data-tab="hired">
    {tab3}
  </div>
  <div class="tab-content" data-tab="search">
    {tab4_html}
  </div>
</div>

<script>
window.SEARCH_DATA = {search_json};
{_js()}
</script>
</body>
</html>'''


# ===== タブ1: 概況 =====
def _render_tab_overview(rows, agg):
    # サマリーカード
    summary = f'''
    <div class="summary-cards">
      <div class="summary-card">
        <div class="summary-num">{agg['active_count']}</div>
        <div class="summary-label">アクティブ応募者</div>
      </div>
      <div class="summary-card card-green">
        <div class="summary-num">{agg['hired_count']}</div>
        <div class="summary-label">採用済み</div>
      </div>
      <div class="summary-card card-coral">
        <div class="summary-num">{agg['action_count']}</div>
        <div class="summary-label">対応が必要</div>
      </div>
      <div class="summary-card card-blue">
        <div class="summary-num">{agg['mendan_yotei']}</div>
        <div class="summary-label">面談予定</div>
      </div>
    </div>'''

    # ステータス別パイプライン
    # アクティブステータスのみバー表示、HIDDEN はグレーテキスト
    active_statuses = {}
    hidden_statuses = {}
    for st, cnt in agg['status_counts'].items():
        if st in HIDDEN_STATUSES or st == '未設定':
            hidden_statuses[st] = cnt
        else:
            active_statuses[st] = cnt

    pipeline_bars = []
    if active_statuses:
        max_count = max(active_statuses.values())
        # 優先順で表示
        ordered = []
        for st in list(ACTION_HIGH) + list(ACTION_MEDIUM) + list(ACTION_MONITORING) + list(HIRED_STATUSES):
            if st in active_statuses:
                ordered.append((st, active_statuses[st]))
        # 残り
        for st, cnt in sorted(active_statuses.items(), key=lambda x: -x[1]):
            if st not in dict(ordered):
                ordered.append((st, cnt))

        for st, cnt in ordered:
            color = STATUS_COLORS.get(st, DEFAULT_COLOR)
            width_pct = max(cnt / max_count * 100, 8)
            pipeline_bars.append(
                f'<div class="pipeline-row">'
                f'<span class="pipeline-label">{_esc(st)}</span>'
                f'<div class="pipeline-bar-wrap">'
                f'<div class="pipeline-bar" style="width:{width_pct}%;background:{color};">{cnt}</div>'
                f'</div></div>'
            )

    # HIDDEN はグレーテキストのみ
    hidden_lines = []
    for st, cnt in sorted(hidden_statuses.items(), key=lambda x: -x[1]):
        hidden_lines.append(f'<span class="hidden-status">{_esc(st)}: {cnt}</span>')

    hidden_html = ''
    if hidden_lines:
        hidden_html = f'<div class="hidden-statuses">{" / ".join(hidden_lines)}</div>'

    pipeline_html = f'''
    <section class="sec">
      <h2 class="sec-title">ステータス別パイプライン</h2>
      {"".join(pipeline_bars)}
      {hidden_html}
    </section>'''

    # 役割別集計テーブル
    role_rows_html = []
    for role, info in sorted(agg['role_counts'].items(), key=lambda x: -x[1]['total']):
        badges = []
        for st, cnt in sorted(info['statuses'].items(), key=lambda x: -x[1]):
            color = STATUS_COLORS.get(st, DEFAULT_COLOR)
            badges.append(f'<span class="status-badge" style="background:{color};">{_esc(st)} {cnt}</span>')
        role_rows_html.append(
            f'<tr><td class="role-name">{_esc(role)}</td>'
            f'<td class="num">{info["total"]}</td>'
            f'<td>{"".join(badges)}</td></tr>'
        )

    role_html = f'''
    <section class="sec">
      <h2 class="sec-title">役割別の応募状況</h2>
      <div class="table-scroll">
        <table class="data-table">
          <thead><tr><th>役割</th><th>応募者数</th><th>ステータス分布</th></tr></thead>
          <tbody>{"".join(role_rows_html)}</tbody>
        </table>
      </div>
    </section>'''

    return summary + pipeline_html + role_html


# ===== タブ2: 対応が必要 =====
def _render_tab_action(agg):
    groups = [
        ('今すぐ対応', agg['action_rows']['high'], True),
        ('確認・評価', agg['action_rows']['medium'], True),
        ('モニタリング', agg['action_rows']['monitoring'], False),
    ]
    # other があればモニタリングの後に追加
    if agg['action_rows']['other']:
        groups.append(('その他', agg['action_rows']['other'], False))

    sections = []
    for group_label, group_rows, expanded in groups:
        if not group_rows:
            continue
        # ステータスごとにサブグループ
        by_status = {}
        for r in group_rows:
            st = r['status'] or '未設定'
            if st not in by_status:
                by_status[st] = []
            by_status[st].append(r)

        sub_sections = []
        for st, st_rows in by_status.items():
            color = STATUS_COLORS.get(st, DEFAULT_COLOR)
            count = len(st_rows)
            open_attr = 'open' if expanded else ''
            # 30件以上は最初20件 + 「すべて表示」
            show_all_btn = ''
            display_rows = st_rows
            if count > 30:
                display_rows = st_rows[:20]
                show_all_btn = f'<div class="show-all-wrap"><button class="show-all-btn" onclick="showAllRows(this)">すべて表示（残り{count - 20}件）</button></div>'

            row_html = _render_applicant_rows(display_rows)
            hidden_html = ''
            if count > 30:
                hidden_rows_html = _render_applicant_rows(st_rows[20:])
                hidden_html = f'<tbody class="hidden-rows" style="display:none;">{hidden_rows_html}</tbody>'

            sub_sections.append(f'''
            <details class="action-group" {open_attr}>
              <summary class="action-summary">
                <span class="status-badge" style="background:{color};">{_esc(st)}</span>
                <span class="action-count">{count}件</span>
              </summary>
              <div class="action-body">
                <table class="data-table">
                  <thead><tr><th>名前</th><th>役割</th><th>期待度</th><th>備考</th><th>面談日</th></tr></thead>
                  <tbody>{row_html}</tbody>
                  {hidden_html}
                </table>
                {show_all_btn}
              </div>
            </details>''')

        sections.append(f'''
        <section class="sec">
          <h2 class="sec-title">{_esc(group_label)}</h2>
          {"".join(sub_sections)}
        </section>''')

    if not sections:
        return '<section class="sec"><p style="color:#94a3b8;text-align:center;">対応が必要な応募者はいません</p></section>'

    return ''.join(sections)


def _render_applicant_rows(rows_list):
    parts = []
    for r in rows_list:
        name_cell = f'<a href="{_attr_esc(r["link"])}" target="_blank" class="applicant-link">{_esc(r["name"])}</a>' if r['link'] else _esc(r['name'])
        note_esc = _attr_esc(r['note'])
        note_short = _esc(r['note'][:40] + '...' if len(r['note']) > 40 else r['note'])
        note_cell = f'<span class="note-text" title="{note_esc}">{note_short}</span>' if r['note'] else ''
        parts.append(
            f'<tr>'
            f'<td>{name_cell}</td>'
            f'<td>{_esc(r["role"])}</td>'
            f'<td><span class="exp-badge">{_esc(r["expectation"])}</span></td>'
            f'<td>{note_cell}</td>'
            f'<td>{_esc(r["interview_date"])}</td>'
            f'</tr>'
        )
    return ''.join(parts)


# ===== タブ3: 採用済み =====
def _render_tab_hired(agg):
    hired_groups = {}
    for r in agg['hired_rows']:
        st = r['status']
        if st not in hired_groups:
            hired_groups[st] = []
        hired_groups[st].append(r)

    # 表示順
    order = ['仮採用', '採用', '契約中']
    sections = []
    for st in order:
        if st not in hired_groups:
            continue
        group = hired_groups[st]
        color = STATUS_COLORS.get(st, DEFAULT_COLOR)
        row_html = []
        for r in group:
            name_cell = f'<a href="{_attr_esc(r["link"])}" target="_blank" class="applicant-link">{_esc(r["name"])}</a>' if r['link'] else _esc(r['name'])
            note_short = _esc(r['note'][:40] + '...' if len(r['note']) > 40 else r['note'])
            note_cell = f'<span class="note-text" title="{_attr_esc(r["note"])}">{note_short}</span>' if r['note'] else ''
            row_html.append(
                f'<tr><td>{name_cell}</td><td>{_esc(r["role"])}</td>'
                f'<td><span class="exp-badge">{_esc(r["expectation"])}</span></td>'
                f'<td>{note_cell}</td></tr>'
            )
        sections.append(f'''
        <section class="sec">
          <h2 class="sec-title"><span class="status-badge" style="background:{color};">{_esc(st)}</span> {len(group)}名</h2>
          <table class="data-table">
            <thead><tr><th>名前</th><th>役割</th><th>期待度</th><th>備考</th></tr></thead>
            <tbody>{"".join(row_html)}</tbody>
          </table>
        </section>''')

    if not sections:
        return '<section class="sec"><p style="color:#94a3b8;text-align:center;">採用済みの応募者はいません</p></section>'

    return ''.join(sections)


# ===== タブ4: 全応募者検索 =====
def _render_tab_search(rows):
    roles = sorted(set(r['role'] for r in rows if r['role']))
    statuses = sorted(set(r['status'] for r in rows if r['status']))
    expectations = sorted(set(r['expectation'] for r in rows if r['expectation']))

    role_options = ''.join(f'<option value="{_attr_esc(v)}">{_esc(v)}</option>' for v in roles)
    status_options = ''.join(f'<option value="{_attr_esc(v)}">{_esc(v)}</option>' for v in statuses)
    exp_options = ''.join(f'<option value="{_attr_esc(v)}">{_esc(v)}</option>' for v in expectations)

    search_data = json.dumps(rows, ensure_ascii=False)

    html = f'''
    <section class="sec">
      <div class="search-bar">
        <input type="text" id="search-name" placeholder="名前で検索..." class="search-input" oninput="searchApplicants()">
      </div>
      <div class="filters">
        <label>役割</label>
        <select id="filter-role" onchange="searchApplicants()">
          <option value="">すべて</option>
          {role_options}
        </select>
        <label>ステータス</label>
        <select id="filter-status" onchange="searchApplicants()">
          <option value="">すべて</option>
          {status_options}
        </select>
        <label>期待度</label>
        <select id="filter-exp" onchange="searchApplicants()">
          <option value="">すべて</option>
          {exp_options}
        </select>
        <span class="filter-count" id="search-count"></span>
      </div>
      <div class="table-scroll">
        <table class="data-table" id="search-table">
          <thead>
            <tr>
              <th class="sortable" data-col="name" onclick="sortColumn('name')">名前 <span class="sort-icon"></span></th>
              <th class="sortable" data-col="role" onclick="sortColumn('role')">役割 <span class="sort-icon"></span></th>
              <th>期待度</th>
              <th class="sortable" data-col="status" onclick="sortColumn('status')">ステータス <span class="sort-icon"></span></th>
              <th>備考</th>
              <th>面談日</th>
            </tr>
          </thead>
          <tbody id="search-tbody"></tbody>
        </table>
      </div>
      <div class="pagination" id="pagination"></div>
    </section>'''

    return html, search_data


# ---------------------------------------------------------------------------
# CSS
# ---------------------------------------------------------------------------
def _css():
    return '''
  :root { --teal: #0d9488; --teal-dark: #0f766e; --coral: #f97316; --gray: #64748b; --border: #e2e8f0; --light: #f8fafc; }
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body { font-family: 'Noto Sans JP', sans-serif; background: var(--light); color: #1e293b; padding: 24px; line-height: 1.6; }
  .container { max-width: 1200px; margin: 0 auto; }

  /* ヘッダー */
  .header { background: linear-gradient(135deg, #0f766e 0%, #0d9488 100%); color: #fff; padding: 24px 28px; border-radius: 12px 12px 0 0; }
  .header h1 { font-size: 20px; margin-bottom: 4px; }
  .header p { font-size: 13px; opacity: 0.9; }

  /* タブバー */
  .tab-bar { display: flex; gap: 0; background: #fff; border-radius: 0 0 12px 12px; border: 1px solid var(--border); border-top: none; margin-bottom: 24px; overflow: hidden; }
  .tab-btn { flex: 1; padding: 14px 20px; font-size: 14px; font-weight: 600; font-family: inherit; border: none; background: #fff; color: var(--gray); cursor: pointer; transition: background 0.2s, color 0.2s; text-align: center; border-right: 1px solid var(--border); }
  .tab-btn:last-child { border-right: none; }
  .tab-btn:hover { background: #f0fdfa; color: var(--teal-dark); }
  .tab-btn.active { background: var(--teal); color: #fff; }
  .tab-content { display: none; }
  .tab-content.active { display: block; }
  .tab-badge { display: inline-block; background: #ef4444; color: #fff; font-size: 11px; font-weight: 700; padding: 1px 7px; border-radius: 10px; margin-left: 6px; vertical-align: middle; }
  .tab-btn.active .tab-badge { background: #fff; color: var(--teal); }

  /* セクション */
  .sec { background: #fff; border-radius: 12px; padding: 20px 24px; margin-bottom: 20px; border: 1px solid var(--border); }
  .sec-title { font-size: 16px; font-weight: 700; color: var(--teal-dark); margin-bottom: 16px; padding-bottom: 8px; border-bottom: 2px solid var(--teal); }

  /* サマリーカード */
  .summary-cards { display: flex; gap: 16px; flex-wrap: wrap; margin-bottom: 24px; }
  .summary-card { background: #f0fdfa; border: 1px solid var(--teal); padding: 20px 28px; border-radius: 12px; text-align: center; flex: 1; min-width: 140px; }
  .summary-card.card-green { background: #f0fdf4; border-color: #16a34a; }
  .summary-card.card-green .summary-num { color: #16a34a; }
  .summary-card.card-blue { background: #eff6ff; border-color: #2563eb; }
  .summary-card.card-blue .summary-num { color: #2563eb; }
  .summary-card.card-coral { background: #fff7ed; border-color: var(--coral); }
  .summary-card.card-coral .summary-num { color: var(--coral); }
  .summary-num { font-size: 32px; font-weight: 700; color: var(--teal); }
  .summary-label { font-size: 13px; color: var(--gray); margin-top: 4px; }

  /* パイプライン */
  .pipeline-row { display: flex; align-items: center; margin-bottom: 8px; }
  .pipeline-label { width: 140px; font-size: 13px; font-weight: 600; text-align: right; padding-right: 12px; flex-shrink: 0; }
  .pipeline-bar-wrap { flex: 1; }
  .pipeline-bar { height: 28px; border-radius: 6px; color: #fff; font-size: 13px; font-weight: 600; display: flex; align-items: center; padding-left: 10px; min-width: 32px; transition: width 0.3s; }
  .hidden-statuses { margin-top: 16px; padding-top: 12px; border-top: 1px solid var(--border); color: #94a3b8; font-size: 13px; }
  .hidden-status { margin-right: 8px; }

  /* ステータスバッジ */
  .status-badge { display: inline-block; padding: 2px 10px; border-radius: 12px; color: #fff; font-size: 12px; font-weight: 600; margin-right: 4px; white-space: nowrap; }
  .exp-badge { display: inline-block; padding: 2px 8px; border-radius: 8px; background: #f1f5f9; color: #475569; font-size: 12px; font-weight: 500; }

  /* データテーブル */
  .data-table { width: 100%; border-collapse: collapse; font-size: 13px; }
  .data-table th { background: #f1f5f9; padding: 10px 12px; text-align: left; font-weight: 600; position: sticky; top: 0; }
  .data-table td { padding: 10px 12px; border-bottom: 1px solid #f1f5f9; }
  .data-table tr:hover { background: #f8fafc; }
  .num { text-align: right; font-weight: 600; }
  .role-name { font-weight: 600; }
  .table-scroll { max-height: 60vh; overflow-y: auto; overflow-x: auto; border: 1px solid var(--border); border-radius: 8px; }

  /* 折りたたみグループ */
  .action-group { margin-bottom: 12px; border: 1px solid var(--border); border-radius: 8px; overflow: hidden; }
  .action-summary { padding: 12px 16px; cursor: pointer; display: flex; align-items: center; gap: 12px; background: #f8fafc; font-weight: 600; font-size: 14px; list-style: none; }
  .action-summary::-webkit-details-marker { display: none; }
  .action-summary::before { content: '▶'; font-size: 10px; color: var(--gray); transition: transform 0.2s; }
  details[open] .action-summary::before { transform: rotate(90deg); }
  .action-count { color: var(--gray); font-weight: 500; font-size: 13px; }
  .action-body { padding: 0; }
  .action-body .data-table { border: none; }

  /* すべて表示ボタン */
  .show-all-wrap { text-align: center; padding: 12px; }
  .show-all-btn { background: none; border: 1px solid var(--border); padding: 8px 20px; border-radius: 8px; font-size: 13px; font-family: inherit; color: var(--teal); cursor: pointer; font-weight: 600; }
  .show-all-btn:hover { background: #f0fdfa; }

  /* 検索バー */
  .search-bar { margin-bottom: 16px; }
  .search-input { width: 100%; padding: 12px 16px; border: 1px solid var(--border); border-radius: 8px; font-size: 14px; font-family: inherit; outline: none; }
  .search-input:focus { border-color: var(--teal); box-shadow: 0 0 0 3px rgba(13,148,136,0.1); }

  /* フィルタ */
  .filters { display: flex; gap: 12px; margin-bottom: 16px; flex-wrap: wrap; align-items: center; }
  .filters label { font-size: 13px; font-weight: 600; color: var(--gray); }
  .filters select { padding: 6px 12px; border: 1px solid var(--border); border-radius: 8px; font-size: 13px; font-family: inherit; background: #fff; }
  .filter-count { font-size: 13px; color: var(--gray); margin-left: auto; }

  /* ソート */
  .sortable { cursor: pointer; user-select: none; }
  .sortable:hover { background: #e2e8f0; }
  .sort-icon { font-size: 10px; color: var(--gray); }

  /* ページネーション */
  .pagination { display: flex; justify-content: center; align-items: center; gap: 8px; margin-top: 16px; flex-wrap: wrap; }
  .page-btn { padding: 6px 14px; border: 1px solid var(--border); border-radius: 6px; background: #fff; font-size: 13px; font-family: inherit; cursor: pointer; color: #1e293b; }
  .page-btn:hover { background: #f0fdfa; border-color: var(--teal); }
  .page-btn.active { background: var(--teal); color: #fff; border-color: var(--teal); }
  .page-btn:disabled { opacity: 0.4; cursor: default; }
  .page-info { font-size: 13px; color: var(--gray); }

  /* リンク */
  .applicant-link { color: var(--teal); text-decoration: none; font-weight: 500; }
  .applicant-link:hover { text-decoration: underline; }
  .note-text { cursor: help; border-bottom: 1px dotted var(--gray); }

  /* レスポンシブ */
  @media (max-width: 768px) {
    body { padding: 12px; }
    .tab-bar { overflow-x: auto; -webkit-overflow-scrolling: touch; }
    .tab-btn { min-width: 100px; flex: none; white-space: nowrap; font-size: 13px; padding: 12px 14px; }
    .summary-cards { gap: 8px; }
    .summary-card { padding: 14px 16px; min-width: 100px; }
    .summary-num { font-size: 24px; }
    .pipeline-label { width: 100px; font-size: 12px; }
    .filters { flex-direction: column; gap: 8px; }
    .table-scroll { max-height: 50vh; }
  }
'''


# ---------------------------------------------------------------------------
# JS
# ---------------------------------------------------------------------------
def _js():
    return r'''
// ==== タブ切り替え ====
document.querySelectorAll('.tab-btn').forEach(function(btn) {
  btn.addEventListener('click', function() {
    var tab = this.dataset.tab;
    document.querySelectorAll('.tab-btn').forEach(function(b) { b.classList.remove('active'); });
    document.querySelectorAll('.tab-content').forEach(function(c) { c.classList.remove('active'); });
    this.classList.add('active');
    document.querySelector('.tab-content[data-tab="' + tab + '"]').classList.add('active');
    history.replaceState(null, '', '#' + tab);
    // 検索タブ初回表示時にデータ描画
    if (tab === 'search' && !window._searchInit) {
      window._searchInit = true;
      searchApplicants();
    }
  });
});
// URL hash support
(function() {
  var hash = location.hash.replace('#', '');
  if (hash) {
    var btn = document.querySelector('.tab-btn[data-tab="' + hash + '"]');
    if (btn) btn.click();
  }
})();

// ==== すべて表示ボタン ====
function showAllRows(btn) {
  var body = btn.closest('.action-body');
  var hidden = body.querySelector('.hidden-rows');
  if (hidden) {
    hidden.style.display = '';
  }
  btn.parentElement.style.display = 'none';
}

// ==== 検索タブ: ページネーション・フィルタ・ソート ====
var PAGE_SIZE = 50;
var currentPage = 1;
var filteredData = [];
var sortCol = '';
var sortAsc = true;

var STATUS_COLORS = {
  '採用': '#16a34a', '仮採用': '#22c55e', '契約中': '#15803d',
  '面談予定': '#2563eb', '面談日程調節中': '#3b82f6', '外部連絡申請中': '#ca8a04',
  'トライアル実施中': '#8b5cf6', '課題提出済み': '#a855f7', '課題確認済み': '#7c3aed', '採用基準〇': '#059669',
  'ヒアリング中': '#0891b2', '返信保留中': '#ea580c',
  '不採用': '#94a3b8', '離脱': '#cbd5e1', '採用基準✖': '#9ca3af',
  '別案件流し': '#6b7280', '不採用→再連絡中': '#d97706', '返信なし': '#9ca3af',
  '未対応': '#ef4444', '未設定': '#d1d5db'
};

function escHtml(t) {
  if (!t) return '';
  return String(t).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function searchApplicants() {
  var nameQ = (document.getElementById('search-name').value || '').toLowerCase();
  var roleF = document.getElementById('filter-role').value;
  var statusF = document.getElementById('filter-status').value;
  var expF = document.getElementById('filter-exp').value;

  filteredData = window.SEARCH_DATA.filter(function(r) {
    if (nameQ && (r.name || '').toLowerCase().indexOf(nameQ) === -1) return false;
    if (roleF && r.role !== roleF) return false;
    if (statusF && r.status !== statusF) return false;
    if (expF && r.expectation !== expF) return false;
    return true;
  });

  // ソート適用
  if (sortCol) {
    filteredData.sort(function(a, b) {
      var va = (a[sortCol] || '').toLowerCase();
      var vb = (b[sortCol] || '').toLowerCase();
      if (va < vb) return sortAsc ? -1 : 1;
      if (va > vb) return sortAsc ? 1 : -1;
      return 0;
    });
  }

  currentPage = 1;
  renderPage();
}

function renderPage() {
  var total = filteredData.length;
  var totalPages = Math.max(1, Math.ceil(total / PAGE_SIZE));
  if (currentPage > totalPages) currentPage = totalPages;

  var start = (currentPage - 1) * PAGE_SIZE;
  var end = Math.min(start + PAGE_SIZE, total);
  var pageData = filteredData.slice(start, end);

  // テーブル描画
  var tbody = document.getElementById('search-tbody');
  var html = '';
  for (var i = 0; i < pageData.length; i++) {
    var r = pageData[i];
    var color = STATUS_COLORS[r.status] || '#64748b';
    var nameCell = r.link
      ? '<a href="' + escHtml(r.link) + '" target="_blank" class="applicant-link">' + escHtml(r.name) + '</a>'
      : escHtml(r.name);
    var noteShort = r.note ? (r.note.length > 40 ? escHtml(r.note.substring(0, 40)) + '...' : escHtml(r.note)) : '';
    var noteCell = r.note ? '<span class="note-text" title="' + escHtml(r.note) + '">' + noteShort + '</span>' : '';
    html += '<tr>'
      + '<td>' + nameCell + '</td>'
      + '<td>' + escHtml(r.role) + '</td>'
      + '<td><span class="exp-badge">' + escHtml(r.expectation) + '</span></td>'
      + '<td><span class="status-badge" style="background:' + color + ';">' + escHtml(r.status) + '</span></td>'
      + '<td>' + noteCell + '</td>'
      + '<td>' + escHtml(r.interview_date) + '</td>'
      + '</tr>';
  }
  tbody.innerHTML = html;

  // カウント表示
  document.getElementById('search-count').textContent = total + '件中 ' + (total > 0 ? start + 1 : 0) + '-' + end + '件表示';

  // ページネーション
  renderPagination(totalPages);

  // ソートアイコン更新
  document.querySelectorAll('#search-table .sort-icon').forEach(function(el) {
    var col = el.closest('th').dataset.col;
    if (col === sortCol) {
      el.textContent = sortAsc ? '▲' : '▼';
    } else {
      el.textContent = '';
    }
  });
}

function renderPagination(totalPages) {
  var pag = document.getElementById('pagination');
  if (totalPages <= 1) { pag.innerHTML = ''; return; }

  var html = '';
  html += '<button class="page-btn" onclick="goPage(1)" ' + (currentPage === 1 ? 'disabled' : '') + '>&laquo;</button>';
  html += '<button class="page-btn" onclick="goPage(' + (currentPage - 1) + ')" ' + (currentPage === 1 ? 'disabled' : '') + '>&lsaquo;</button>';

  // ページ番号（最大7個表示）
  var startP = Math.max(1, currentPage - 3);
  var endP = Math.min(totalPages, startP + 6);
  if (endP - startP < 6) startP = Math.max(1, endP - 6);

  for (var p = startP; p <= endP; p++) {
    html += '<button class="page-btn' + (p === currentPage ? ' active' : '') + '" onclick="goPage(' + p + ')">' + p + '</button>';
  }

  html += '<button class="page-btn" onclick="goPage(' + (currentPage + 1) + ')" ' + (currentPage === totalPages ? 'disabled' : '') + '>&rsaquo;</button>';
  html += '<button class="page-btn" onclick="goPage(' + totalPages + ')" ' + (currentPage === totalPages ? 'disabled' : '') + '>&raquo;</button>';

  pag.innerHTML = html;
}

function goPage(p) {
  var totalPages = Math.max(1, Math.ceil(filteredData.length / PAGE_SIZE));
  if (p < 1 || p > totalPages) return;
  currentPage = p;
  renderPage();
  // スクロールトップ
  document.getElementById('search-table').scrollIntoView({behavior: 'smooth', block: 'start'});
}

function sortColumn(col) {
  if (sortCol === col) {
    sortAsc = !sortAsc;
  } else {
    sortCol = col;
    sortAsc = true;
  }
  searchApplicants();
}
'''


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------
def main():
    rows = load_csv()
    if not rows:
        print("No data found. Exiting.")
        return

    agg = aggregate(rows)

    # JSON中間データ出力
    output_json = BASE / 'recruitment_data.json'
    with open(output_json, 'w', encoding='utf-8') as f:
        json.dump({'rows': rows, 'aggregation': agg}, f, ensure_ascii=False, indent=2, default=str)

    # HTML生成
    html = render_html(rows, agg)
    output_html = BASE / 'index.html'
    with open(output_html, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"Generated {output_json}")
    print(f"Generated {output_html}")
    print(f"  Total applicants: {agg['total']} (空行除外済み)")
    print(f"  Active: {agg['active_count']}, Hired: {agg['hired_count']}, Action needed: {agg['action_count']}")


if __name__ == '__main__':
    main()
