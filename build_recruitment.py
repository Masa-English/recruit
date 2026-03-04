#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
採用管理ダッシュボード — タブベースUI
CSVをパースし、6タブ構成のHTMLダッシュボードを生成する。
"""
import csv
import json
import os
from datetime import date
from pathlib import Path
from statistics import median

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
                'date_applied': (row.get('応募日') or '').strip(),
                'date_hearing': (row.get('ヒアリング日') or '').strip(),
                'date_interview': (row.get('面談実施日') or '').strip(),
                'date_trial': (row.get('課題・トライアル日') or '').strip(),
                'date_hired': (row.get('採用日') or '').strip(),
            })
    return rows


# ---------------------------------------------------------------------------
# WBS / 予実ヘルパー
# ---------------------------------------------------------------------------
WBS_STAGES = [
    ('応募', 'date_applied'),
    ('ヒアリング', 'date_hearing'),
    ('面談', 'date_interview'),
    ('トライアル', 'date_trial'),
    ('採用', 'date_hired'),
]

WBS_TRANSITIONS = [
    ('応募→ヒアリング', 'date_applied', 'date_hearing'),
    ('ヒアリング→面談', 'date_hearing', 'date_interview'),
    ('面談→トライアル', 'date_interview', 'date_trial'),
    ('トライアル→採用', 'date_trial', 'date_hired'),
]


def _parse_date(s):
    """YYYY-MM-DD を date に変換。無効なら None。"""
    if not s:
        return None
    try:
        parts = s.split('-')
        return date(int(parts[0]), int(parts[1]), int(parts[2]))
    except (ValueError, IndexError):
        return None


def calc_pipeline_stages(rows):
    """各応募者の最も進んだステージを判定し、ステージ別人数を返す。"""
    stage_counts = {label: 0 for label, _ in WBS_STAGES}
    for r in rows:
        last_stage = None
        for label, key in WBS_STAGES:
            if r.get(key):
                last_stage = label
        if last_stage:
            stage_counts[last_stage] += 1
    return stage_counts


def calc_durations(rows):
    """ステージ間所要日数を計算。"""
    results = []
    for label, key_from, key_to in WBS_TRANSITIONS:
        days_list = []
        for r in rows:
            d1 = _parse_date(r.get(key_from, ''))
            d2 = _parse_date(r.get(key_to, ''))
            if d1 and d2:
                diff = (d2 - d1).days
                if diff >= 0:
                    days_list.append(diff)
        if days_list:
            results.append({
                'label': label,
                'avg': round(sum(days_list) / len(days_list), 1),
                'median': round(median(days_list), 1),
                'max': max(days_list),
                'count': len(days_list),
            })
        else:
            results.append({'label': label, 'avg': None, 'median': None, 'max': None, 'count': 0})
    return results


def load_targets():
    """役割_目標.csv を読み込み。ファイルが無ければ None。"""
    target_path = BASE / '役割_目標.csv'
    if not target_path.exists():
        return None
    targets = {}
    with open(target_path, 'r', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            role = (row.get('役割') or '').strip()
            try:
                num = int((row.get('目標人数') or '0').strip())
            except ValueError:
                num = 0
            if role:
                targets[role] = num
    return targets


def calc_yojitsu(hired_rows, targets):
    """目標 vs 実績を計算。"""
    # 実績: 採用済みの役割別人数と名前
    actual = {}
    for r in hired_rows:
        role = r['role'] or '未設定'
        if role not in actual:
            actual[role] = {'count': 0, 'names': []}
        actual[role]['count'] += 1
        actual[role]['names'].append(r['name'])

    # 役割別データ構築
    all_roles = set()
    if targets:
        all_roles.update(targets.keys())
    all_roles.update(actual.keys())

    result = []
    total_target = 0
    total_actual = 0
    unset_roles = []  # 目標未設定だが実績がある

    for role in sorted(all_roles):
        target = targets.get(role, 0) if targets else 0
        act = actual.get(role, {'count': 0, 'names': []})
        pct = (act['count'] / target * 100) if target > 0 else (100 if act['count'] > 0 else 0)
        has_target = targets is not None and role in targets
        result.append({
            'role': role,
            'target': target,
            'actual': act['count'],
            'pct': round(pct, 1),
            'names': act['names'],
            'has_target': has_target,
        })
        if has_target:
            total_target += target
        total_actual += act['count']

        if not has_target and act['count'] > 0:
            unset_roles.append({'role': role, 'count': act['count'], 'names': act['names']})

    return {
        'items': [i for i in result if i['has_target']],
        'total_target': total_target,
        'total_actual': total_actual,
        'total_pct': round(total_actual / total_target * 100, 1) if total_target > 0 else 0,
        'unset_roles': unset_roles,
        'targets_loaded': targets is not None,
    }


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

    # WBS計算
    pipeline_stages = calc_pipeline_stages(rows)
    durations = calc_durations(rows)

    # 個別タイムライン: 2つ以上の日付がある応募者
    timeline_rows = []
    for r in rows:
        dates_filled = sum(1 for _, key in WBS_STAGES if r.get(key))
        if dates_filled >= 2:
            d_first = None
            d_last = None
            for _, key in WBS_STAGES:
                d = _parse_date(r.get(key, ''))
                if d:
                    if d_first is None or d < d_first:
                        d_first = d
                    if d_last is None or d > d_last:
                        d_last = d
            total_days = (d_last - d_first).days if d_first and d_last else None
            timeline_rows.append({**r, 'total_days': total_days})

    # 予実計算
    targets = load_targets()
    yojitsu = calc_yojitsu(hired_rows, targets)

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
        'pipeline_stages': pipeline_stages,
        'durations': durations,
        'timeline_rows': timeline_rows,
        'yojitsu': yojitsu,
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
    # --- タブ4: WBS進捗 ---
    tab4_wbs = _render_tab_wbs(agg)
    # --- タブ5: 予実管理 ---
    tab5_yojitsu = _render_tab_yojitsu(agg)
    # --- タブ6: 全応募者検索 (JSON埋め込み) ---
    tab6_html, search_json = _render_tab_search(rows)

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
    <button class="tab-btn" data-tab="wbs">WBS進捗</button>
    <button class="tab-btn" data-tab="yojitsu">予実管理</button>
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
  <div class="tab-content" data-tab="wbs">
    {tab4_wbs}
  </div>
  <div class="tab-content" data-tab="yojitsu">
    {tab5_yojitsu}
  </div>
  <div class="tab-content" data-tab="search">
    {tab6_html}
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


# ===== タブ4: WBS進捗 =====
def _render_tab_wbs(agg):
    stages = agg['pipeline_stages']
    durations = agg['durations']
    timeline_rows = agg['timeline_rows']

    total_with_dates = sum(stages.values())

    if total_with_dates == 0:
        return '<section class="sec"><p style="color:#94a3b8;text-align:center;padding:40px 0;">WBS日付データがまだ入力されていません</p></section>'

    # --- セクション1: パイプラインファネル ---
    funnel_bars = []
    prev_count = None
    for label, _ in WBS_STAGES:
        count = stages[label]
        if count == 0 and prev_count is None:
            prev_count = 0
            continue
        pct_text = ''
        if prev_count and prev_count > 0:
            rate = count / prev_count * 100
            pct_text = f' ({rate:.0f}%)'
        width_pct = max(count / max(total_with_dates, 1) * 100, 6) if count > 0 else 0
        funnel_bars.append(
            f'<div class="pipeline-row">'
            f'<span class="pipeline-label">{_esc(label)}</span>'
            f'<div class="pipeline-bar-wrap">'
            f'<div class="funnel-bar" style="width:{width_pct}%;">{count}{pct_text}</div>'
            f'</div></div>'
        )
        prev_count = count

    funnel_html = f'''
    <section class="sec">
      <h2 class="sec-title">パイプラインファネル</h2>
      <p style="font-size:13px;color:#64748b;margin-bottom:12px;">各ステージの到達人数（最も進んだステージで集計）。括弧内は前ステージからの転換率。</p>
      {"".join(funnel_bars)}
    </section>'''

    # --- セクション2: ステージ間所要日数 ---
    dur_rows = []
    for d in durations:
        if d['count'] == 0:
            dur_rows.append(
                f'<tr><td>{_esc(d["label"])}</td><td colspan="3" style="color:#94a3b8;">データなし</td><td>0</td></tr>'
            )
        else:
            avg_cls = 'dur-green' if d['avg'] < 5 else ('dur-orange' if d['avg'] < 15 else 'dur-red')
            med_cls = 'dur-green' if d['median'] < 5 else ('dur-orange' if d['median'] < 15 else 'dur-red')
            max_cls = 'dur-green' if d['max'] < 5 else ('dur-orange' if d['max'] < 15 else 'dur-red')
            dur_rows.append(
                f'<tr><td>{_esc(d["label"])}</td>'
                f'<td class="{avg_cls}">{d["avg"]}日</td>'
                f'<td class="{med_cls}">{d["median"]}日</td>'
                f'<td class="{max_cls}">{d["max"]}日</td>'
                f'<td>{d["count"]}</td></tr>'
            )

    dur_html = f'''
    <section class="sec">
      <h2 class="sec-title">ステージ間所要日数</h2>
      <div class="table-scroll">
        <table class="data-table">
          <thead><tr><th>遷移</th><th>平均</th><th>中央値</th><th>最大</th><th>件数</th></tr></thead>
          <tbody>{"".join(dur_rows)}</tbody>
        </table>
      </div>
    </section>'''

    # --- セクション3: 個別タイムライン ---
    if timeline_rows:
        tl_rows_html = []
        for r in timeline_rows:
            name_cell = f'<a href="{_attr_esc(r["link"])}" target="_blank" class="applicant-link">{_esc(r["name"])}</a>' if r.get('link') else _esc(r['name'])
            total_d = f'{r["total_days"]}日' if r.get('total_days') is not None else '-'
            tl_rows_html.append(
                f'<tr>'
                f'<td>{name_cell}</td>'
                f'<td>{_esc(r.get("role", ""))}</td>'
                f'<td>{_esc(r.get("date_applied", ""))}</td>'
                f'<td>{_esc(r.get("date_hearing", ""))}</td>'
                f'<td>{_esc(r.get("date_interview", ""))}</td>'
                f'<td>{_esc(r.get("date_trial", ""))}</td>'
                f'<td>{_esc(r.get("date_hired", ""))}</td>'
                f'<td class="num">{total_d}</td>'
                f'</tr>'
            )
        tl_html = f'''
        <section class="sec">
          <h2 class="sec-title">個別タイムライン</h2>
          <p style="font-size:13px;color:#64748b;margin-bottom:12px;">2つ以上の日付が入力されている応募者を表示</p>
          <div class="table-scroll">
            <table class="data-table">
              <thead><tr><th>名前</th><th>役割</th><th>応募日</th><th>ヒアリング日</th><th>面談実施日</th><th>課題日</th><th>採用日</th><th>合計日数</th></tr></thead>
              <tbody>{"".join(tl_rows_html)}</tbody>
            </table>
          </div>
        </section>'''
    else:
        tl_html = ''

    return funnel_html + dur_html + tl_html


# ===== タブ5: 予実管理 =====
def _render_tab_yojitsu(agg):
    yj = agg['yojitsu']

    if not yj['targets_loaded']:
        return '<section class="sec"><p style="color:#94a3b8;text-align:center;padding:40px 0;">役割_目標.csv が未作成です。目標人数を設定してください。</p></section>'

    # --- セクション1: サマリーカード ---
    pct_color = '#16a34a' if yj['total_pct'] >= 100 else ('#0d9488' if yj['total_pct'] >= 50 else '#f97316')
    summary = f'''
    <div class="summary-cards">
      <div class="summary-card">
        <div class="summary-num">{yj['total_target']}</div>
        <div class="summary-label">目標合計</div>
      </div>
      <div class="summary-card card-green">
        <div class="summary-num">{yj['total_actual']}</div>
        <div class="summary-label">採用実績</div>
      </div>
      <div class="summary-card" style="border-color:{pct_color};">
        <div class="summary-num" style="color:{pct_color};">{yj['total_pct']}%</div>
        <div class="summary-label">達成率</div>
      </div>
    </div>'''

    # --- セクション2: 役割別プログレスバー ---
    progress_items = []
    for item in yj['items']:
        pct = item['pct']
        if pct >= 100:
            bar_color = '#16a34a'
            check = ' ✓'
        elif pct >= 50:
            bar_color = '#0d9488'
            check = ''
        elif pct >= 25:
            bar_color = '#f97316'
            check = ''
        else:
            bar_color = '#ef4444'
            check = ''
        bar_width = min(pct, 100)
        names_html = ', '.join(_esc(n) for n in item['names']) if item['names'] else 'なし'
        detail_id = f'yj-detail-{_attr_esc(item["role"])}'
        progress_items.append(f'''
        <div class="progress-item">
          <div class="progress-header" onclick="toggleYojitsuDetail(this)">
            <span class="progress-role">{_esc(item['role'])}</span>
            <span class="progress-nums">{item['actual']} / {item['target']}{check}</span>
          </div>
          <div class="progress-bar-wrap">
            <div class="progress-bar" style="width:{bar_width}%;background:{bar_color};"></div>
            <span class="progress-pct">{pct:.0f}%</span>
          </div>
          <div class="progress-detail" style="display:none;">
            <p class="progress-names">{names_html}</p>
          </div>
        </div>''')

    progress_html = f'''
    <section class="sec">
      <h2 class="sec-title">役割別 目標 vs 実績</h2>
      {"".join(progress_items)}
    </section>'''

    # --- セクション3: 目標未設定の役割 ---
    unset_html = ''
    if yj['unset_roles']:
        unset_items = []
        for u in yj['unset_roles']:
            names = ', '.join(_esc(n) for n in u['names'])
            unset_items.append(f'<tr><td class="role-name">{_esc(u["role"])}</td><td class="num">{u["count"]}</td><td>{names}</td></tr>')
        unset_html = f'''
        <section class="sec">
          <h2 class="sec-title">目標未設定の役割（実績あり）</h2>
          <p style="font-size:13px;color:#64748b;margin-bottom:12px;">役割_目標.csv に目標が設定されていない役割です。</p>
          <div class="table-scroll">
            <table class="data-table">
              <thead><tr><th>役割</th><th>採用数</th><th>採用者</th></tr></thead>
              <tbody>{"".join(unset_items)}</tbody>
            </table>
          </div>
        </section>'''

    return summary + progress_html + unset_html


# ===== タブ6: 全応募者検索 =====
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

  /* WBS ファネルバー */
  .funnel-bar { height: 32px; border-radius: 6px; color: #fff; font-size: 13px; font-weight: 600; display: flex; align-items: center; padding-left: 10px; min-width: 40px; background: linear-gradient(90deg, #0d9488, #14b8a6); transition: width 0.3s; }

  /* WBS 所要日数 色分け */
  .dur-green { color: #16a34a; font-weight: 600; }
  .dur-orange { color: #f97316; font-weight: 600; }
  .dur-red { color: #ef4444; font-weight: 600; }

  /* 予実 プログレスバー */
  .progress-item { margin-bottom: 16px; }
  .progress-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px; cursor: pointer; padding: 4px 0; }
  .progress-header:hover { opacity: 0.8; }
  .progress-role { font-weight: 600; font-size: 14px; }
  .progress-nums { font-size: 14px; font-weight: 600; color: var(--gray); }
  .progress-bar-wrap { position: relative; background: #f1f5f9; border-radius: 8px; height: 24px; overflow: hidden; }
  .progress-bar { height: 100%; border-radius: 8px; transition: width 0.4s ease; min-width: 0; }
  .progress-pct { position: absolute; right: 10px; top: 50%; transform: translateY(-50%); font-size: 12px; font-weight: 600; color: #475569; }
  .progress-detail { padding: 8px 0 0 0; }
  .progress-names { font-size: 13px; color: var(--gray); padding: 8px 12px; background: #f8fafc; border-radius: 6px; }

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

// ==== 予実管理: 詳細トグル ====
function toggleYojitsuDetail(el) {
  var detail = el.parentElement.querySelector('.progress-detail');
  if (detail) {
    detail.style.display = detail.style.display === 'none' ? 'block' : 'none';
  }
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
