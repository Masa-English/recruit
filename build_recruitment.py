#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
採用管理ダッシュボード用のCSVをパースし、HTMLに埋め込むJSONを生成する。
実行後、index.html が更新される。
"""
import csv
import json
import os
from pathlib import Path

BASE = Path(__file__).parent


def load_csv():
    """採用管理CSVを読み込み、辞書リストで返す"""
    csv_path = BASE / '採用管理.csv'
    if not csv_path.exists():
        print(f"Warning: {csv_path} not found")
        return []

    with open(csv_path, 'r', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        rows = []
        for row in reader:
            rows.append({
                'name': (row.get('お名前') or '').strip(),
                'link': (row.get('Crowdworks返信リンク') or '').strip(),
                'role': (row.get('役割') or '').strip(),
                'expectation': (row.get('期待度') or '').strip(),
                'status': (row.get('現状のステータス') or '').strip(),
                'note': (row.get('備考') or '').strip(),
                'interview_date': (row.get('面談日') or '').strip(),
            })
    return rows


def aggregate(rows):
    """集計データを生成"""
    total = len(rows)
    status_counts = {}
    role_counts = {}
    expectation_counts = {}
    interview_list = []

    for r in rows:
        # ステータス別
        st = r['status'] or '未設定'
        status_counts[st] = status_counts.get(st, 0) + 1

        # 役割別
        role = r['role'] or '未設定'
        if role not in role_counts:
            role_counts[role] = {'total': 0, 'statuses': {}}
        role_counts[role]['total'] += 1
        role_counts[role]['statuses'][st] = role_counts[role]['statuses'].get(st, 0) + 1

        # 期待度別
        exp = r['expectation'] or '未設定'
        expectation_counts[exp] = expectation_counts.get(exp, 0) + 1

        # 面談予定者
        if st == '面談予定':
            interview_list.append(r)

    return {
        'total': total,
        'status_counts': status_counts,
        'role_counts': role_counts,
        'expectation_counts': expectation_counts,
        'interview_list': interview_list,
        'kari_saiyo': status_counts.get('仮採用', 0),
        'mendan_yotei': status_counts.get('面談予定', 0),
        'mitaiou': status_counts.get('未対応', 0),
    }


def _esc(t):
    if not t:
        return ''
    return str(t).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')


def _attr_esc(t):
    if not t:
        return ''
    return str(t).replace('&', '&amp;').replace('"', '&quot;').replace('<', '&lt;').replace('>', '&gt;')


def render_html(rows, agg):
    """ダッシュボードHTMLをレンダリング"""
    data_js = json.dumps(rows, ensure_ascii=False)

    # --- セクション A: サマリーカード ---
    summary_html = f'''
    <div class="summary-cards">
      <div class="summary-card">
        <div class="summary-num">{agg['total']}</div>
        <div class="summary-label">総応募者数</div>
      </div>
      <div class="summary-card card-green">
        <div class="summary-num">{agg['kari_saiyo']}</div>
        <div class="summary-label">仮採用</div>
      </div>
      <div class="summary-card card-blue">
        <div class="summary-num">{agg['mendan_yotei']}</div>
        <div class="summary-label">面談予定</div>
      </div>
      <div class="summary-card card-orange">
        <div class="summary-num">{agg['mitaiou']}</div>
        <div class="summary-label">未対応</div>
      </div>
    </div>'''

    # --- セクション B: ステータス別パイプライン ---
    pipeline_order = ['未対応', '外部連絡申請中', '面談予定', '仮採用', '不採用', '離脱']
    status_colors = {
        '仮採用': '#16a34a',
        '面談予定': '#2563eb',
        '外部連絡申請中': '#ca8a04',
        '不採用': '#94a3b8',
        '離脱': '#dc2626',
        '未対応': '#ea580c',
    }
    max_count = max(agg['status_counts'].values()) if agg['status_counts'] else 1
    pipeline_bars = []
    for st in pipeline_order:
        cnt = agg['status_counts'].get(st, 0)
        if cnt == 0:
            continue
        color = status_colors.get(st, '#64748b')
        width_pct = max(cnt / max_count * 100, 8)
        pipeline_bars.append(
            f'<div class="pipeline-row">'
            f'<span class="pipeline-label">{_esc(st)}</span>'
            f'<div class="pipeline-bar-wrap">'
            f'<div class="pipeline-bar" style="width:{width_pct}%;background:{color};">{cnt}</div>'
            f'</div></div>'
        )
    # 未分類ステータスも表示
    for st, cnt in agg['status_counts'].items():
        if st not in pipeline_order and cnt > 0:
            color = '#64748b'
            width_pct = max(cnt / max_count * 100, 8)
            pipeline_bars.append(
                f'<div class="pipeline-row">'
                f'<span class="pipeline-label">{_esc(st)}</span>'
                f'<div class="pipeline-bar-wrap">'
                f'<div class="pipeline-bar" style="width:{width_pct}%;background:{color};">{cnt}</div>'
                f'</div></div>'
            )
    pipeline_html = '\n'.join(pipeline_bars)

    # --- セクション C: 役割別の応募状況 ---
    role_rows_html = []
    for role, info in sorted(agg['role_counts'].items(), key=lambda x: -x[1]['total']):
        status_badges = []
        for st, cnt in sorted(info['statuses'].items()):
            color = status_colors.get(st, '#64748b')
            status_badges.append(f'<span class="status-badge" style="background:{color};">{_esc(st)} {cnt}</span>')
        role_rows_html.append(
            f'<tr><td class="role-name">{_esc(role)}</td>'
            f'<td class="num">{info["total"]}</td>'
            f'<td>{"".join(status_badges)}</td></tr>'
        )
    role_table_html = '\n'.join(role_rows_html)

    # --- セクション D: 応募者一覧テーブル ---
    # フィルタ用ユニーク値
    roles = sorted(set(r['role'] for r in rows if r['role']))
    statuses = sorted(set(r['status'] for r in rows if r['status']))
    expectations = sorted(set(r['expectation'] for r in rows if r['expectation']))

    role_options = ''.join(f'<option value="{_attr_esc(v)}">{_esc(v)}</option>' for v in roles)
    status_options = ''.join(f'<option value="{_attr_esc(v)}">{_esc(v)}</option>' for v in statuses)
    exp_options = ''.join(f'<option value="{_attr_esc(v)}">{_esc(v)}</option>' for v in expectations)

    applicant_rows = []
    for r in rows:
        st = r['status']
        color = status_colors.get(st, '#64748b')
        name_cell = f'<a href="{_attr_esc(r["link"])}" target="_blank" class="applicant-link">{_esc(r["name"])}</a>' if r['link'] else _esc(r['name'])
        note_esc = _attr_esc(r['note'])
        note_short = _esc(r['note'][:30] + '...' if len(r['note']) > 30 else r['note'])
        note_cell = f'<span class="note-text" title="{note_esc}">{note_short}</span>' if r['note'] else ''
        applicant_rows.append(
            f'<tr data-role="{_attr_esc(r["role"])}" data-status="{_attr_esc(r["status"])}" data-exp="{_attr_esc(r["expectation"])}">'
            f'<td>{name_cell}</td>'
            f'<td>{_esc(r["role"])}</td>'
            f'<td><span class="exp-badge">{_esc(r["expectation"])}</span></td>'
            f'<td><span class="status-badge" style="background:{color};">{_esc(st)}</span></td>'
            f'<td>{note_cell}</td>'
            f'<td>{_esc(r["interview_date"])}</td>'
            f'</tr>'
        )
    applicant_table_html = '\n'.join(applicant_rows)

    # --- セクション E: 面談予定リスト ---
    interview_rows = []
    for r in sorted(agg['interview_list'], key=lambda x: x.get('interview_date') or 'z'):
        interview_rows.append(
            f'<tr>'
            f'<td>{_esc(r["interview_date"] or "未定")}</td>'
            f'<td>{_esc(r["name"])}</td>'
            f'<td>{_esc(r["role"])}</td>'
            f'<td>{_esc(r["expectation"])}</td>'
            f'<td class="note-cell">{_esc(r["note"])}</td>'
            f'</tr>'
        )
    interview_table_html = '\n'.join(interview_rows) if interview_rows else '<tr><td colspan="5" style="text-align:center;color:#94a3b8;">面談予定の応募者はいません</td></tr>'

    return f'''<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>採用管理ダッシュボード</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
  :root {{ --teal: #0d9488; --teal-dark: #0f766e; --coral: #f97316; --gray: #64748b; --border: #e2e8f0; --light: #f8fafc; }}
  * {{ margin: 0; padding: 0; box-sizing: border-box; }}
  body {{ font-family: 'Noto Sans JP', sans-serif; background: var(--light); color: #1e293b; padding: 24px; line-height: 1.6; }}
  .container {{ max-width: 1200px; margin: 0 auto; }}
  .header {{ background: linear-gradient(135deg, #0f766e 0%, #0d9488 100%); color: #fff; padding: 24px 28px; border-radius: 12px; margin-bottom: 24px; }}
  .header h1 {{ font-size: 20px; margin-bottom: 4px; }}
  .header p {{ font-size: 13px; opacity: 0.9; }}
  .header-links {{ font-size: 13px; margin-top: 12px; opacity: 0.9; }}
  .header-links a {{ color: #fff; text-decoration: underline; margin-right: 16px; }}
  .sec {{ background: #fff; border-radius: 12px; padding: 20px 24px; margin-bottom: 20px; border: 1px solid var(--border); }}
  .sec-title {{ font-size: 16px; font-weight: 700; color: var(--teal-dark); margin-bottom: 16px; padding-bottom: 8px; border-bottom: 2px solid var(--teal); }}

  /* サマリーカード */
  .summary-cards {{ display: flex; gap: 16px; flex-wrap: wrap; margin-bottom: 24px; }}
  .summary-card {{ background: #f0fdfa; border: 1px solid var(--teal); padding: 20px 28px; border-radius: 12px; text-align: center; flex: 1; min-width: 140px; }}
  .summary-card.card-green {{ background: #f0fdf4; border-color: #16a34a; }}
  .summary-card.card-green .summary-num {{ color: #16a34a; }}
  .summary-card.card-blue {{ background: #eff6ff; border-color: #2563eb; }}
  .summary-card.card-blue .summary-num {{ color: #2563eb; }}
  .summary-card.card-orange {{ background: #fff7ed; border-color: #ea580c; }}
  .summary-card.card-orange .summary-num {{ color: #ea580c; }}
  .summary-num {{ font-size: 32px; font-weight: 700; color: var(--teal); }}
  .summary-label {{ font-size: 13px; color: var(--gray); margin-top: 4px; }}

  /* パイプライン */
  .pipeline-row {{ display: flex; align-items: center; margin-bottom: 8px; }}
  .pipeline-label {{ width: 130px; font-size: 13px; font-weight: 600; text-align: right; padding-right: 12px; flex-shrink: 0; }}
  .pipeline-bar-wrap {{ flex: 1; }}
  .pipeline-bar {{ height: 28px; border-radius: 6px; color: #fff; font-size: 13px; font-weight: 600; display: flex; align-items: center; padding-left: 10px; min-width: 32px; transition: width 0.3s; }}

  /* ステータスバッジ */
  .status-badge {{ display: inline-block; padding: 2px 10px; border-radius: 12px; color: #fff; font-size: 12px; font-weight: 600; margin-right: 4px; white-space: nowrap; }}

  /* 期待度バッジ */
  .exp-badge {{ display: inline-block; padding: 2px 8px; border-radius: 8px; background: #f1f5f9; color: #475569; font-size: 12px; font-weight: 500; }}

  /* 役割テーブル */
  .role-name {{ font-weight: 600; }}

  /* データテーブル共通 */
  .data-table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
  .data-table th {{ background: #f1f5f9; padding: 10px 12px; text-align: left; font-weight: 600; position: sticky; top: 0; }}
  .data-table td {{ padding: 10px 12px; border-bottom: 1px solid #f1f5f9; }}
  .data-table tr:hover {{ background: #f8fafc; }}
  .num {{ text-align: right; font-weight: 600; }}

  /* フィルタ */
  .filters {{ display: flex; gap: 12px; margin-bottom: 16px; flex-wrap: wrap; align-items: center; }}
  .filters label {{ font-size: 13px; font-weight: 600; color: var(--gray); }}
  .filters select {{ padding: 6px 12px; border: 1px solid var(--border); border-radius: 8px; font-size: 13px; font-family: inherit; background: #fff; }}
  .filter-count {{ font-size: 13px; color: var(--gray); margin-left: auto; }}

  /* 応募者リンク */
  .applicant-link {{ color: var(--teal); text-decoration: none; font-weight: 500; }}
  .applicant-link:hover {{ text-decoration: underline; }}

  /* 備考 */
  .note-text {{ cursor: help; border-bottom: 1px dotted var(--gray); }}
  .note-cell {{ max-width: 300px; white-space: normal; word-break: break-word; font-size: 12px; }}

  /* 面談セクション */
  .interview-section {{ border-left: 4px solid #2563eb; }}

  /* スクロール */
  .table-scroll {{ max-height: 60vh; overflow-y: auto; overflow-x: auto; border: 1px solid var(--border); border-radius: 8px; }}

  /* レスポンシブ */
  @media (max-width: 768px) {{
    body {{ padding: 12px; }}
    .summary-cards {{ gap: 8px; }}
    .summary-card {{ padding: 14px 16px; min-width: 100px; }}
    .summary-num {{ font-size: 24px; }}
    .pipeline-label {{ width: 90px; font-size: 12px; }}
    .filters {{ flex-direction: column; gap: 8px; }}
    .table-scroll {{ max-height: 50vh; }}
  }}
</style>
</head>
<body>
<div class="container">
  <header class="header">
    <h1>採用管理ダッシュボード</h1>
    <p>応募者のステータス・役割別状況を一覧表示</p>
    <div class="header-links">
    </div>
  </header>

  <!-- セクション A: サマリーカード -->
  {summary_html}

  <!-- セクション B: ステータス別パイプライン -->
  <section class="sec">
    <h2 class="sec-title">ステータス別パイプライン</h2>
    {pipeline_html}
  </section>

  <!-- セクション C: 役割別の応募状況 -->
  <section class="sec">
    <h2 class="sec-title">役割別の応募状況</h2>
    <table class="data-table">
      <thead><tr><th>役割</th><th>応募者数</th><th>ステータス分布</th></tr></thead>
      <tbody>
        {role_table_html}
      </tbody>
    </table>
  </section>

  <!-- セクション D: 応募者一覧テーブル -->
  <section class="sec">
    <h2 class="sec-title">応募者一覧</h2>
    <div class="filters">
      <label>役割</label>
      <select id="filter-role" onchange="applyFilters()">
        <option value="">すべて</option>
        {role_options}
      </select>
      <label>ステータス</label>
      <select id="filter-status" onchange="applyFilters()">
        <option value="">すべて</option>
        {status_options}
      </select>
      <label>期待度</label>
      <select id="filter-exp" onchange="applyFilters()">
        <option value="">すべて</option>
        {exp_options}
      </select>
      <span class="filter-count" id="filter-count">{len(rows)}件表示</span>
    </div>
    <div class="table-scroll">
      <table class="data-table" id="applicant-table">
        <thead><tr><th>名前</th><th>役割</th><th>期待度</th><th>ステータス</th><th>備考</th><th>面談日</th></tr></thead>
        <tbody>
          {applicant_table_html}
        </tbody>
      </table>
    </div>
  </section>

  <!-- セクション E: 面談予定リスト -->
  <section class="sec interview-section">
    <h2 class="sec-title">面談予定リスト</h2>
    <table class="data-table">
      <thead><tr><th>面談日</th><th>名前</th><th>役割</th><th>期待度</th><th>備考</th></tr></thead>
      <tbody>
        {interview_table_html}
      </tbody>
    </table>
  </section>
</div>

<script>
  window.RECRUITMENT_DATA = {data_js};

  function applyFilters() {{
    var role = document.getElementById('filter-role').value;
    var status = document.getElementById('filter-status').value;
    var exp = document.getElementById('filter-exp').value;
    var rows = document.querySelectorAll('#applicant-table tbody tr');
    var count = 0;
    rows.forEach(function(tr) {{
      var show = true;
      if (role && tr.dataset.role !== role) show = false;
      if (status && tr.dataset.status !== status) show = false;
      if (exp && tr.dataset.exp !== exp) show = false;
      tr.style.display = show ? '' : 'none';
      if (show) count++;
    }});
    document.getElementById('filter-count').textContent = count + '件表示';
  }}
</script>
</body>
</html>'''


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
    print(f"  Total applicants: {agg['total']}")
    print(f"  仮採用: {agg['kari_saiyo']}, 面談予定: {agg['mendan_yotei']}, 未対応: {agg['mitaiou']}")


if __name__ == '__main__':
    main()
