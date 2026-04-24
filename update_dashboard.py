#!/usr/bin/env python3
"""
Fetch quiz data from Google Sheets and generate the dashboard HTML.
Supports both service account authentication and public CSV export.
"""

import requests
import csv
import io
import re
import os
from pathlib import Path

# Try to import gspread for service account authentication
try:
    import gspread
    from google.oauth2.service_account import Credentials
    GSPREAD_AVAILABLE = True
except ImportError:
    GSPREAD_AVAILABLE = False

SHEET_URL = "https://docs.google.com/spreadsheets/d/1w2CPrgmlVu4CmxlQY-T9J6g5NcNLKgx3eg4wMlZdAhI/edit?gid=0#gid=0"
SHEET_ID = "1w2CPrgmlVu4CmxlQY-T9J6g5NcNLKgx3eg4wMlZdAhI"

def extract_sheet_id(url):
    """Extract sheet ID from Google Sheets URL"""
    match = re.search(r'/spreadsheets/d/([a-zA-Z0-9-_]+)', url)
    return match.group(1) if match else None

def fetch_sheet_data_gspread(sheet_id):
    """
    Fetch data from Google Sheet using gspread and service account.
    Requires sheets-key.json file.
    """
    if not GSPREAD_AVAILABLE:
        return None
    
    try:
        # Look for sheets-key.json in current directory or via environment variable
        key_file = os.getenv('SHEETS_KEY_FILE', 'sheets-key.json')
        
        if not Path(key_file).exists():
            return None
        
        scope = [
            'https://www.googleapis.com/auth/spreadsheets.readonly',
            'https://www.googleapis.com/auth/drive.readonly'
        ]
        
        creds = Credentials.from_service_account_file(key_file, scopes=scope)
        client = gspread.authorize(creds)
        spreadsheet = client.open_by_key(sheet_id)
        worksheet = spreadsheet.worksheet("Sheet1")
        
        # Get all records
        records = worksheet.get_all_records()
        print(f"✅ Connected via service account")
        return records
    
    except Exception as e:
        print(f"⚠️  Service account auth failed: {e}")
        return None

def fetch_sheet_data_csv(sheet_url):
    """
    Fetch data from public Google Sheet via CSV export.
    Sheet must be shared with "Anyone with the link" (Viewer access).
    """
    try:
        sheet_id = extract_sheet_id(sheet_url)
        if not sheet_id:
            print("❌ Could not extract sheet ID from URL")
            return None
        
        # Construct CSV export URL
        csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid=0"
        print(f"📊 Fetching via public CSV export...")
        
        response = requests.get(csv_url, timeout=10)
        if response.status_code != 200:
            print(f"❌ Failed to fetch sheet (Status: {response.status_code})")
            return None
        
        # Parse CSV
        csv_file = io.StringIO(response.text)
        reader = csv.DictReader(csv_file)
        records = list(reader)
        return records
    
    except requests.exceptions.RequestException as e:
        print(f"❌ Network error: {e}")
        return None
    except Exception as e:
        print(f"❌ Error fetching data: {e}")
        return None

def fetch_sheet_data(sheet_url, sheet_id):
    """
    Fetch sheet data - tries service account first, falls back to public CSV.
    """
    # Try service account first
    data = fetch_sheet_data_gspread(sheet_id)
    if data:
        return data
    
    # Fall back to public CSV
    data = fetch_sheet_data_csv(sheet_url)
    if data:
        return data
    
    print("❌ Could not fetch data via service account or public CSV")
    return None

def generate_dashboard_html(quiz_data):
    """Generate HTML dashboard from quiz data"""
    
    if not quiz_data:
        return None
    
    # Parse the data
    quiz_rows = [row for row in quiz_data if row.get('Set No.') and row.get('Topic')]
    players = ['Jameer', 'Akhil', 'Sreehari', 'Ameen', 'Nadeem']
    team_map = {
        "Victoria's Secrets": ['Jameer', 'Akhil'],
        'Khiladis of Kerala': ['Sreehari', 'Ameen', 'Nadeem']
    }
    
    # Calculate totals
    player_totals = {}
    for player in players:
        total = sum(float(row.get(player, 0) or 0) for row in quiz_rows if row.get(player))
        player_totals[player] = int(total)
    
    # Sort by total
    sorted_players = sorted(player_totals.items(), key=lambda x: x[1], reverse=True)
    
    # Team totals
    team_totals = {}
    for team, members in team_map.items():
        total = sum(float(row.get(m, 0) or 0) for row in quiz_rows for m in members if row.get(m))
        team_totals[team] = int(total)
    
    # Winners per quiz
    winners = []
    for row in quiz_rows:
        scores = {p: float(row.get(p, 0) or 0) for p in players}
        max_score = max(scores.values())
        top = [p for p in players if scores[p] == max_score]
        winners.append({
            'set': int(float(row['Set No.'])),
            'topic': row['Topic'],
            'top_score': int(max_score),
            'winners': top,
            'joint': len(top) > 1
        })
    
    # Win counts
    win_counts = {}
    for w in winners:
        for p in w['winners']:
            win_counts[p] = win_counts.get(p, 0) + 1
    
    # Build HTML table rows for full results
    results_rows_html = ""
    for row in quiz_rows:
        set_no = int(float(row['Set No.']))
        topic = row['Topic']
        scores = {p: float(row.get(p, 0) or 0) for p in players}
        max_score = max(scores.values())
        
        results_rows_html += f"""            <tr>
              <td>{set_no}</td>
              <td>{topic}</td>"""
        for p in players:
            score = int(scores[p])
            is_top = score == max_score
            class_attr = ' class="top-score"' if is_top else ''
            results_rows_html += f"<td{class_attr}>{score}</td>"
        results_rows_html += f"""
              <td>{int(float(row.get('Maximum Score', 0) or 0))}</td>
            </tr>
"""
    
    # Build winners table
    winners_table_html = ""
    for w in winners:
        winners_str = " &amp; ".join(w['winners']) if w['joint'] else w['winners'][0]
        winners_table_html += f"""            <tr>
              <td>{w['set']}</td>
              <td>{w['topic']}</td>
              <td>{winners_str}</td>
              <td>{w['top_score']}</td>
            </tr>
"""
    
    # Build HTML
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Vahid's Written Quiz Season 1 Dashboard</title>
  <style>
    body {{
      margin: 0;
      font-family: Inter, system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
      background: #0f172a;
      color: #e2e8f0;
    }}
    .page {{
      max-width: 1200px;
      margin: 0 auto;
      padding: 32px;
    }}
    h1, h2, h3 {{
      margin: 0;
      color: #f8fafc;
    }}
    header {{
      padding-bottom: 24px;
      border-bottom: 1px solid rgba(148, 163, 184, 0.24);
      margin-bottom: 32px;
    }}
    .subheading {{
      margin-top: 8px;
      color: #94a3b8;
    }}
    .grid {{
      display: grid;
      gap: 20px;
      grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
    }}
    .card {{
      background: rgba(15, 23, 42, 0.88);
      border: 1px solid rgba(148, 163, 184, 0.12);
      border-radius: 18px;
      padding: 22px;
      box-shadow: 0 20px 45px rgba(15, 23, 42, 0.28);
    }}
    .metric {{
      font-size: 3rem;
      line-height: 1;
      margin: 0;
      color: #f8fafc;
    }}
    .metric-label {{
      color: #94a3b8;
      margin-top: 8px;
      font-size: 0.95rem;
    }}
    .leaderboard-table,
    .results-table,
    .wins-table {{
      width: 100%;
      border-collapse: collapse;
      margin-top: 16px;
      font-size: 0.95rem;
      color: #e2e8f0;
    }}
    .leaderboard-table th,
    .leaderboard-table td,
    .results-table th,
    .results-table td,
    .wins-table th,
    .wins-table td {{
      padding: 12px 14px;
      border-bottom: 1px solid rgba(148, 163, 184, 0.12);
      text-align: left;
    }}
    .leaderboard-table th,
    .results-table th,
    .wins-table th {{
      background: rgba(15, 23, 42, 0.95);
      color: #cbd5e1;
      font-weight: 600;
    }}
    .top-score {{
      background: rgba(16, 185, 129, 0.18);
      color: #a7f3d0;
      font-weight: 700;
    }}
    .section {{
      margin-bottom: 40px;
    }}
    .section-header {{
      display: flex;
      align-items: flex-end;
      justify-content: space-between;
      gap: 18px;
    }}
    .section-header h2 {{
      margin-bottom: 8px;
    }}
    .note {{
      color: #94a3b8;
      margin-top: 10px;
      font-size: 0.95rem;
      line-height: 1.6;
    }}
    .callout {{
      margin-top: 20px;
      padding: 18px 20px;
      border-radius: 16px;
      background: rgba(59, 130, 246, 0.08);
      border: 1px solid rgba(59, 130, 246, 0.18);
      color: #e2e8f0;
    }}
    .small-text {{
      color: #cbd5e1;
      font-size: 0.9rem;
    }}
    .table-title {{
      margin-top: 0;
      margin-bottom: 12px;
      color: #e2e8f0;
    }}
    .results-table tbody tr:nth-child(odd) {{
      background: rgba(148, 163, 184, 0.03);
    }}
    .results-table tbody tr:hover {{
      background: rgba(100, 116, 139, 0.16);
    }}
    .updated-at {{
      text-align: center;
      color: #64748b;
      font-size: 0.85rem;
      margin-top: 40px;
      padding-top: 20px;
      border-top: 1px solid rgba(148, 163, 184, 0.12);
    }}
  </style>
</head>
<body>
  <div class="page">
    <header>
      <h1>Vahid's Written Quiz Season 1</h1>
      <p class="subheading">Drivers' Championship, Constructors' Championship, Grand Prix wins, and Full Quiz Results.</p>
    </header>

    <section class="section">
      <div class="section-header">
        <div>
          <h2>Drivers' Championship Leaderboard</h2>
          <p class="note">Ranks based on total individual scores across all quiz rounds.</p>
        </div>
      </div>
      <div class="card">
        <table class="leaderboard-table">
          <thead>
            <tr>
              <th>Rank</th>
              <th>Driver</th>
              <th>Total Score</th>
            </tr>
          </thead>
          <tbody>
"""
    
    for rank, (player, total) in enumerate(sorted_players, 1):
        is_top = rank == 1
        class_attr = ' class="top-score"' if is_top else ''
        html += f"""            <tr>
              <td>{rank}</td>
              <td>{player}</td>
              <td{class_attr}>{total}</td>
            </tr>
"""
    
    html += """          </tbody>
        </table>
      </div>
    </section>

    <section class="section">
      <div class="section-header">
        <div>
          <h2>Constructors' Championship Results</h2>
          <p class="note">Team totals are the sum of their members' scores over all quiz rounds.</p>
        </div>
      </div>
      <div class="grid">
"""
    
    sorted_teams = sorted(team_totals.items(), key=lambda x: x[1], reverse=True)
    for team, total in sorted_teams:
        html += f"""        <div class="card">
          <p class="metric">{total}</p>
          <p class="metric-label">{team}</p>
        </div>
"""
    
    winner_team = sorted_teams[0][0]
    winner_score = sorted_teams[0][1]
    runner_up_score = sorted_teams[1][1]
    
    html += f"""      </div>
      <div class="callout">
        <strong>Constructors' Leader:</strong> {winner_team} with a {winner_score}–{runner_up_score} lead.
      </div>
    </section>

    <section class="section">
      <div class="section-header">
        <div>
          <h2>Grand Prix Wins</h2>
          <p class="note">Joint victories are included.</p>
        </div>
      </div>
      <div class="card">
        <table class="wins-table">
          <thead>
            <tr>
              <th>Driver</th>
              <th>Grand Prix Wins</th>
            </tr>
          </thead>
          <tbody>
"""
    
    sorted_wins = sorted(win_counts.items(), key=lambda x: x[1], reverse=True)
    for player in players:
        wins = win_counts.get(player, 0)
        is_top = wins == sorted_wins[0][1] if sorted_wins else False
        class_attr = ' class="top-score"' if is_top else ''
        html += f"""            <tr>
              <td>{player}</td>
              <td{class_attr}>{wins}</td>
            </tr>
"""
    
    html += f"""          </tbody>
        </table>
        <div class="note">Joint winners occurred in {sum(1 for w in winners if w['joint'])} rounds.</div>
      </div>
      <div class="card">
        <h3 class="table-title">Grand Prix Winners by Round</h3>
        <table class="results-table">
          <thead>
            <tr>
              <th>Set</th>
              <th>Topic</th>
              <th>Winner(s)</th>
              <th>Top Score</th>
            </tr>
          </thead>
          <tbody>
{winners_table_html}          </tbody>
        </table>
      </div>
    </section>

    <section class="section">
      <div class="section-header">
        <div>
          <h2>Full Quiz Results</h2>
          <p class="note">Each round includes the score of every driver. Top scorers are highlighted.</p>
        </div>
      </div>
      <div class="card">
        <table class="results-table">
          <thead>
            <tr>
              <th>Set</th>
              <th>Topic</th>
              <th>Jameer</th>
              <th>Akhil</th>
              <th>Sreehari</th>
              <th>Ameen</th>
              <th>Nadeem</th>
              <th>Max Score</th>
            </tr>
          </thead>
          <tbody>
{results_rows_html}          </tbody>
        </table>
      </div>
    </section>

    <div class="updated-at">
      Dashboard auto-generated from Google Sheets
    </div>
  </div>
</body>
</html>
"""
    return html

if __name__ == "__main__":
    print("🔄 Fetching quiz data from Google Sheets...")
    quiz_data = fetch_sheet_data(SHEET_URL, SHEET_ID)
    
    if quiz_data:
        print(f"✅ Fetched {len(quiz_data)} rows")
        html = generate_dashboard_html(quiz_data)
        
        if html:
            output_file = "written_quiz_dashboard.html"
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(html)
            print(f"✅ Dashboard saved to {output_file}")
            print(f"📤 Push this to GitHub Pages to make it live!")
    else:
        print("❌ Failed to fetch data. Check authentication.")
