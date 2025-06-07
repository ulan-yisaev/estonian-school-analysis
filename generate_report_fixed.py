#!/usr/bin/env python3

import pandas as pd
import unicodedata
import os
import shutil
from datetime import datetime

def normalize_school_name(name):
    """Normalize school name for comparison"""
    if pd.isna(name) or not isinstance(name, str):
        return ""
    normalized = unicodedata.normalize('NFC', name.strip())
    return normalized.replace('­', '')

def parse_decimal_comma(value):
    """Parse Estonian decimal comma format"""
    if pd.isna(value) or value == "" or value == "—":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).replace(',', '.'))
    except:
        return None

def format_score(score):
    """Format score with Estonian comma"""
    if score is None:
        return "—"
    return f"{score:.1f}".replace('.', ',')

def backup_previous_report(report_filename):
    """Backup previous HTML report with timestamp"""
    if os.path.exists(report_filename):
        # Create backup directory if it doesn't exist
        backup_dir = 'report_backups'
        if not os.path.exists(backup_dir):
            os.makedirs(backup_dir)
            print(f"📁 Created backup directory: {backup_dir}")
        
        # Get current timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Create backup filename
        name, ext = os.path.splitext(report_filename)
        backup_filename = f"{name}_{timestamp}{ext}"
        backup_path = os.path.join(backup_dir, backup_filename)
        
        # Move the file to backup
        shutil.move(report_filename, backup_path)
        print(f"📦 Backed up previous report: {backup_path}")
        return backup_path
    else:
        print("📝 No previous report found - creating new one")
        return None

def normalize(text: str) -> str:
    """NFC normalize, strip extra whitespace and lowercase"""
    text = unicodedata.normalize('NFKC', text)
    return text.strip().lower()

# Expand this list whenever you add new "odd‐ball" Tallinn schools
EXTRA_TALLINN_SCHOOLS = {
    'gustav adolfi gümnaasium',
    'jakob westholmi gümnaasium',
    'vanalinna hariduskolleegium',
    'kadrioru saksa gümnaasium',
    'rocca al mare kool',
    'pelgulinna gümnaasium',
    'pirita majandusgümnaasium',
    'ebs gümnaasium',
    'audentese spordigümnaasium'
}

# Private schools in Tallinn (not suitable for public school selection)
PRIVATE_SCHOOLS = {
    'ebs gümnaasium',
    'rocca al mare kool',
    'audentese spordigümnaasium'
}

def is_tallinn_school(school_name: str) -> bool:
    """
    Returns True if the given school_name is located in Tallinn.
    1) Any name containing 'tallinn'
    2) Any of the known exception names in EXTRA_TALLINN_SCHOOLS
    """
    if not school_name:
        return False

    name = normalize(school_name)

    # 1) names containing the string 'tallinn'
    if 'tallinn' in name:
        return True

    # 2) exact-match against our expanded exception set
    #    (using substring match to catch variants)
    for exc in EXTRA_TALLINN_SCHOOLS:
        if exc in name:
            return True

    return False

def is_private_school(school_name: str) -> bool:
    """
    Returns True if the given school_name is a private school.
    """
    if not school_name:
        return False
    
    name = normalize(school_name)
    
    for private_school in PRIVATE_SCHOOLS:
        if private_school in name:
            return True
    
    return False

def read_csv_data(filename, year):
    """Read and process CSV data for a specific year"""
    print(f"Reading {filename}...")
    df = pd.read_csv(filename)
    
    data = {}
    for _, row in df.iterrows():
        school = normalize_school_name(row.get('Kool', ''))
        if school and is_tallinn_school(school):  # Filter for Tallinn schools
            data[school] = {
                'Place': row.get('Place'),
                'Kokku': parse_decimal_comma(row.get('Kokku')),
                'Math': parse_decimal_comma(row.get('Matemaatika')),
                'Estonian': parse_decimal_comma(row.get('Eesti keel')),
                'English': parse_decimal_comma(row.get('Inglise keel'))
            }
    
    print(f"Found {len(data)} Tallinn schools in {year}")
    return data

# Load all original CSV data
print("Loading original CSV data files...")
data_2018 = read_csv_data('Edetabel 2018.csv', '2018')
data_2023 = read_csv_data('Edetabel 2023.csv', '2023')
data_2024 = read_csv_data('Edetabel 2024.csv', '2024')

# Combine all schools (union of all years)
all_schools = set()
all_schools.update(data_2018.keys())
all_schools.update(data_2023.keys())
all_schools.update(data_2024.keys())

print(f"Total unique Tallinn schools across all years: {len(all_schools)}")

# Create consolidated data
schools_data = []
for school in all_schools:
    school_2018 = data_2018.get(school, {})
    school_2023 = data_2023.get(school, {})
    school_2024 = data_2024.get(school, {})
    
    # Format places as integers (remove .0)
    def format_place(place):
        if place == '—' or place is None or pd.isna(place):
            return '—'
        try:
            return str(int(float(place)))
        except:
            return str(place)
    
    school_record = {
        'School': school,
        # Use original NATIONAL places from CSV files (among ALL Estonian schools)
        'Place_2018': format_place(school_2018.get('Place', '—')),
        'Place_2023': format_place(school_2023.get('Place', '—')),
        'Place_2024': format_place(school_2024.get('Place', '—')),
        # Use original scores
        'Kokku_2018': school_2018.get('Kokku'),
        'Kokku_2023': school_2023.get('Kokku'),
        'Kokku_2024': school_2024.get('Kokku'),
        'Math_2018': school_2018.get('Math'),
        'Math_2023': school_2023.get('Math'),
        'Math_2024': school_2024.get('Math'),
        'Estonian_2018': school_2018.get('Estonian'),
        'Estonian_2023': school_2023.get('Estonian'),
        'Estonian_2024': school_2024.get('Estonian'),
        'English_2018': school_2018.get('English'),
        'English_2023': school_2023.get('English'),
        'English_2024': school_2024.get('English')
    }
    
    schools_data.append(school_record)

# Sort by 2024 total score (for display order)
schools_data.sort(key=lambda x: x['Kokku_2024'] if x['Kokku_2024'] else -1, reverse=True)

# Calculate trends and categories
for school in schools_data:
    # Score trends
    if school['Kokku_2023'] and school['Kokku_2024']:
        school['trend_1yr'] = school['Kokku_2024'] - school['Kokku_2023']
    else:
        school['trend_1yr'] = None
        
    if school['Kokku_2018'] and school['Kokku_2024']:
        school['trend_6yr'] = school['Kokku_2024'] - school['Kokku_2018']
    else:
        school['trend_6yr'] = None
    
    # Place trends (improvement = lower number)
    def parse_place(place_str):
        """Parse place string to number, return None if not valid"""
        if place_str == '—' or place_str is None:
            return None
        try:
            return int(place_str)
        except (ValueError, TypeError):
            return None
    
    place_2018 = parse_place(school['Place_2018'])
    place_2023 = parse_place(school['Place_2023'])
    place_2024 = parse_place(school['Place_2024'])
    
    if place_2018 and place_2024:
        school['place_change_6yr'] = place_2018 - place_2024  # positive = improvement
    else:
        school['place_change_6yr'] = None
        
    if place_2023 and place_2024:
        school['place_change_1yr'] = place_2023 - place_2024  # positive = improvement
    else:
        school['place_change_1yr'] = None
    
    # Performance category
    kokku = school['Kokku_2024']
    if kokku is None:
        school['category'] = 'no-data'
    elif kokku >= 250:
        school['category'] = 'excellent'
    elif kokku >= 200:
        school['category'] = 'very-good'
    elif kokku >= 150:
        school['category'] = 'good'
    else:
        school['category'] = 'average'
    
    # Check if school is private
    school['is_private'] = is_private_school(school['School'])
    
    # Trend category based on place improvement
    place_trend = school['place_change_1yr'] or 0
    if place_trend > 5:  # improved by more than 5 places
        school['trend_cat'] = 'improving'
        school['trend_arrow'] = '↗️'
    elif place_trend < -5:  # declined by more than 5 places
        school['trend_cat'] = 'declining'
        school['trend_arrow'] = '↘️'
    else:
        school['trend_cat'] = 'stable'
        school['trend_arrow'] = '→'

print(f"Processed {len(schools_data)} schools")

# Calculate statistics
stats = {
    'total_schools': len(schools_data),
    'top_performer': max((s for s in schools_data if s['Kokku_2024']), key=lambda x: x['Kokku_2024'], default={'School': 'N/A', 'Kokku_2024': 0}),
    'excellent_count': len([s for s in schools_data if s['category'] == 'excellent']),
    'improving_count': len([s for s in schools_data if s['trend_cat'] == 'improving'])
}

# Generate HTML with national rankings clearly indicated
html_content = f'''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Tallinn Schools Selection Guide 2024 - National Rankings</title>
    <style>
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 20px;
            background-color: #f5f7fa;
            color: #333;
        }}
        
        .header {{
            text-align: center;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 40px 20px;
            border-radius: 15px;
            margin-bottom: 30px;
            box-shadow: 0 8px 25px rgba(0,0,0,0.15);
        }}
        
        .header h1 {{
            margin: 0;
            font-size: 2.5em;
            font-weight: 300;
        }}
        
        .header p {{
            margin: 10px 0 0 0;
            font-size: 1.1em;
            opacity: 0.9;
        }}
        
        .national-notice {{
            background: #e0f2fe;
            border: 1px solid #0288d1;
            border-radius: 8px;
            padding: 16px;
            margin-bottom: 24px;
            text-align: center;
        }}
        
        .national-notice h4 {{
            color: #0277bd;
            margin-bottom: 8px;
        }}
        
        .summary-cards {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            margin-bottom: 40px;
        }}
        
        .summary-card {{
            background: white;
            padding: 25px;
            border-radius: 12px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
            text-align: center;
            border-left: 4px solid #667eea;
        }}
        
        .summary-card h3 {{
            margin: 0 0 10px 0;
            color: #667eea;
            font-size: 1.3em;
        }}
        
        .summary-card .number {{
            font-size: 2.5em;
            font-weight: bold;
            color: #333;
            margin: 10px 0;
        }}
        
        .recommendations {{
            background: white;
            padding: 30px;
            border-radius: 12px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
            margin-bottom: 30px;
        }}
        
        .recommendations h2 {{
            color: #667eea;
            margin-top: 0;
            border-bottom: 2px solid #eee;
            padding-bottom: 10px;
        }}
        
        .recommendation-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
            gap: 20px;
            margin-top: 20px;
        }}
        
        .recommendation-item {{
            padding: 20px;
            border-radius: 8px;
            border-left: 4px solid;
        }}
        
        .rec-excellent {{ 
            background-color: #f0f9ff; 
            border-left-color: #0ea5e9; 
        }}
        
        .rec-improving {{ 
            background-color: #f0fdf4; 
            border-left-color: #22c55e; 
        }}
        
        .rec-consistent {{ 
            background-color: #fefce8; 
            border-left-color: #eab308; 
        }}
        
        .rec-strategic {{
            background-color: #f8fafc;
            border-left-color: #8b5cf6;
        }}
        
        .competitive-notice {{
            background: #fef3c7;
            border: 1px solid #f59e0b;
            border-radius: 8px;
            padding: 16px;
            margin-bottom: 24px;
        }}
        
        .competitive-notice h4 {{
            color: #d97706;
            margin-bottom: 8px;
        }}
        
        .competitive-schools {{
            font-style: italic;
            color: #92400e;
            font-weight: 500;
        }}
        
        .highly-competitive {{
            opacity: 0.4;
            font-size: 0.9em;
            color: #6b7280 !important;
        }}
        
        .highly-competitive .school-name {{
            color: #9ca3af !important;
        }}
        
        .highly-competitive td {{
            color: #9ca3af !important;
        }}
        
        .private-school {{
            opacity: 0.3;
            font-size: 0.9em;
            color: #d97706 !important;
            background-color: #fef3c7 !important;
        }}
        
        .private-school .school-name {{
            color: #92400e !important;
        }}
        
        .private-school td {{
            color: #92400e !important;
        }}
        
        .schools-table {{
            background: white;
            border-radius: 12px;
            overflow: hidden;
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
            margin-bottom: 30px;
        }}
        
        .table-header {{
            background: #667eea;
            color: white;
            padding: 20px;
            font-size: 1.3em;
            font-weight: 500;
        }}
        
        table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 0.9em;
        }}
        
        th, td {{
            padding: 12px 8px;
            text-align: center;
            border-bottom: 1px solid #eee;
        }}
        
        th {{
            background-color: #f8fafc;
            font-weight: 600;
            color: #4a5568;
            position: sticky;
            top: 0;
        }}
        
        .school-name {{
            text-align: left !important;
            font-weight: 500;
            max-width: 200px;
            padding-left: 15px !important;
        }}
        
        .place-cell {{
            font-weight: bold;
        }}
        
        .score-cell {{
            font-weight: 500;
        }}
        
        /* Performance categories */
        .excellent {{ background-color: #dcfce7; color: #166534; }}
        .very-good {{ background-color: #dbeafe; color: #1e40af; }}
        .good {{ background-color: #fef3c7; color: #92400e; }}
        .average {{ background-color: #fee2e2; color: #991b1b; }}
        .no-data {{ background-color: #f3f4f6; color: #6b7280; }}
        
        /* Trend indicators */
        .trend-arrow {{
            font-weight: bold;
            font-size: 1.2em;
        }}
        
        .trend-up {{ color: #16a34a; }}
        .trend-down {{ color: #dc2626; }}
        .trend-stable {{ color: #6b7280; }}
        
        .responsive-table {{
            overflow-x: auto;
        }}
        
        @media (max-width: 768px) {{
            .header h1 {{ font-size: 2em; }}
            .summary-cards {{ grid-template-columns: 1fr; }}
            .recommendation-grid {{ grid-template-columns: 1fr; }}
        }}
        
        .footer {{
            text-align: center;
            margin-top: 40px;
            padding: 20px;
            color: #6b7280;
            font-size: 0.9em;
        }}
    </style>
</head>
<body>
    <div class="header">
        <h1>🏫 Tallinn Schools Selection Guide 2024</h1>
        <p>National Rankings & Performance Analysis (2018-2024)</p>
    </div>

    <div class="national-notice">
        <h4>📊 National Context</h4>
        <p><strong>All rankings shown are NATIONAL rankings among ALL Estonian schools</strong> (not just Tallinn schools). This gives you the true competitive context and shows where Tallinn schools stand nationally.</p>
    </div>

    <div class="summary-cards">
        <div class="summary-card">
            <h3>Tallinn Schools</h3>
            <div class="number">{stats['total_schools']}</div>
            <p>gymnasiums analyzed</p>
        </div>
        <div class="summary-card">
            <h3>Best in Estonia</h3>
            <div class="number">{format_score(stats['top_performer']['Kokku_2024'])}</div>
            <p>{stats['top_performer']['School']}</p>
        </div>
        <div class="summary-card">
            <h3>Excellence Tier</h3>
            <div class="number">{stats['excellent_count']}</div>
            <p>Schools scoring 250+ points</p>
        </div>
        <div class="summary-card">
            <h3>Improving</h3>
            <div class="number">{stats['improving_count']}</div>
            <p>Schools trending upward</p>
        </div>
    </div>
    
    <div class="recommendations">
        <h2>🎯 Realistic School Selection Recommendations</h2>
        
        <div class="competitive-notice">
            <h4>⚠️ Highly Competitive Schools (Extremely Difficult Admission)</h4>
            <p>The following schools have very high entrance requirements and limited acceptance rates. Consider them as "reach" schools only:</p>
            <p class="competitive-schools">Tallinna Reaalkool, Tallinna Inglise Kolledž, Gustav Adolfi Gümnaasium, Tallinna Prantsuse Lütseum, Tallinna 21. Kool, Kadrioru Saksa Gümnaasium</p>
        </div>
        
        <div style="background: #fff7ed; border: 1px solid #ea580c; border-radius: 8px; padding: 16px; margin-bottom: 24px;">
            <h4 style="color: #c2410c; margin-bottom: 8px;">🏫 Private Schools (Not Included in Public School Selection)</h4>
            <p>The following are private schools with tuition fees and are excluded from public school selection considerations:</p>
            <p style="font-style: italic; color: #9a3412; font-weight: 500;">EBS Gümnaasium, Rocca al Mare Kool, Audentese Spordigümnaasium</p>
        </div>
        
        <div class="recommendation-grid">
            <div class="recommendation-item rec-excellent">
                <h4>🏆 Excellent & Accessible</h4>
                <p><strong>High-quality schools with better admission chances:</strong></p>
                <ul>'''

# Add accessible excellent schools (exclude competitive and private schools)
accessible_excellent = [s for s in schools_data if s['category'] in ['very-good'] and s['Kokku_2024'] and 
                        not any(comp in s['School'] for comp in ['Tallinna Reaalkool', 'Tallinna Inglise Kolledž', 'Gustav Adolfi', 'Tallinna Prantsuse', 'Tallinna 21. Kool', 'Kadrioru Saksa']) and
                        not s.get('is_private', False)]
accessible_excellent.sort(key=lambda x: x['Kokku_2024'], reverse=True)

for school in accessible_excellent[:3]:
    national_rank = school['Place_2024'] if school['Place_2024'] != '—' else 'N/A'
    html_content += f'''
                    <li>{school['School']} - {format_score(school['Kokku_2024'])} points (National #{national_rank})</li>'''

html_content += '''
                </ul>
                <p><em>These schools offer excellent education quality with more realistic admission prospects.</em></p>
            </div>
            
            <div class="recommendation-item rec-improving">
                <h4>📈 Rising Stars (Best Value)</h4>
                <p><strong>Schools showing significant improvement trajectory:</strong></p>
                <ul>'''

# Add improving schools (exclude private schools)
improving_schools = [s for s in schools_data if s['trend_cat'] == 'improving' and s['Kokku_2024'] and not s.get('is_private', False)]
improving_schools.sort(key=lambda x: x['place_change_1yr'] or 0, reverse=True)

for school in improving_schools[:4]:
    place_change = school['place_change_1yr'] or 0
    national_rank = school['Place_2024'] if school['Place_2024'] != '—' else 'N/A'
    html_content += f'''
                    <li>{school['School']} - National #{national_rank} (+{place_change} places improvement)</li>'''

html_content += '''
                </ul>
                <p><em>These schools are rapidly improving and offer great opportunities for growth.</em></p>
            </div>
            
            <div class="recommendation-item rec-consistent">
                <h4>🎯 Solid & Reliable Choices</h4>
                <p><strong>Dependable schools with good performance:</strong></p>
                <ul>'''

# Add solid performing schools (exclude private schools)
solid_schools = [s for s in schools_data if s['category'] in ['good'] and s['Kokku_2024'] and s['trend_cat'] == 'stable' and not s.get('is_private', False)]
solid_schools.sort(key=lambda x: x['Kokku_2024'], reverse=True)

for school in solid_schools[:4]:
    national_rank = school['Place_2024'] if school['Place_2024'] != '—' else 'N/A'
    html_content += f'''
                    <li>{school['School']} - {format_score(school['Kokku_2024'])} points (National #{national_rank})</li>'''

html_content += '''
                </ul>
                <p><em>These schools provide consistent quality education with reasonable admission requirements.</em></p>
            </div>

        </div>
    </div>
    
    <div class="schools-table">
        <div class="table-header">
            📊 Tallinn Schools with National Rankings & Performance Analysis
        </div>
        <div class="responsive-table">
            <table>
                <thead>
                    <tr>
                        <th class="school-name">School Name</th>
                        <th>2018<br>National Rank</th>
                        <th>2023<br>National Rank</th>
                        <th>2024<br>National Rank</th>
                        <th>Trend</th>
                        <th>2024<br>Total</th>
                        <th>2024<br>Math</th>
                        <th>2024<br>Estonian</th>
                        <th>2024<br>English</th>
                        <th>Performance</th>
                    </tr>
                </thead>
                <tbody>'''

# Generate table rows
highly_competitive = ['Tallinna Reaalkool', 'Tallinna Inglise Kolledž', 'Gustav Adolfi Gümnaasium', 
                     'Tallinna Prantsuse Lütseum', 'Tallinna 21. Kool', 'Kadrioru Saksa Gümnaasium']

for school in schools_data:
    is_competitive = any(comp in school['School'] for comp in highly_competitive)
    is_private = school.get('is_private', False)
    
    # Determine row class and styling
    if is_private:
        row_class = ' class="private-school"'
        school_name = school['School'] + ' 💼'
        performance = 'Private School'
    elif is_competitive:
        row_class = ' class="highly-competitive"'
        school_name = school['School'] + ' ⭐'
        performance = 'Highly Competitive'
    else:
        row_class = ''
        school_name = school['School']
        performance = school['category'].title().replace('-', ' ')
    
    # Format trend display
    trend_class = f"trend-{school['trend_cat']}"
    trend_arrow = school.get('trend_arrow', '→')
    
    html_content += f'''
                    <tr{row_class}>
                        <td class="school-name">{school_name}</td>
                        <td class="place-cell">{school['Place_2018'] if school['Place_2018'] != '—' else '—'}</td>
                        <td class="place-cell">{school['Place_2023'] if school['Place_2023'] != '—' else '—'}</td>
                        <td class="place-cell">{school['Place_2024'] if school['Place_2024'] != '—' else '—'}</td>
                        <td class="{trend_class}"><span class="trend-arrow">{trend_arrow}</span></td>
                        <td class="score-cell {school['category']}">{format_score(school['Kokku_2024'])}</td>
                        <td class="score-cell">{format_score(school['Math_2024'])}</td>
                        <td class="score-cell">{format_score(school['Estonian_2024'])}</td>
                        <td class="score-cell">{format_score(school['English_2024'])}</td>
                        <td class="{school['category']}">{performance}</td>
                    </tr>'''

html_content += '''
                </tbody>
            </table>
        </div>
    </div>
    
    <div class="application-strategy" style="background: #f8fafc; border-radius: 12px; padding: 24px; margin: 30px 0; border-left: 4px solid #0ea5e9;">
        <h2 style="color: #0f172a; margin-bottom: 16px;">📋 Practical Application Strategy</h2>
        
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 20px;">
            <div style="background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
                <h4 style="color: #16a34a; margin-bottom: 12px;">🎯 Recommended Application List</h4>
                <ul style="margin: 8px 0; padding-left: 20px;">'''

# Add top realistic recommendations with national rankings
accessible_schools = [s for s in schools_data if s['category'] in ['very-good', 'good'] and 
                     not any(comp in s['School'] for comp in highly_competitive) and s['Kokku_2024'] and
                     not s.get('is_private', False)]
accessible_schools.sort(key=lambda x: x['Kokku_2024'], reverse=True)

if accessible_schools:
    top_school = accessible_schools[0]
    national_rank = top_school['Place_2024'] if top_school['Place_2024'] != '—' else 'N/A'
    html_content += f'''
                    <li><strong>Top Choice:</strong> {top_school['School']} (National #{national_rank})</li>'''
    
    if len(accessible_schools) > 2:
        school1, school2 = accessible_schools[1], accessible_schools[2]
        rank1 = school1['Place_2024'] if school1['Place_2024'] != '—' else 'N/A'
        rank2 = school2['Place_2024'] if school2['Place_2024'] != '—' else 'N/A'
        html_content += f'''
                    <li><strong>Solid Options:</strong> {school1['School']} (#{rank1}), {school2['School']} (#{rank2})</li>'''
    
    improving_accessible = [s for s in improving_schools if not any(comp in s['School'] for comp in highly_competitive) and not s.get('is_private', False)][:2]
    if improving_accessible:
        improving_text = ', '.join([f"{s['School']} (#{s['Place_2024'] if s['Place_2024'] != '—' else 'N/A'})" for s in improving_accessible])
        html_content += f'''
                    <li><strong>Rising Stars:</strong> {improving_text}</li>'''

html_content += '''
                </ul>
            </div>
            
            <div style="background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
                <h4 style="color: #dc2626; margin-bottom: 12px;">⚡ Understanding National Rankings</h4>
                <ul style="margin: 8px 0; padding-left: 20px;">
                    <li><strong>Rankings 1-10:</strong> Elite national level - extremely competitive</li>
                    <li><strong>Rankings 11-25:</strong> Excellent schools - very competitive</li>
                    <li><strong>Rankings 26-50:</strong> Very good schools - competitive</li>
                    <li><strong>Rankings 51+:</strong> Good schools - more accessible</li>
                    <li><strong>Consider national context:</strong> Even rank 50+ represents top third nationally</li>
                </ul>
            </div>

        </div>
        
        <div style="background: #e0f2fe; padding: 16px; border-radius: 8px; margin-top: 20px; text-align: center;">
            <p style="color: #0277bd; font-weight: 500; margin: 0;">
                💡 <strong>Pro Tip:</strong> Rankings show national competition level. A school ranked 30th nationally is still in the top 25% of all Estonian schools. Apply to schools across different ranking tiers for the best chances.
            </p>
        </div>
    </div>
    
    <div class="footer">
        <p>📈 Data Analysis Period: 2018-2024 | 🎯 Focus: National Rankings & Realistic School Selection</p>
        <p>Generated from original CSV sources | ⭐ = Highly Competitive Admission | All rankings are NATIONAL</p>
        <div style="background: #fef9c3; border: 1px solid #facc15; border-radius: 8px; padding: 16px; margin-top: 20px;">
            <p style="color: #854d0e; font-weight: 500; margin: 0; text-align: center;">
                ⚠️ <strong>Important Note:</strong> These rankings are based on State Exam Results in Gymnasiums and do not reflect the true quality of education. Each year, schools may have students with different levels of preparation, talents, and backgrounds, which significantly affects these results.
            </p>
        </div>
        <div style="background: #f0f9ff; border: 1px solid #0ea5e9; border-radius: 8px; padding: 16px; margin-top: 16px;">
            <p style="color: #0c4a6e; font-size: 0.9em; margin: 0; text-align: center;">
                📊 <strong>Data Sources:</strong> All data are based on open sources from Postimees:<br>
                <a href="https://rus.postimees.ee/6465418/luchshaya-estonskaya-shkola-nahoditsya-v-tallinne" style="color: #0369a1; text-decoration: none;">2018 School Rankings</a> • 
                <a href="https://rus.postimees.ee/7953716/bolshoy-obzor-rezultaty-gosekzamenov-v-gimnaziyah-kto-vozglavlyaet-reyting-russkoyazychnyh-shkol-estonii" style="color: #0369a1; text-decoration: none;">2024 State Exam Results</a> • 
                <a href="https://services.postimees.ee/infography/2025/2507-koolide-edetabel/eestiTOP/index.html?id=eestiTOP" style="color: #0369a1; text-decoration: none;">Interactive Rankings</a>
            </p>
        </div>
    </div>
    
    <script>
        // Add some interactivity
        document.addEventListener('DOMContentLoaded', function() {{
            // Highlight row on hover
            const rows = document.querySelectorAll('tbody tr');
            rows.forEach(row => {{
                row.addEventListener('mouseenter', function() {{
                    this.style.backgroundColor = '#f8fafc';
                }});
                row.addEventListener('mouseleave', function() {{
                    this.style.backgroundColor = '';
                }});
            }});
        }});
    </script>
</body>
</html>'''

# Backup previous report before creating new one
report_filename = 'school_selection_report_corrected.html'
backup_previous_report(report_filename)

# Write the new HTML file
with open(report_filename, 'w', encoding='utf-8') as f:
    f.write(html_content)

print("✅ Enhanced report generated: school_selection_report_corrected.html")
print("✅ Data sourced directly from original CSV files")
print("✅ Preserves NATIONAL rankings (not just Tallinn rankings)")
print(f"✅ {len(schools_data)} schools processed with accurate national context")
print("✅ Clear indication that rankings are national, not local") 