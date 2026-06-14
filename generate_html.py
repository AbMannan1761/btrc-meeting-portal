import json
import os

db_path = r"C:\Users\mannan\.gemini\antigravity\brain\d4fb72b8-9e63-42ba-8680-37ffcb1867fa\scratch\meetings_db.json"
output_html_path = r"c:\Users\mannan\OneDrive - Bangladesh Telecommunication Regulatory Commission\1. E & I\Meeting\BTRC_Meeting_Minutes_Reader.html"

# Load the database
with open(db_path, "r", encoding="utf-8") as f:
    master_db = json.load(f)

# HTML/CSS/JS template content
html_template = """<!DOCTYPE html>
<html lang="bn">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>BTRC Commission Meeting Minutes & Decisions Portal</title>
    <!-- Google Fonts -->
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Outfit:wght@300;400;500;600;700;800&family=Noto+Sans+Bengali:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    
    <style>
        :root {
            --bg-primary: #0b111e;
            --bg-secondary: #131c2e;
            --bg-card: rgba(20, 30, 50, 0.65);
            --bg-glass: rgba(15, 23, 42, 0.45);
            --border-glass: rgba(255, 255, 255, 0.07);
            --text-primary: #f3f4f6;
            --text-secondary: #9ca3af;
            --text-muted: #6b7280;
            --accent-blue: #3b82f6;
            --accent-green: #10b981;
            --accent-amber: #f59e0b;
            --accent-red: #f43f5e;
            --accent-purple: #8b5cf6;
            --shadow-glass: 0 8px 32px 0 rgba(0, 0, 0, 0.37);
            --transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        }

        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            background-color: var(--bg-primary);
            color: var(--text-primary);
            font-family: 'Inter', 'Noto Sans Bengali', sans-serif;
            background-image: 
                radial-gradient(at 10% 20%, rgba(59, 130, 246, 0.08) 0px, transparent 50%),
                radial-gradient(at 90% 10%, rgba(139, 92, 246, 0.08) 0px, transparent 50%),
                radial-gradient(at 50% 80%, rgba(16, 185, 129, 0.05) 0px, transparent 50%);
            background-attachment: fixed;
            min-height: 100vh;
            line-height: 1.6;
        }

        header {
            background: rgba(11, 17, 30, 0.8);
            backdrop-filter: blur(12px);
            border-bottom: 1px solid var(--border-glass);
            padding: 1.25rem 2rem;
            position: sticky;
            top: 0;
            z-index: 100;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }

        .header-title-container {
            display: flex;
            align-items: center;
            gap: 1rem;
        }

        .logo-placeholder {
            width: 48px;
            height: 48px;
            background: linear-gradient(135deg, var(--accent-blue), var(--accent-purple));
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-family: 'Outfit', sans-serif;
            font-weight: 800;
            font-size: 1.25rem;
            box-shadow: 0 0 15px rgba(59, 130, 246, 0.4);
            border: 2px solid rgba(255, 255, 255, 0.1);
        }

        .header-title h1 {
            font-family: 'Outfit', 'Noto Sans Bengali', sans-serif;
            font-size: 1.5rem;
            font-weight: 700;
            background: linear-gradient(to right, #ffffff, #9ca3af);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }

        .header-title p {
            font-size: 0.8rem;
            color: var(--text-secondary);
        }

        .tab-switcher {
            display: flex;
            gap: 0.5rem;
            background: rgba(255, 255, 255, 0.03);
            border: 1px solid var(--border-glass);
            padding: 0.35rem;
            border-radius: 30px;
        }

        .tab-btn {
            background: transparent;
            border: none;
            color: var(--text-secondary);
            padding: 0.5rem 1.25rem;
            border-radius: 20px;
            font-weight: 500;
            font-size: 0.9rem;
            cursor: pointer;
            transition: var(--transition);
            font-family: 'Noto Sans Bengali', sans-serif;
        }

        .tab-btn.active {
            background: linear-gradient(135deg, var(--accent-blue), var(--accent-purple));
            color: #ffffff;
            box-shadow: 0 4px 15px rgba(59, 130, 246, 0.3);
        }

        .container {
            max-width: 1400px;
            margin: 0 auto;
            padding: 2rem;
        }

        /* Stats Row */
        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
            gap: 1.5rem;
            margin-bottom: 2.5rem;
        }

        .stat-card {
            background: var(--bg-card);
            backdrop-filter: blur(16px);
            border: 1px solid var(--border-glass);
            border-radius: 16px;
            padding: 1.5rem;
            box-shadow: var(--shadow-glass);
            transition: var(--transition);
            position: relative;
            overflow: hidden;
        }

        .stat-card::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            width: 4px;
            height: 100%;
            background: var(--accent-blue);
        }

        .stat-card.green::before { background: var(--accent-green); }
        .stat-card.amber::before { background: var(--accent-amber); }
        .stat-card.red::before { background: var(--accent-red); }

        .stat-card:hover {
            transform: translateY(-5px);
            border-color: rgba(255, 255, 255, 0.15);
        }

        .stat-card h3 {
            font-size: 0.85rem;
            color: var(--text-secondary);
            font-family: 'Noto Sans Bengali', sans-serif;
            margin-bottom: 0.5rem;
        }

        .stat-value {
            font-size: 1.8rem;
            font-weight: 700;
            font-family: 'Outfit', sans-serif;
            color: #ffffff;
        }

        .stat-subtitle {
            font-size: 0.75rem;
            color: var(--text-muted);
            margin-top: 0.25rem;
        }

        /* Search & Filters */
        .search-filters-row {
            background: var(--bg-glass);
            backdrop-filter: blur(10px);
            border: 1px solid var(--border-glass);
            border-radius: 16px;
            padding: 1.25rem;
            margin-bottom: 2.5rem;
            display: flex;
            gap: 1rem;
            flex-wrap: wrap;
            align-items: center;
        }

        .search-container {
            flex: 1;
            min-width: 300px;
            position: relative;
        }

        .search-input {
            width: 100%;
            background: rgba(0, 0, 0, 0.3);
            border: 1px solid var(--border-glass);
            border-radius: 10px;
            padding: 0.75rem 1rem 0.75rem 2.5rem;
            color: var(--text-primary);
            font-size: 0.95rem;
            transition: var(--transition);
            font-family: 'Noto Sans Bengali', 'Inter', sans-serif;
        }

        .search-input:focus {
            outline: none;
            border-color: var(--accent-blue);
            box-shadow: 0 0 10px rgba(59, 130, 246, 0.25);
        }

        .search-icon {
            position: absolute;
            left: 0.85rem;
            top: 50%;
            transform: translateY(-50%);
            color: var(--text-muted);
            pointer-events: none;
        }

        .filter-select {
            background: rgba(0, 0, 0, 0.3);
            border: 1px solid var(--border-glass);
            border-radius: 10px;
            padding: 0.75rem 1.5rem 0.75rem 1rem;
            color: var(--text-primary);
            font-size: 0.9rem;
            cursor: pointer;
            transition: var(--transition);
            font-family: 'Noto Sans Bengali', sans-serif;
            min-width: 180px;
        }

        .filter-select:focus {
            outline: none;
            border-color: var(--accent-blue);
        }

        /* Dashboard Meetings Grid (Layer 1) */
        .meetings-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(280px, 1fr));
            gap: 1.5rem;
        }

        .meeting-card {
            background: var(--bg-card);
            backdrop-filter: blur(16px);
            border: 1px solid var(--border-glass);
            border-radius: 16px;
            padding: 1.5rem;
            box-shadow: var(--shadow-glass);
            cursor: pointer;
            transition: var(--transition);
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            min-height: 200px;
            position: relative;
            overflow: hidden;
        }

        .meeting-card::after {
            content: '';
            position: absolute;
            top: 0;
            right: 0;
            width: 50px;
            height: 50px;
            background: linear-gradient(135deg, transparent, rgba(59, 130, 246, 0.05));
            border-radius: 0 0 0 100%;
            transition: var(--transition);
        }

        .meeting-card:hover {
            transform: translateY(-6px);
            border-color: rgba(59, 130, 246, 0.3);
            box-shadow: 0 12px 30px rgba(59, 130, 246, 0.15);
        }

        .meeting-card:hover::after {
            background: linear-gradient(135deg, transparent, rgba(59, 130, 246, 0.2));
            width: 70px;
            height: 70px;
        }

        .meeting-num-badge {
            align-self: flex-start;
            background: rgba(59, 130, 246, 0.15);
            color: #60a5fa;
            border: 1px solid rgba(59, 130, 246, 0.25);
            font-size: 0.8rem;
            font-weight: 600;
            padding: 0.25rem 0.75rem;
            border-radius: 20px;
            margin-bottom: 1rem;
            font-family: 'Noto Sans Bengali', sans-serif;
        }

        .meeting-card h2 {
            font-family: 'Noto Sans Bengali', sans-serif;
            font-size: 1.25rem;
            font-weight: 600;
            margin-bottom: 0.75rem;
            color: #ffffff;
        }

        .meeting-card .date {
            font-size: 0.85rem;
            color: var(--text-secondary);
            margin-bottom: 1rem;
            display: flex;
            align-items: center;
            gap: 0.35rem;
            font-family: 'Noto Sans Bengali', sans-serif;
        }

        .meeting-info-row {
            display: flex;
            justify-content: space-between;
            font-size: 0.8rem;
            color: var(--text-muted);
            border-top: 1px solid rgba(255, 255, 255, 0.05);
            padding-top: 0.75rem;
            margin-top: auto;
        }

        .meeting-info-item strong {
            color: var(--text-secondary);
        }

        /* 3-Layer Split View (Layer 2 & 3) */
        .viewer-layout {
            display: grid;
            grid-template-columns: 350px 1fr;
            gap: 2rem;
            background: var(--bg-glass);
            border: 1px solid var(--border-glass);
            border-radius: 20px;
            padding: 1.5rem;
            box-shadow: var(--shadow-glass);
            min-height: 70vh;
        }

        .agenda-sidebar {
            border-right: 1px solid var(--border-glass);
            padding-right: 1.5rem;
            display: flex;
            flex-direction: column;
            max-height: 75vh;
            overflow-y: auto;
        }

        .back-btn {
            background: rgba(255, 255, 255, 0.05);
            border: 1px solid var(--border-glass);
            color: var(--text-primary);
            padding: 0.65rem 1.25rem;
            border-radius: 10px;
            cursor: pointer;
            font-weight: 500;
            font-size: 0.9rem;
            display: inline-flex;
            align-items: center;
            gap: 0.5rem;
            margin-bottom: 1.5rem;
            transition: var(--transition);
            font-family: 'Noto Sans Bengali', sans-serif;
        }

        .back-btn:hover {
            background: rgba(255, 255, 255, 0.1);
            border-color: var(--accent-blue);
        }

        .sidebar-meeting-info {
            margin-bottom: 1.5rem;
        }

        .sidebar-meeting-info h2 {
            font-size: 1.35rem;
            font-family: 'Noto Sans Bengali', sans-serif;
        }

        .sidebar-meeting-info p {
            font-size: 0.85rem;
            color: var(--text-secondary);
            font-family: 'Noto Sans Bengali', sans-serif;
        }

        .agenda-list {
            list-style: none;
            display: flex;
            flex-direction: column;
            gap: 0.75rem;
        }

        .agenda-item {
            background: rgba(255, 255, 255, 0.02);
            border: 1px solid var(--border-glass);
            border-radius: 12px;
            padding: 1rem;
            cursor: pointer;
            transition: var(--transition);
        }

        .agenda-item:hover {
            background: rgba(255, 255, 255, 0.05);
            border-color: rgba(59, 130, 246, 0.2);
        }

        .agenda-item.active {
            background: rgba(59, 130, 246, 0.1);
            border-color: rgba(59, 130, 246, 0.5);
            box-shadow: inset 0 0 10px rgba(59, 130, 246, 0.1);
        }

        .agenda-item-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 0.5rem;
        }

        .agenda-badge {
            font-size: 0.7rem;
            background: rgba(255, 255, 255, 0.08);
            padding: 0.15rem 0.5rem;
            border-radius: 4px;
            color: var(--text-secondary);
        }

        .agenda-item.active .agenda-badge {
            background: var(--accent-blue);
            color: #ffffff;
        }

        .agenda-item-title {
            font-size: 0.85rem;
            font-weight: 500;
            color: var(--text-primary);
            display: -webkit-box;
            -webkit-line-clamp: 2;
            -webkit-box-orient: vertical;
            overflow: hidden;
            text-overflow: ellipsis;
            font-family: 'Noto Sans Bengali', sans-serif;
        }

        .detail-viewer {
            padding-left: 0.5rem;
            display: grid;
            grid-template-columns: 1.1fr 1fr;
            gap: 1.5rem;
            align-items: start;
            max-height: 75vh;
            overflow-y: auto;
        }

        .detail-core-column {
            display: flex;
            flex-direction: column;
            gap: 1.5rem;
        }

        /* View Toggle styling */
        .view-toggle-container {
            display: flex;
            gap: 0.5rem;
            margin-left: auto;
        }

        .toggle-view-btn {
            background: rgba(255, 255, 255, 0.03);
            border: 1px solid var(--border-glass);
            color: var(--text-secondary);
            padding: 0.65rem 1.25rem;
            border-radius: 10px;
            font-weight: 500;
            font-size: 0.9rem;
            cursor: pointer;
            transition: var(--transition);
            font-family: 'Noto Sans Bengali', sans-serif;
            display: flex;
            align-items: center;
            gap: 0.5rem;
        }

        .toggle-view-btn:hover {
            background: rgba(255, 255, 255, 0.08);
            border-color: rgba(59, 130, 246, 0.2);
        }

        .toggle-view-btn.active {
            background: rgba(59, 130, 246, 0.15);
            border-color: var(--accent-blue);
            color: #ffffff;
            box-shadow: 0 0 10px rgba(59, 130, 246, 0.15);
        }


        /* Premium Search Panel */
        .search-panel-container {
            background: var(--bg-card);
            backdrop-filter: blur(16px);
            border: 1px solid var(--border-glass);
            border-radius: 16px;
            padding: 1.5rem;
            margin-bottom: 2rem;
            box-shadow: var(--shadow-glass);
            display: flex;
            flex-direction: column;
            gap: 1.25rem;
        }

        .search-panel-header {
            display: flex;
            align-items: center;
            gap: 0.75rem;
            color: var(--accent-blue);
            border-bottom: 1px solid rgba(255, 255, 255, 0.05);
            padding-bottom: 0.75rem;
        }

        .search-panel-header h2 {
            font-size: 1.15rem;
            font-weight: 600;
            color: #ffffff;
            font-family: 'Noto Sans Bengali', sans-serif;
        }

        .search-panel-body {
            display: flex;
            flex-direction: column;
            gap: 1.25rem;
        }

        .main-search-row {
            display: flex;
            gap: 1rem;
        }

        .search-input-premium {
            flex: 1;
            background: rgba(0, 0, 0, 0.4);
            border: 1px solid var(--border-glass);
            border-radius: 10px;
            padding: 0.85rem 1.25rem;
            color: var(--text-primary);
            font-size: 1rem;
            transition: var(--transition);
            font-family: 'Noto Sans Bengali', 'Inter', sans-serif;
        }

        .search-input-premium:focus {
            outline: none;
            border-color: var(--accent-blue);
            box-shadow: 0 0 10px rgba(59, 130, 246, 0.25);
        }

        .search-action-btn {
            background: linear-gradient(135deg, var(--accent-blue), var(--accent-purple));
            color: #ffffff;
            border: none;
            padding: 0.85rem 2rem;
            border-radius: 10px;
            font-weight: 600;
            cursor: pointer;
            transition: var(--transition);
            font-family: 'Noto Sans Bengali', sans-serif;
            box-shadow: 0 4px 15px rgba(59, 130, 246, 0.3);
        }

        .search-action-btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 6px 20px rgba(59, 130, 246, 0.4);
        }

        .filters-row-premium {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 1.25rem;
            align-items: flex-end;
        }

        .filter-group {
            display: flex;
            flex-direction: column;
            gap: 0.5rem;
        }

        .filter-group label {
            font-size: 0.8rem;
            color: var(--text-secondary);
            font-family: 'Noto Sans Bengali', sans-serif;
            font-weight: 500;
        }

        .filter-select-premium {
            background: rgba(0, 0, 0, 0.4);
            border: 1px solid var(--border-glass);
            border-radius: 8px;
            padding: 0.75rem 1rem;
            color: var(--text-primary);
            font-size: 0.9rem;
            cursor: pointer;
            transition: var(--transition);
            font-family: 'Noto Sans Bengali', sans-serif;
            outline: none;
        }

        .filter-select-premium:focus {
            border-color: var(--accent-blue);
        }

        .view-toggle-group {
            align-items: stretch;
        }

        .view-toggle-container-premium {
            display: flex;
            background: rgba(0, 0, 0, 0.3);
            border: 1px solid var(--border-glass);
            padding: 0.25rem;
            border-radius: 8px;
            gap: 0.25rem;
        }

        .toggle-view-btn-premium {
            flex: 1;
            background: transparent;
            border: none;
            color: var(--text-secondary);
            padding: 0.5rem;
            border-radius: 6px;
            font-size: 0.85rem;
            cursor: pointer;
            transition: var(--transition);
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 0.35rem;
            font-family: 'Noto Sans Bengali', sans-serif;
        }

        .toggle-view-btn-premium.active {
            background: rgba(59, 130, 246, 0.15);
            color: #ffffff;
            border: 1px solid rgba(59, 130, 246, 0.25);
        }

        .search-stats-premium {
            font-size: 0.85rem;
            color: var(--text-muted);
            font-family: 'Noto Sans Bengali', sans-serif;
            border-top: 1px solid rgba(255, 255, 255, 0.05);
            padding-top: 0.75rem;
        }

        .search-stats-premium strong {
            color: var(--accent-blue);
        }

        /* Dashboard Content Layout */
        .dashboard-content-wrapper {
            display: grid;
            grid-template-columns: 1fr;
            gap: 2rem;
            transition: var(--transition);
        }

        .dashboard-content-wrapper.list-mode {
            grid-template-columns: 280px 1fr;
        }

        .toc-sidebar {
            background: var(--bg-card);
            backdrop-filter: blur(16px);
            border: 1px solid var(--border-glass);
            border-radius: 16px;
            padding: 1.25rem;
            box-shadow: var(--shadow-glass);
            position: sticky;
            top: 90px;
            max-height: calc(100vh - 130px);
            overflow-y: auto;
            display: flex;
            flex-direction: column;
            gap: 1rem;
        }

        .toc-sidebar.hidden {
            display: none !important;
        }

        .toc-sidebar h3 {
            font-family: 'Noto Sans Bengali', sans-serif;
            font-size: 1rem;
            color: #ffffff;
            border-bottom: 1px solid var(--border-glass);
            padding-bottom: 0.5rem;
        }

        .toc-links {
            display: flex;
            flex-direction: column;
            gap: 0.5rem;
        }

        .toc-link {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 0.5rem 0.75rem;
            border-radius: 8px;
            color: var(--text-secondary);
            text-decoration: none;
            font-size: 0.85rem;
            font-family: 'Noto Sans Bengali', sans-serif;
            transition: var(--transition);
            background: rgba(255, 255, 255, 0.01);
            border: 1px solid transparent;
        }

        .toc-link:hover {
            color: #ffffff;
            background: rgba(255, 255, 255, 0.05);
            border-color: rgba(59, 130, 246, 0.2);
        }

        .toc-link.active {
            color: #ffffff;
            background: rgba(59, 130, 246, 0.15);
            border-color: var(--accent-blue);
        }

        .toc-link-badge {
            background: rgba(255, 255, 255, 0.05);
            padding: 0.1rem 0.4rem;
            border-radius: 4px;
            font-size: 0.75rem;
            color: var(--text-muted);
        }

        .toc-link.active .toc-link-badge {
            background: var(--accent-blue);
            color: #ffffff;
        }

        .meeting-section-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 2px solid rgba(255, 255, 255, 0.07);
            padding-bottom: 0.5rem;
            margin-top: 2rem;
            margin-bottom: 1.25rem;
            scroll-margin-top: 100px;
        }

        .meeting-section-header:first-of-type {
            margin-top: 0;
        }

        .meeting-section-header h2 {
            font-family: 'Noto Sans Bengali', sans-serif;
            font-size: 1.5rem;
            color: #ffffff;
            display: flex;
            align-items: center;
            gap: 0.75rem;
        }

        .meeting-section-header h2::before {
            content: '';
            display: inline-block;
            width: 4px;
            height: 1.25rem;
            background: var(--accent-blue);
            border-radius: 2px;
        }

        .meeting-section-date {
            font-size: 0.9rem;
            color: var(--text-secondary);
            font-family: 'Noto Sans Bengali', sans-serif;
        }

        @media (max-width: 1024px) {
            .dashboard-content-wrapper.list-mode {
                grid-template-columns: 1fr;
            }
            .toc-sidebar {
                position: static;
                max-height: 200px;
            }
        }

        /* Collapsible Agenda List Card */
        .agenda-list-card {
            background: var(--bg-card);
            backdrop-filter: blur(16px);
            border: 1px solid var(--border-glass);
            border-radius: 16px;
            margin-bottom: 1.25rem;
            box-shadow: var(--shadow-glass);
            transition: var(--transition);
            display: flex;
            flex-direction: column;
            animation: fadeIn 0.4s ease-out forwards;
            overflow: hidden;
        }

        .agenda-list-card:hover {
            border-color: rgba(59, 130, 246, 0.2);
            box-shadow: 0 8px 25px rgba(0, 0, 0, 0.25);
        }

        .agenda-list-card-header {
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
            padding: 1.25rem 1.5rem;
            cursor: pointer;
            user-select: none;
            transition: var(--transition);
            gap: 1.5rem;
        }

        .agenda-list-card.expanded .agenda-list-card-header {
            background: rgba(255, 255, 255, 0.02);
            border-bottom: 1px solid rgba(255, 255, 255, 0.08);
        }

        .header-main-info {
            display: flex;
            align-items: flex-start;
            gap: 0.75rem;
            flex: 1;
            min-width: 0;
        }

        .header-subject-preview {
            font-size: 0.95rem;
            color: var(--text-secondary);
            font-family: 'Noto Sans Bengali', sans-serif;
            flex: 1;
            line-height: 1.5;
            margin-left: 0.25rem;
            word-break: break-word;
        }

        .agenda-no-badge {
            background: rgba(255, 255, 255, 0.05);
            color: var(--text-secondary);
            border: 1px solid var(--border-glass);
            font-size: 0.75rem;
            font-weight: 600;
            padding: 0.25rem 0.6rem;
            border-radius: 6px;
            font-family: 'Noto Sans Bengali', sans-serif;
            white-space: nowrap;
            flex-shrink: 0;
        }

        .meta-badge {
            font-size: 0.75rem;
            font-weight: 600;
            padding: 0.25rem 0.75rem;
            border-radius: 12px;
            text-transform: uppercase;
            white-space: nowrap;
            flex-shrink: 0;
        }

        .header-right-info {
            display: flex;
            align-items: flex-start;
            gap: 1rem;
            flex-shrink: 0;
            margin-top: 0.15rem;
        }

        .header-fine-badge {
            color: var(--accent-red);
            font-weight: 700;
            font-size: 0.85rem;
            font-family: 'Noto Sans Bengali', sans-serif;
            background: rgba(244, 63, 94, 0.1);
            padding: 0.2rem 0.5rem;
            border-radius: 4px;
            border: 1px solid rgba(244, 63, 94, 0.2);
            white-space: nowrap;
        }

        .toggle-arrow {
            font-size: 0.8rem;
            color: var(--text-muted);
            transition: transform 0.3s cubic-bezier(0.4, 0, 0.2, 1);
            margin-top: 0.25rem;
        }

        .agenda-list-card.expanded .toggle-arrow {
            transform: rotate(180deg);
            color: var(--accent-blue);
        }

        .agenda-list-card-body {
            max-height: 0;
            opacity: 0;
            overflow: hidden;
            transition: max-height 0.4s cubic-bezier(0.4, 0, 0.2, 1), opacity 0.3s ease, padding 0.3s ease;
            padding: 0 1.5rem;
            display: flex;
            flex-direction: column;
            gap: 1.5rem;
        }

        .agenda-list-card.expanded .agenda-list-card-body {
            max-height: 4000px;
            opacity: 1;
            padding: 1.5rem;
        }

        /* Expanded parts */
        .agenda-expanded-part {
            display: flex;
            flex-direction: column;
            gap: 0.5rem;
        }

        .part-title {
            font-size: 0.85rem;
            font-weight: 600;
            color: var(--accent-blue);
            font-family: 'Noto Sans Bengali', sans-serif;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            border-left: 3px solid var(--accent-blue);
            padding-left: 0.5rem;
            margin-bottom: 0.25rem;
        }

        .part-content {
            font-family: 'Noto Sans Bengali', sans-serif;
            line-height: 1.7;
            font-size: 0.95rem;
            color: var(--text-secondary);
        }

        .part-content.subject-text {
            font-size: 1.2rem;
            color: #ffffff;
            font-weight: 600;
            line-height: 1.4;
        }

        .part-content.decision-box {
            background: linear-gradient(135deg, rgba(244, 63, 94, 0.07) 0%, rgba(20, 30, 50, 0.3) 100%);
            border: 1px solid rgba(244, 63, 94, 0.25);
            border-left: 6px solid var(--accent-red);
            border-radius: 12px;
            padding: 1.25rem 1.5rem;
            margin-top: 0.25rem;
        }

        /* Detail Block A (Subject) */
        .block-subject {
            background: rgba(255, 255, 255, 0.01);
            border: 1px solid var(--border-glass);
            border-radius: 16px;
            padding: 1.5rem;
        }

        .block-header-meta {
            display: flex;
            justify-content: space-between;
            flex-wrap: wrap;
            gap: 0.75rem;
            margin-bottom: 1rem;
        }

        .meta-badge {
            font-size: 0.75rem;
            font-weight: 600;
            padding: 0.25rem 0.75rem;
            border-radius: 12px;
            text-transform: uppercase;
        }

        .meta-badge.category {
            background: rgba(139, 92, 246, 0.15);
            color: #a78bfa;
            border: 1px solid rgba(139, 92, 246, 0.3);
        }

        .meta-badge.status {
            background: rgba(245, 158, 11, 0.15);
            color: #fbbf24;
            border: 1px solid rgba(245, 158, 11, 0.3);
        }

        .meta-badge.status.resolved {
            background: rgba(16, 185, 129, 0.15);
            color: #34d399;
            border: 1px solid rgba(16, 185, 129, 0.3);
        }

        .block-subject h2 {
            font-family: 'Noto Sans Bengali', sans-serif;
            font-size: 1.4rem;
            font-weight: 700;
            color: #ffffff;
            line-height: 1.4;
        }

        .inspector-row {
            margin-top: 1rem;
            font-size: 0.85rem;
            color: var(--text-secondary);
            display: flex;
            gap: 1.5rem;
            border-top: 1px solid rgba(255, 255, 255, 0.05);
            padding-top: 0.75rem;
        }

        /* Detail Block B (Decision - Crimson Callout Card) */
        .block-decision {
            background: linear-gradient(135deg, rgba(244, 63, 94, 0.07) 0%, rgba(20, 30, 50, 0.3) 100%);
            border: 1px solid rgba(244, 63, 94, 0.25);
            border-left: 6px solid var(--accent-red);
            border-radius: 16px;
            padding: 1.75rem;
            box-shadow: 0 10px 20px rgba(0, 0, 0, 0.2);
            position: relative;
        }

        .block-decision h3 {
            font-family: 'Noto Sans Bengali', sans-serif;
            font-size: 1.1rem;
            font-weight: 700;
            color: var(--accent-red);
            margin-bottom: 0.75rem;
            display: flex;
            align-items: center;
            gap: 0.5rem;
        }

        .decision-text {
            font-family: 'Noto Sans Bengali', sans-serif;
            font-size: 1.05rem;
            color: #ffe4e6;
            line-height: 1.7;
            white-space: pre-line;
        }

        /* Detail Block C (Accordion & Tables) */
        .block-details-accordion {
            display: flex;
            flex-direction: column;
            gap: 1rem;
        }

        details {
            background: rgba(255, 255, 255, 0.02);
            border: 1px solid var(--border-glass);
            border-radius: 12px;
            overflow: hidden;
            transition: var(--transition);
        }

        details[open] {
            border-color: rgba(255, 255, 255, 0.15);
            background: rgba(255, 255, 255, 0.03);
        }

        summary {
            padding: 1.25rem;
            font-weight: 600;
            font-size: 1rem;
            cursor: pointer;
            color: var(--text-primary);
            user-select: none;
            outline: none;
            display: flex;
            justify-content: space-between;
            align-items: center;
            font-family: 'Noto Sans Bengali', sans-serif;
        }

        summary::-webkit-details-marker {
            display: none;
        }

        summary::after {
            content: '▼';
            font-size: 0.8rem;
            color: var(--text-muted);
            transition: var(--transition);
        }

        details[open] summary::after {
            transform: rotate(180deg);
            color: var(--accent-blue);
        }

        .accordion-content {
            padding: 1.25rem;
            border-top: 1px solid var(--border-glass);
            font-size: 0.95rem;
            color: var(--text-secondary);
            font-family: 'Noto Sans Bengali', sans-serif;
            white-space: pre-line;
        }

        /* Violation and Stats Tables */
        .table-responsive {
            width: 100%;
            overflow-x: auto;
            margin-top: 1rem;
            border-radius: 10px;
            border: 1px solid var(--border-glass);
        }

        table {
            width: 100%;
            border-collapse: collapse;
            text-align: left;
            font-size: 0.85rem;
        }

        th {
            background: rgba(0, 0, 0, 0.4);
            color: #ffffff;
            font-weight: 600;
            padding: 0.75rem 1rem;
            border-bottom: 2px solid var(--border-glass);
        }

        td {
            padding: 0.75rem 1rem;
            border-bottom: 1px solid var(--border-glass);
            color: var(--text-secondary);
        }

        tr:nth-child(even) td {
            background: rgba(255, 255, 255, 0.01);
        }

        tr:hover td {
            background: rgba(255, 255, 255, 0.03);
            color: #ffffff;
        }

        /* Tab 2 Styles (Directives & coordination MoM) */
        .directive-card {
            background: var(--bg-card);
            border: 1px solid var(--border-glass);
            border-radius: 16px;
            padding: 1.75rem;
            margin-bottom: 1.5rem;
            box-shadow: var(--shadow-glass);
        }

        .directive-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 1px solid var(--border-glass);
            padding-bottom: 1rem;
            margin-bottom: 1.25rem;
        }

        .directive-header h2 {
            font-family: 'Noto Sans Bengali', sans-serif;
            font-size: 1.35rem;
            color: #ffffff;
        }

        .directive-header span {
            font-size: 0.85rem;
            color: var(--text-secondary);
            font-family: 'Noto Sans Bengali', sans-serif;
        }

        .directive-topic {
            margin-bottom: 1.5rem;
            background: rgba(255, 255, 255, 0.01);
            border: 1px solid rgba(255, 255, 255, 0.03);
            border-radius: 12px;
            padding: 1.25rem;
            transition: var(--transition);
        }

        .directive-topic:hover {
            background: rgba(255, 255, 255, 0.02);
            border-color: rgba(59, 130, 246, 0.15);
        }

        .directive-topic h3 {
            font-size: 1.05rem;
            font-family: 'Noto Sans Bengali', sans-serif;
            color: #60a5fa;
            margin-bottom: 0.5rem;
            display: flex;
            gap: 0.5rem;
            align-items: center;
        }

        .directive-topic h3 .badge {
            background: rgba(59, 130, 246, 0.15);
            padding: 0.15rem 0.5rem;
            border-radius: 6px;
            font-size: 0.8rem;
        }

        .directive-topic p {
            font-size: 0.95rem;
            color: var(--text-secondary);
            line-height: 1.6;
            margin-bottom: 0.75rem;
        }

        .directive-decision {
            background: rgba(16, 185, 129, 0.05);
            border: 1px solid rgba(16, 185, 129, 0.15);
            border-left: 4px solid var(--accent-green);
            padding: 0.75rem 1rem;
            border-radius: 6px;
            font-size: 0.9rem;
            color: #a7f3d0;
        }

        .directive-instruction-item {
            display: flex;
            gap: 1rem;
            padding: 1rem;
            border-bottom: 1px solid rgba(255, 255, 255, 0.05);
        }

        .directive-instruction-num {
            background: rgba(139, 92, 246, 0.15);
            color: #c084fc;
            width: 32px;
            height: 32px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-weight: 700;
            flex-shrink: 0;
            font-family: 'Noto Sans Bengali', sans-serif;
        }

        .directive-instruction-content h4 {
            font-family: 'Noto Sans Bengali', sans-serif;
            font-size: 1.05rem;
            margin-bottom: 0.25rem;
        }

        .directive-instruction-content p {
            font-size: 0.95rem;
            color: var(--text-secondary);
            line-height: 1.5;
        }

        /* Animations */
        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(10px); }
            to { opacity: 1; transform: translateY(0); }
        }

        .fade-in {
            animation: fadeIn 0.4s ease-out forwards;
        }

        /* Utilities */
        .hidden { display: none !important; }

        /* Scrollbar styles */
        ::-webkit-scrollbar {
            width: 8px;
            height: 8px;
        }
        ::-webkit-scrollbar-track {
            background: rgba(0, 0, 0, 0.1);
        }
        ::-webkit-scrollbar-thumb {
            background: rgba(255, 255, 255, 0.1);
            border-radius: 10px;
        }
        ::-webkit-scrollbar-thumb:hover {
            background: rgba(255, 255, 255, 0.2);
        }

        @media (max-width: 1200px) {
            .viewer-layout {
                grid-template-columns: 1fr;
            }
            .agenda-sidebar {
                border-right: none;
                border-bottom: 1px solid var(--border-glass);
                padding-right: 0;
                padding-bottom: 1.5rem;
                max-height: 300px;
            }
            .detail-viewer {
                grid-template-columns: 1fr;
                padding-left: 0;
            }
        }
    </style>
</head>
<body>

    <header>
        <div class="header-title-container">
            <div class="logo-placeholder">BTRC</div>
            <div class="header-title">
                <h1>বাংলাদেশ টেলিযোগাযোগ নিয়ন্ত্রণ কমিশন (বিটিআরসি)</h1>
                <p>কমিশন সভার কার্যবিবরণী ও সিদ্ধান্ত ট্র্যাকিং পোর্টাল</p>
            </div>
        </div>
        <div class="tab-switcher">
            <button class="tab-btn active" onclick="switchTab('meetings')">কমিশন সভার সিদ্ধান্ত</button>
            <button class="tab-btn" onclick="switchTab('directives')">সমন্বয় ও নির্দেশনাবলী</button>
        </div>
    </header>

    <div class="container">
        
        <!-- Tab 1: Commission Meetings -->
        <div id="tab-meetings-content" class="fade-in">
            
            <!-- Statistics Row -->
            <div class="stats-grid">
                <div class="stat-card">
                    <h3>মোট কমিশন সভা</h3>
                    <div id="stat-meetings-count" class="stat-value">০</div>
                    <div class="stat-subtitle">২৭৮তম থেকে ৩০৭তম সভার রেকর্ড</div>
                </div>
                <div class="stat-card green">
                    <h3>মোট এজেন্ডা সংখ্যা</h3>
                    <div id="stat-agendas-count" class="stat-value">০</div>
                    <div class="stat-subtitle">বিটিআরসি কমিশন সভার মোট এজেন্ডা</div>
                </div>
                <div class="stat-card red">
                    <h3>মোট প্রশাসনিক জরিমানা</h3>
                    <div id="stat-fines-count" class="stat-value">০</div>
                    <div class="stat-subtitle">টাকা জরিমানা আরোপিত (সকল বিভাগ)</div>
                </div>
                <div class="stat-card amber">
                    <h3>মোট রেভিনিউ শেয়ারিং দাবি</h3>
                    <div id="stat-claims-count" class="stat-value">০</div>
                    <div class="stat-subtitle">কোটি টাকা বকেয়া দাবি ট্র্যাকিং (সকল বিভাগ)</div>
                </div>
            </div>

            <!-- Dashboard Grid View (Layer 1) -->
            <div id="layer-grid-view">
                <!-- Dedicated Global Search & Analytics Panel -->
                <div class="search-panel-container">
                    <div class="search-panel-header">
                        <svg width="18" height="18" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="2" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"></circle><path d="m21 21-4.3-4.3"></path></svg>
                        <h2>কমিশন সভার এজেন্ডা ও সিদ্ধান্ত অনুসন্ধান প্যানেল</h2>
                    </div>
                    <div class="search-panel-body">
                        <div class="main-search-row">
                            <input type="text" id="search-bar" class="search-input-premium" placeholder="প্রতিষ্ঠান, সভার নম্বর, এজেন্ডার বিষয়বস্তু বা সিদ্ধান্ত দিয়ে অনুসন্ধান করুন...">
                            <button class="search-action-btn" onclick="filterMeetings()">অনুসন্ধান</button>
                        </div>
                        <div class="filters-row-premium">
                            <div class="filter-group">
                                <label for="filter-category">লাইসেন্স/সেবার ধরণ</label>
                                <select id="filter-category" class="filter-select-premium">
                                    <option value="">সকল ক্যাটাগরি</option>
                                    <option value="ISP">ISP (ইন্টারনেট)</option>
                                    <option value="MNO">MNO (মোবাইল অপারেটর)</option>
                                    <option value="IIG">IIG (গেটওয়ে)</option>
                                    <option value="ITC">ITC (ইন্টারন্যাশনাল ক্যারিয়ার)</option>
                                    <option value="TVAS">TVAS</option>
                                    <option value="IPTSP">IPTSP</option>
                                </select>
                            </div>
                            
                            <div class="filter-group">
                                <label for="filter-department">সংশ্লিষ্ট বিভাগ/দপ্তর</label>
                                <select id="filter-department" class="filter-select-premium">
                                    <option value="">সকল বিভাগ/দপ্তর</option>
                                    <option value="ei">এনফোর্সমেন্ট অ্যান্ড ইন্সপেকশন (E&I)</option>
                                    <option value="legal">আইন ও লাইসেন্সিং (Legal)</option>
                                    <option value="revenue">অর্থ, হিসাব ও রাজস্ব (Revenue)</option>
                                    <option value="spectrum">স্পেকট্রাম (Spectrum)</option>
                                    <option value="systems">সিস্টেমস অ্যান্ড সার্ভিসেস (S&S)</option>
                                </select>
                            </div>

                            <div class="filter-group">
                                <label for="filter-status">সিদ্ধান্ত/কেইসের অবস্থা</label>
                                <select id="filter-status" class="filter-select-premium">
                                    <option value="">সকল অবস্থা</option>
                                    <option value="জরিমানা">জরিমানা আরোপিত</option>
                                    <option value="ক্যাপিং">ব্যান্ডউইডথ ক্যাপিং</option>
                                    <option value="কারণ দর্শানো">কারণ দর্শানো (Show Cause)</option>
                                    <option value="নিষ্পন্ন">নিষ্পন্ন</option>
                                    <option value="চলমান">চলমান</option>
                                </select>
                            </div>

                            <div class="filter-group view-toggle-group">
                                <label>ভিউ মোড</label>
                                <div class="view-toggle-container-premium">
                                    <button id="btn-view-grid" class="toggle-view-btn-premium" onclick="setViewMode('grid')">
                                        <svg width="14" height="14" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="2" viewBox="0 0 24 24"><rect width="7" height="9" x="3" y="3" rx="1"></rect><rect width="7" height="5" x="14" y="3" rx="1"></rect><rect width="7" height="9" x="14" y="12" rx="1"></rect><rect width="7" height="5" x="3" y="16" rx="1"></rect></svg>
                                        গ্রিড
                                    </button>
                                    <button id="btn-view-list" class="toggle-view-btn-premium active" onclick="setViewMode('list')">
                                        <svg width="14" height="14" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="2" viewBox="0 0 24 24"><path d="M8 6h13M8 12h13M8 18h13M3 6h.01M3 12h.01M3 18h.01"></path></svg>
                                        লিস্ট
                                    </button>
                                </div>
                            </div>
                        </div>
                        <div class="search-stats-premium">
                            সর্বমোট এজেন্ডা পাওয়া গিয়েছে: <strong id="search-results-count">০</strong> টি
                        </div>
                    </div>
                </div>

                <div class="dashboard-content-wrapper list-mode">
                    <!-- TOC Sidebar for List View -->
                    <div id="list-toc-sidebar" class="toc-sidebar">
                        <h3>কমিশন সভাসমূহ</h3>
                        <div id="toc-links-container" class="toc-links">
                            <!-- TOC links will be dynamically inserted here -->
                        </div>
                    </div>

                    <div class="dashboard-main-area">
                        <div id="meetings-list-grid" class="meetings-grid hidden">
                            <!-- Cards will be dynamically inserted here -->
                        </div>

                        <div id="agendas-list-view">
                            <!-- Cards will be dynamically inserted here -->
                        </div>
                    </div>
                </div>
            </div>

            <!-- Details Split View (Layer 2 & 3) -->
            <div id="layer-details-view" class="viewer-layout hidden">
                
                <!-- Left Sidebar: Agenda list of Selected Meeting (Layer 2) -->
                <div class="agenda-sidebar">
                    <button class="back-btn" onclick="showDashboard()">
                        <svg width="16" height="16" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="2" viewBox="0 0 24 24"><path d="M19 12H5m7 7-7-7 7-7"></path></svg>
                        ড্যাশবোর্ডে ফিরে যান
                    </button>
                    <div class="sidebar-meeting-info">
                        <h2 id="sidebar-meeting-title">কমিশন সভা</h2>
                        <p id="sidebar-meeting-date">তারিখ: </p>
                    </div>
                    <ul id="agenda-sidebar-list" class="agenda-list">
                        <!-- Agenda items will be inserted here -->
                    </ul>
                </div>

                <!-- Right View: Selected Agenda Details (Layer 3) -->
                <div id="detail-viewer-content" class="detail-viewer">
                    <!-- Left Column: Core Info & Decision -->
                    <div class="detail-core-column">
                        <!-- Subject (Block A) -->
                        <div class="block-subject">
                            <div class="block-header-meta">
                                <span id="detail-meta-category" class="meta-badge category">ISP</span>
                                <span id="detail-meta-status" class="meta-badge status">চলমান</span>
                            </div>
                            <h2 id="detail-subject">এজেন্ডা বিষয়বস্তু</h2>
                            <div class="inspector-row">
                                <div id="detail-inspector"><strong>দায়িত্বপ্রাপ্ত কর্মকর্তা:</strong> -</div>
                                <div id="detail-due-status"><strong>অবস্থা:</strong> -</div>
                            </div>
                        </div>

                        <!-- Decision (Block B) -->
                        <div class="block-decision">
                            <h3>
                                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="2" viewBox="0 0 24 24"><path d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3Z"></path></svg>
                                কমিশন সভার সিদ্ধান্ত (E&I decision)
                            </h3>
                            <div id="detail-decision-text" class="decision-text">
                                কমিশনের সিদ্ধান্তের বিবরণ এখানে প্রদর্শিত হবে।
                            </div>
                        </div>
                    </div>

                    <!-- Right Column: Accordion Details (Block C) -->
                    <div class="block-details-accordion">
                        <details open>
                            <summary>প্রেক্ষাপট ও পরিদর্শনের সারসংক্ষেপ</summary>
                            <div id="detail-summary" class="accordion-content">
                                প্রেক্ষাপটের বিবরণ...
                            </div>
                        </details>
                        <details id="detail-tables-accordion">
                            <summary>ব্যত্যয়ের বিবরণ ও পরিমাপক টেবিল</summary>
                            <div id="detail-tables-container" class="accordion-content">
                                <!-- Tables dynamically generated -->
                            </div>
                        </details>
                        <details open>
                            <summary>বাস্তবায়নকারী বিভাগ/দপ্তর</summary>
                            <div id="detail-implementation" class="accordion-content">
                                বাস্তবায়ন সংক্রান্ত...
                            </div>
                        </details>
                    </div>

                </div>

            </div>

        </div>

        <!-- Tab 2: Coordination Meetings & Directives -->
        <div id="tab-directives-content" class="fade-in hidden">
            <!-- MoM Section -->
            <div id="directives-mom-container"></div>
            <!-- Commissioner Instructions Section -->
            <div id="directives-instructions-container"></div>
        </div>

    </div>

    <script>
        // Embed compiled master database
        const masterDb = __DB_JSON_PLACEHOLDER__;
        
        const meetings = masterDb.meetings;
        const coordinationMom = masterDb.coordination_mom;
        const commissionerInstructions = masterDb.commissioner_instructions;

        let activeMeetingIndex = null;
        let activeAgendaIndex = null;
        let viewMode = 'list';

        // Bengali and English number helpers
        const banglaDigits = {'০':'0','১':'1','২':'2','৩':'3','৪':'4','৫':'5','৬':'6','৭':'7','৮':'8','৯':'9'};
        const englishDigits = {'0':'০','1':'১','2':'২','3':'৩','4':'৪','5':'৫','6':'৬','7':'৭','8':'৮','9':'৯'};

        function bdToEn(str) {
            if (!str) return 0;
            // Clean non-digits except commas/periods and convert Bangla
            let cleaned = str.toString().replace(/[০-৯]/g, d => banglaDigits[d]);
            cleaned = cleaned.replace(/,/g, '');
            const val = parseFloat(cleaned);
            return isNaN(val) ? 0 : val;
        }

        function enToBd(num) {
            if (num === undefined || num === null) return '';
            let str = num.toLocaleString('en-US');
            return str.replace(/[0-9]/g, d => englishDigits[d]);
        }

        // Initialize and compute stats
        function init() {
            // Compute total stats
            document.getElementById('stat-meetings-count').innerText = enToBd(meetings.length) + 'টি';
            
            let totalAgendas = 0;
            let totalFines = 0;
            let totalClaims = 0;

            meetings.forEach(m => {
                totalAgendas += m.agendas.length;
                m.agendas.forEach(a => {
                    if (a.fine_amount) {
                        totalFines += bdToEn(a.fine_amount);
                    }
                    if (a.revenue_sharing_claim) {
                        totalClaims += bdToEn(a.revenue_sharing_claim);
                    }
                });
            });

            document.getElementById('stat-agendas-count').innerText = enToBd(totalAgendas) + 'টি';
            
            // Format fines in Crore/Lakh
            if (totalFines >= 10000000) {
                document.getElementById('stat-fines-count').innerText = enToBd(parseFloat((totalFines / 10000000).toFixed(2))) + ' কোটি';
            } else {
                document.getElementById('stat-fines-count').innerText = enToBd(parseFloat((totalFines / 100000).toFixed(2))) + ' লাখ';
            }

            // Format claims in Crore
            document.getElementById('stat-claims-count').innerText = enToBd(parseFloat((totalClaims / 10000000).toFixed(2))) + ' কোটি';

            // Render Layer 1 list view by default
            setViewMode('list');
            
            // Render Tab 2 directives and coordination minutes
            renderDirectivesTab();

            // Set up search and filters
            document.getElementById('search-bar').addEventListener('input', filterMeetings);
            document.getElementById('filter-category').addEventListener('change', filterMeetings);
            document.getElementById('filter-department').addEventListener('change', filterMeetings);
            document.getElementById('filter-status').addEventListener('change', filterMeetings);
        }

        // Switch Main Tabs
        function switchTab(tab) {
            const tabBtnMeetings = document.querySelectorAll('.tab-btn')[0];
            const tabBtnDirectives = document.querySelectorAll('.tab-btn')[1];
            const contentMeetings = document.getElementById('tab-meetings-content');
            const contentDirectives = document.getElementById('tab-directives-content');

            if (tab === 'meetings') {
                tabBtnMeetings.classList.add('active');
                tabBtnDirectives.classList.remove('active');
                contentMeetings.classList.remove('hidden');
                contentDirectives.classList.add('hidden');
            } else {
                tabBtnMeetings.classList.remove('active');
                tabBtnDirectives.classList.add('active');
                contentMeetings.classList.add('hidden');
                contentDirectives.classList.remove('hidden');
            }
        }

        // Render Layer 1 Dashboard grid
        function renderMeetingsGrid(filteredMeetings = meetings) {
            const grid = document.getElementById('meetings-list-grid');
            grid.innerHTML = '';

            if (filteredMeetings.length === 0) {
                grid.innerHTML = `<div style="grid-column: 1/-1; text-align: center; color: var(--text-secondary); padding: 3rem; font-family: 'Noto Sans Bengali', sans-serif;">কোনো সভা বা সিদ্ধান্ত পাওয়া যায়নি।</div>`;
                return;
            }

            filteredMeetings.forEach((m, index) => {
                // Find actual index in primary meetings array
                const origIndex = meetings.findIndex(origM => origM.meeting_number === m.meeting_number);
                
                // Calculate meeting specific total fines
                let meetingFines = 0;
                m.agendas.forEach(a => {
                    if (a.fine_amount) meetingFines += bdToEn(a.fine_amount);
                });

                let finesStr = '০ টাকা';
                if (meetingFines > 0) {
                    if (meetingFines >= 10000000) {
                        finesStr = enToBd(parseFloat((meetingFines / 10000000).toFixed(2))) + ' কোটি টাকা';
                    } else if (meetingFines >= 100000) {
                        finesStr = enToBd(parseFloat((meetingFines / 100000).toFixed(2))) + ' লাখ টাকা';
                    } else {
                        finesStr = enToBd(meetingFines) + ' টাকা';
                    }
                }

                const card = document.createElement('div');
                card.className = 'meeting-card fade-in';
                card.onclick = () => selectMeeting(origIndex);

                card.innerHTML = `
                    <div>
                        <span class="meeting-num-badge">${m.meeting_number}</span>
                        <h2>${m.meeting_number} কমিশন সভা</h2>
                        <div class="date">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="2" viewBox="0 0 24 24"><rect width="18" height="18" x="3" y="4" rx="2" ry="2"></rect><path d="M16 2v4M8 2v4m-5 4h18"></path></svg>
                            ${m.meeting_date}
                        </div>
                    </div>
                    <div class="meeting-info-row">
                        <div class="meeting-info-item">এজেন্ডা: <strong>${enToBd(m.agendas.length)}টি</strong></div>
                        <div class="meeting-info-item" style="color: ${meetingFines > 0 ? 'var(--accent-red)' : 'var(--text-muted)'}">জরিমানা: <strong>${finesStr}</strong></div>
                    </div>
                `;
                grid.appendChild(card);
            });
        }

        // View Mode Toggle Handler
        function setViewMode(mode) {
            viewMode = mode;
            const btnGrid = document.getElementById('btn-view-grid');
            const btnList = document.getElementById('btn-view-list');
            const gridContainer = document.getElementById('meetings-list-grid');
            const listContainer = document.getElementById('agendas-list-view');
            const sidebar = document.getElementById('list-toc-sidebar');
            const wrapper = document.querySelector('.dashboard-content-wrapper');
            
            if (mode === 'grid') {
                btnGrid.classList.add('active');
                btnList.classList.remove('active');
                gridContainer.classList.remove('hidden');
                listContainer.classList.add('hidden');
                sidebar.classList.add('hidden');
                wrapper.classList.remove('list-mode');
            } else {
                btnGrid.classList.remove('active');
                btnList.classList.add('active');
                gridContainer.classList.add('hidden');
                listContainer.classList.remove('hidden');
                sidebar.classList.remove('hidden');
                wrapper.classList.add('list-mode');
            }
            filterMeetings();
        }

        // Search and Filters Logic
        function filterMeetings() {
            const query = document.getElementById('search-bar').value.toLowerCase().trim();
            const category = document.getElementById('filter-category').value;
            const status = document.getElementById('filter-status').value;
            const dept = document.getElementById('filter-department').value;

            const matchedAgendas = [];
            const matchedMeetings = [];

            meetings.forEach(m => {
                const matchingAgendasInMeeting = m.agendas.filter(a => {
                    const matchesCategory = !category || a.subject.toUpperCase().includes(category) || a.details.presentation_summary.toUpperCase().includes(category);
                    
                    let matchesStatus = true;
                    if (status) {
                        if (status === 'জরিমানা') {
                            matchesStatus = parseFloat(bdToEn(a.fine_amount)) > 0 || a.decision.includes('জরিমানা') || (a.fine_amount && a.fine_amount !== 'nan');
                        } else if (status === 'ক্যাপিং') {
                            matchesStatus = a.decision.includes('ক্যাপিং') || a.decision.includes('ক্যাপ');
                        } else if (status === 'কারণ দর্শানো') {
                            matchesStatus = a.decision.includes('কারণ দর্শানো') || a.decision.includes('Show Cause') || a.details.case_status.includes('কারণ দর্শানো');
                        } else if (status === 'নিষ্পন্ন') {
                            matchesStatus = a.details.case_status.includes('নিষ্পন্ন') || a.details.case_status.includes('আদায়');
                        } else if (status === 'চলমান') {
                            matchesStatus = a.details.case_status.includes('চলমান') || a.details.case_status.includes('শুনানি');
                        }
                    }

                    let matchesDept = true;
                    if (dept) {
                        const implText = (a.details.implementation || '').toLowerCase();
                        const inspectorText = (a.details.assigned_inspector || '').toLowerCase();
                        const subjectText = a.subject.toLowerCase();
                        
                        if (dept === 'ei') {
                            matchesDept = implText.includes('ইএন্ডআই') || implText.includes('enforcement') || implText.includes('e&i') || inspectorText.includes('e&i') || implText.includes('এনফোর্সমেন্ট');
                        } else if (dept === 'legal') {
                            matchesDept = implText.includes('আইন') || implText.includes('লাইসেন্সিং') || implText.includes('legal') || implText.includes('licensing') || implText.includes('এলএল') || subjectText.includes('আইনি') || subjectText.includes('মামলা');
                        } else if (dept === 'revenue') {
                            matchesDept = implText.includes('অর্থ') || implText.includes('হিসাব') || implText.includes('রাজস্ব') || implText.includes('revenue') || implText.includes('finance') || subjectText.includes('বকেয়া') || subjectText.includes('দাবি');
                        } else if (dept === 'spectrum') {
                            matchesDept = implText.includes('স্পেকট্রাম') || implText.includes('spectrum');
                        } else if (dept === 'systems') {
                            matchesDept = implText.includes('সিস্টেম') || implText.includes('systems') || implText.includes('সার্ভিস');
                        }
                    }

                    const matchesSearch = !query || 
                        m.meeting_number.toLowerCase().includes(query) || 
                        a.subject.toLowerCase().includes(query) || 
                        a.decision.toLowerCase().includes(query) || 
                        a.details.presentation_summary.toLowerCase().includes(query);

                    return matchesCategory && matchesStatus && matchesDept && matchesSearch;
                });

                if (matchingAgendasInMeeting.length > 0) {
                    matchedMeetings.push({
                        ...m,
                        agendas: matchingAgendasInMeeting
                    });
                    
                    matchingAgendasInMeeting.forEach(a => {
                        matchedAgendas.push({
                            meeting_number: m.meeting_number,
                            meeting_date: m.meeting_date,
                            agenda: a
                        });
                    });
                }
            });

            // Update search stats count
            document.getElementById('search-results-count').innerText = enToBd(matchedAgendas.length);

            if (viewMode === 'grid') {
                renderMeetingsGrid(matchedMeetings);
            } else {
                renderAgendasList(matchedAgendas);
            }
        }

        // Render List View of all Agendas (Continuous scrolling grouped by meeting with TOC)
        function renderAgendasList(agendasList) {
            const container = document.getElementById('agendas-list-view');
            const tocContainer = document.getElementById('toc-links-container');
            container.innerHTML = '';
            tocContainer.innerHTML = '';

            if (agendasList.length === 0) {
                container.innerHTML = `<div style="text-align: center; color: var(--text-secondary); padding: 3rem; font-family: 'Noto Sans Bengali', sans-serif;">কোনো এজেন্ডা বা সিদ্ধান্ত পাওয়া যায়নি।</div>`;
                tocContainer.innerHTML = `<div style="color: var(--text-muted); font-size: 0.8rem; text-align: center; font-family: 'Noto Sans Bengali', sans-serif;">খালি</div>`;
                return;
            }

            // Group by meeting
            const meetingsGrouped = {};
            agendasList.forEach(item => {
                if (!meetingsGrouped[item.meeting_number]) {
                    meetingsGrouped[item.meeting_number] = {
                        meeting_number: item.meeting_number,
                        meeting_date: item.meeting_date,
                        agendas: []
                    };
                }
                meetingsGrouped[item.meeting_number].agendas.push(item.agenda);
            });

            // Render the grouped meetings
            Object.values(meetingsGrouped).forEach(mGroup => {
                const mNo = mGroup.meeting_number;
                const mDate = mGroup.meeting_date;
                
                // Add TOC Link
                const tocLink = document.createElement('a');
                tocLink.href = `#meeting-sec-${mNo}`;
                tocLink.className = 'toc-link';
                tocLink.onclick = (e) => {
                    e.preventDefault();
                    const targetEl = document.getElementById(`meeting-sec-${mNo}`);
                    if (targetEl) {
                        targetEl.scrollIntoView({ behavior: 'smooth' });
                    }
                    // Set active class
                    document.querySelectorAll('.toc-link').forEach(link => link.classList.remove('active'));
                    tocLink.classList.add('active');
                };
                tocLink.innerHTML = `
                    <span>${mNo} কমিশন সভা</span>
                    <span class="toc-link-badge">${enToBd(mGroup.agendas.length)}</span>
                `;
                tocContainer.appendChild(tocLink);

                // Add Section Header in list view
                const secHeader = document.createElement('div');
                secHeader.id = `meeting-sec-${mNo}`;
                secHeader.className = 'meeting-section-header';
                secHeader.innerHTML = `
                    <h2>${mNo} কমিশন সভা</h2>
                    <span class="meeting-section-date">তারিখ: ${mDate}</span>
                `;
                container.appendChild(secHeader);

                // Add agendas
                mGroup.agendas.forEach(a => {
                    let cat = 'E&I';
                    if (a.subject.includes('ISP') || a.details.presentation_summary.includes('ISP')) cat = 'ISP';
                    else if (a.subject.includes('MNO') || a.details.presentation_summary.includes('MNO') || a.subject.includes('গ্রামীণ') || a.subject.includes('রবি')) cat = 'MNO';
                    else if (a.subject.includes('IIG')) cat = 'IIG';
                    else if (a.subject.includes('ITC')) cat = 'ITC';
                    else if (a.subject.includes('TVAS')) cat = 'TVAS';
                    else if (a.subject.includes('IPTSP')) cat = 'IPTSP';

                    const status = a.details.case_status || 'চলমান';
                    const isResolved = status.includes('নিষ্পন্ন') || status.includes('অনুমোদিত');

                    // Render tables if they exist
                    let tablesHtml = '';
                    if (a.details.tables && a.details.tables.length > 0) {
                        a.details.tables.forEach(tableData => {
                            tablesHtml += `<div class="table-responsive"><table>`;
                            let thead = '<tr>';
                            tableData.headers.forEach(h => { thead += `<th>${h}</th>`; });
                            thead += '</tr>';
                            let tbody = '';
                            tableData.rows.forEach(r => {
                                tbody += '<tr>';
                                r.forEach(cell => { tbody += `<td>${cell}</td>`; });
                                tbody += '</tr>';
                            });
                            tablesHtml += `<thead>${thead}</thead><tbody>${tbody}</tbody></table></div>`;
                        });
                    }

                    const card = document.createElement('div');
                    card.className = 'agenda-list-card';
                    card.onclick = (e) => toggleAgendaCard(card, e);
                    
                    let formattedDecision = a.decision;
                    if (a.fine_amount) {
                        formattedDecision = formattedDecision.replace(a.fine_amount, `<strong>${a.fine_amount}</strong>`);
                    }

                    // Display the full subject without truncation
                    let subjectPreview = a.subject;

                    card.innerHTML = `
                        <div class="agenda-list-card-header">
                            <div class="header-main-info">
                                <span class="agenda-no-badge">এজেন্ডা ${enToBd(a.agenda_no)}</span>
                                <span class="meta-badge category">${cat}</span>
                                <span class="meta-badge status ${isResolved ? 'resolved' : ''}">${status}</span>
                                <span class="header-subject-preview">${subjectPreview}</span>
                            </div>
                            <div class="header-right-info">
                                ${a.fine_amount ? `<span class="header-fine-badge">জরিমানা: ${enToBd(bdToEn(a.fine_amount))} টাকা</span>` : ''}
                                <span class="toggle-arrow">▼</span>
                            </div>
                        </div>
                        <div class="agenda-list-card-body">
                            <!-- Part 1: Subject -->
                            <div class="agenda-expanded-part">
                                <div class="part-title">১. এজেন্ডার বিষয় (Subject)</div>
                                <div class="part-content subject-text">${a.subject}</div>
                            </div>
                            
                            <!-- Part 2: Body -->
                            <div class="agenda-expanded-part">
                                <div class="part-title">২. প্রেক্ষাপট ও পরিদর্শনের সারসংক্ষেপ (Body)</div>
                                <div class="part-content">
                                    <div style="margin-bottom: 0.75rem;">${a.details.presentation_summary}</div>
                                    ${tablesHtml}
                                    <div style="margin-top: 1rem; font-size: 0.85rem; color: var(--text-muted);">
                                        <strong>দায়িত্বপ্রাপ্ত কর্মকর্তা:</strong> ${a.details.assigned_inspector || 'এনফোর্সমেন্ট টিম'} | 
                                        <strong>বাস্তবায়নকারী বিভাগ/দপ্তর:</strong> ${a.details.implementation || 'এনফোর্সমেন্ট এন্ড ইন্সপেকশন ডিরেক্টরেট'}
                                    </div>
                                </div>
                            </div>
                            
                            <!-- Part 3: Decision -->
                            <div class="agenda-expanded-part">
                                <div class="part-title">৩. কমিশন সভার সিদ্ধান্ত (Decision)</div>
                                <div class="part-content decision-box">
                                    <div class="decision-text">${formattedDecision}</div>
                                </div>
                            </div>
                        </div>
                    `;
                    container.appendChild(card);
                });
            });
            
            // Set first link active by default
            const firstLink = tocContainer.querySelector('.toc-link');
            if (firstLink) firstLink.classList.add('active');
        }

        // Toggle Agenda Card Expanded State
        function toggleAgendaCard(cardElement, event) {
            // Do not toggle if click was inside link, details summary or table
            if (event.target.closest('a') || event.target.closest('button') || event.target.closest('details') || event.target.closest('table')) {
                return;
            }
            cardElement.classList.toggle('expanded');
        }

        // Select and Open Meeting details (Layer 2)
        function selectMeeting(index) {
            activeMeetingIndex = index;
            const meeting = meetings[index];

            document.getElementById('sidebar-meeting-title').innerText = meeting.meeting_number + ' কমিশন সভা';
            document.getElementById('sidebar-meeting-date').innerText = 'তারিখ: ' + meeting.meeting_date;

            const list = document.getElementById('agenda-sidebar-list');
            list.innerHTML = '';

            meeting.agendas.forEach((a, aIdx) => {
                const item = document.createElement('li');
                item.className = `agenda-item ${aIdx === 0 ? 'active' : ''}`;
                item.onclick = () => selectAgenda(aIdx);
                
                item.innerHTML = `
                    <div class="agenda-item-header">
                        <span class="agenda-badge">এজেন্ডা ${enToBd(a.agenda_no)}</span>
                        ${a.fine_amount ? '<span style="color: var(--accent-red); font-size: 0.7rem; font-weight: 700;">জরিমানা</span>' : ''}
                    </div>
                    <div class="agenda-item-title">${a.subject}</div>
                `;
                list.appendChild(item);
            });

            // Toggle screens
            document.getElementById('layer-grid-view').classList.add('hidden');
            document.getElementById('layer-details-view').classList.remove('hidden');

            // Select first agenda by default
            selectAgenda(0);
        }

        // Show Dashboard grid (Back from details)
        function showDashboard() {
            document.getElementById('layer-grid-view').classList.remove('hidden');
            document.getElementById('layer-details-view').classList.add('hidden');
            activeMeetingIndex = null;
            activeAgendaIndex = null;
        }

        // Select Agenda item (Layer 3)
        function selectAgenda(agendaIndex) {
            activeAgendaIndex = agendaIndex;
            const meeting = meetings[activeMeetingIndex];
            const agenda = meeting.agendas[agendaIndex];

            // Update active state in sidebar
            const items = document.querySelectorAll('.agenda-item');
            items.forEach((item, idx) => {
                if (idx === agendaIndex) {
                    item.classList.add('active');
                } else {
                    item.classList.remove('active');
                }
            });

            // Subject (Block A)
            document.getElementById('detail-subject').innerText = agenda.subject;
            
            // Detect and display category
            let cat = 'E&I';
            if (agenda.subject.includes('ISP') || agenda.details.presentation_summary.includes('ISP')) cat = 'ISP';
            else if (agenda.subject.includes('MNO') || agenda.details.presentation_summary.includes('MNO') || agenda.subject.includes('গ্রামীণ') || agenda.subject.includes('রবি')) cat = 'MNO';
            else if (agenda.subject.includes('IIG')) cat = 'IIG';
            else if (agenda.subject.includes('ITC')) cat = 'ITC';
            else if (agenda.subject.includes('TVAS')) cat = 'TVAS';
            else if (agenda.subject.includes('IPTSP')) cat = 'IPTSP';
            
            document.getElementById('detail-meta-category').innerText = cat;

            // Status details
            const status = agenda.details.case_status || 'চলমান';
            const statusBadge = document.getElementById('detail-meta-status');
            statusBadge.innerText = status;
            if (status.includes('নিষ্পন্ন') || status.includes('অনুমোদিত')) {
                statusBadge.className = 'meta-badge status resolved';
            } else {
                statusBadge.className = 'meta-badge status';
            }

            document.getElementById('detail-inspector').innerHTML = `<strong>দায়িত্বপ্রাপ্ত কর্মকর্তা:</strong> ${agenda.details.assigned_inspector || 'এনফোর্সমেন্ট টিম'}`;
            document.getElementById('detail-due-status').innerHTML = `<strong>কেইস অবস্থা:</strong> ${status}`;

            // Decision (Block B - Crimson Callout Card)
            let formattedDecision = agenda.decision;
            if (agenda.fine_amount) {
                formattedDecision = formattedDecision.replace(agenda.fine_amount, `<strong>${agenda.fine_amount}</strong>`);
            }
            document.getElementById('detail-decision-text').innerHTML = formattedDecision;

            // Accordion Details (Block C)
            document.getElementById('detail-summary').innerText = agenda.details.presentation_summary;
            document.getElementById('detail-implementation').innerText = agenda.details.implementation || 'এনফোর্সমেন্ট এন্ড ইন্সপেকশন ডিরেক্টরেট';

            // Violation tables rendering
            const tablesContainer = document.getElementById('detail-tables-container');
            const tablesAccordion = document.getElementById('detail-tables-accordion');
            
            if (agenda.details.tables && agenda.details.tables.length > 0) {
                tablesAccordion.style.display = 'block';
                tablesAccordion.setAttribute('open', 'true');
                tablesContainer.innerHTML = '';

                agenda.details.tables.forEach(tableData => {
                    const tableResponsive = document.createElement('div');
                    tableResponsive.className = 'table-responsive';
                    
                    const table = document.createElement('table');
                    
                    // Headers
                    let theadHtml = '<tr>';
                    tableData.headers.forEach(h => {
                        theadHtml += `<th>${h}</th>`;
                    });
                    theadHtml += '</tr>';
                    
                    // Rows
                    let tbodyHtml = '';
                    tableData.rows.forEach(r => {
                        tbodyHtml += '<tr>';
                        r.forEach(cell => {
                            tbodyHtml += `<td>${cell}</td>`;
                        });
                        tbodyHtml += '</tr>';
                    });
                    
                    table.innerHTML = `<thead>${theadHtml}</thead><tbody>${tbodyHtml}</tbody>`;
                    tableResponsive.appendChild(table);
                    tablesContainer.appendChild(tableResponsive);
                });
            } else {
                tablesAccordion.style.display = 'none';
                tablesAccordion.removeAttribute('open');
                tablesContainer.innerHTML = '';
            }
        }

        // Render Tab 2 directives and coordination minutes
        function renderDirectivesTab() {
            const momContainer = document.getElementById('directives-mom-container');
            momContainer.innerHTML = '';

            coordinationMom.forEach(mom => {
                const momCard = document.createElement('div');
                momCard.className = 'directive-card fade-in';
                
                let topicsHtml = '';
                mom.topics.forEach((t, idx) => {
                    topicsHtml += `
                        <div class="directive-topic">
                            <h3><span class="badge">${enToBd(idx + 1)}</span> ${t.section}</h3>
                            <p>${t.summary}</p>
                            <div class="directive-decision">
                                <strong>মাননীয় কমিশনারের সিদ্ধান্ত:</strong> ${t.decision}
                            </div>
                        </div>
                    `;
                });

                momCard.innerHTML = `
                    <div class="directive-header">
                        <h2>${mom.title}</h2>
                        <span>তারিখ: ${mom.date} | সভাপতিত্বে: ${mom.chairperson}</span>
                    </div>
                    <div style="font-size: 0.9rem; color: var(--text-secondary); margin-bottom: 1.5rem; background: rgba(255,255,255,0.02); padding: 1rem; border-radius: 8px; border: 1px solid var(--border-glass)">
                        <strong>অংশগ্রহণকারী:</strong> ${mom.attendees}
                    </div>
                    <div>
                        ${topicsHtml}
                    </div>
                `;
                momContainer.appendChild(momCard);
            });

            const instructionsContainer = document.getElementById('directives-instructions-container');
            instructionsContainer.innerHTML = '';

            commissionerInstructions.forEach(ins => {
                const insCard = document.createElement('div');
                insCard.className = 'directive-card fade-in';
                
                let instListHtml = '';
                ins.instructions.forEach(item => {
                    instListHtml += `
                        <div class="directive-instruction-item">
                            <div class="directive-instruction-num">${item.id}</div>
                            <div class="directive-instruction-content">
                                <h4>${item.title}</h4>
                                <p>${item.text}</p>
                            </div>
                        </div>
                    `;
                });

                insCard.innerHTML = `
                    <div class="directive-header">
                        <h2>${ins.title}</h2>
                        <span>উৎস: ${ins.source} | তারিখ: ${ins.date}</span>
                    </div>
                    <div style="background: rgba(255,255,255,0.01); border: 1px solid var(--border-glass); border-radius: 12px; overflow: hidden;">
                        ${instListHtml}
                    </div>
                `;
                instructionsContainer.appendChild(insCard);
            });
        }

        // Start initialization on window load
        window.onload = init;
    </script>
</body>
</html>
"""

# Serialize the database content and embed it
db_json_string = json.dumps(master_db, ensure_ascii=False)
html_content = html_template.replace("__DB_JSON_PLACEHOLDER__", db_json_string)

# Overwrite / Write the file to path
with open(output_html_path, "w", encoding="utf-8") as f:
    f.write(html_content)

print(f"Interactive Reader HTML successfully compiled and saved to: {output_html_path}")
print(f"HTML size: {os.path.getsize(output_html_path)} bytes")
