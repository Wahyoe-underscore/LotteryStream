import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import secrets
import hashlib
from io import BytesIO, StringIO
import time
import re
import requests
import json
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

PRIZE_CONFIG_FILE = "prize_config.json"

PRIZE_TIERS = [
    {"name": "Tokopedia Rp.100.000,-", "icon": "🛒", "count": 175, "start": 1, "end": 175},
    {"name": "Indomaret Rp.100.000,-", "icon": "🏪", "count": 175, "start": 176, "end": 350},
    {"name": "Bensin Rp.100.000,-", "icon": "⛽", "count": 175, "start": 351, "end": 525},
    {"name": "SNL Rp.100.000,-", "icon": "🎵", "count": 175, "start": 526, "end": 700},
]

SHUFFLE_CONFIG = [
    {"name": "Sesi 1", "count": 30, "prize": ""},
    {"name": "Sesi 2", "count": 30, "prize": ""},
    {"name": "Sesi 3", "count": 30, "prize": ""},
]

WHEEL_CONFIG = {"count": 10}

def load_prize_config():
    if os.path.exists(PRIZE_CONFIG_FILE):
        try:
            with open(PRIZE_CONFIG_FILE, 'r') as f:
                return json.load(f)
        except:
            return None
    return None

def save_prize_config(config):
    with open(PRIZE_CONFIG_FILE, 'w') as f:
        json.dump(config, f, indent=2)

def calculate_total_winners(prize_tiers):
    return sum(tier["count"] for tier in prize_tiers)

def secure_shuffle(items):
    items = list(items)
    for i in range(len(items) - 1, 0, -1):
        j = secrets.randbelow(i + 1)
        items[i], items[j] = items[j], items[i]
    return items

def is_eligible_for_prize(name, phone):
    name_str = str(name).upper() if pd.notna(name) else ""
    phone_str = str(phone) if pd.notna(phone) else ""
    if "VIP" in name_str or name_str.startswith("F ") or name_str.endswith(" F") or " F " in name_str:
        return False
    if phone_str.startswith("F") or phone_str.upper() == "F":
        return False
    return True

def get_prize_dynamic(rank, prize_tiers):
    for tier in prize_tiers:
        if tier["start"] <= rank <= tier["end"]:
            return tier["name"]
    return "Hadiah"

def mask_phone(phone):
    phone_str = str(phone) if pd.notna(phone) else ""
    if len(phone_str) >= 4:
        return f"****{phone_str[-4:]}"
    return phone_str

def create_spinning_wheel_html(participants, winner, wheel_size=400):
    num_segments = min(len(participants), 20)
    display_participants = participants[:num_segments]
    
    if winner not in display_participants:
        display_participants[-1] = winner
    
    winner_index = display_participants.index(winner)
    
    colors = [
        '#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7',
        '#DDA0DD', '#98D8C8', '#F7DC6F', '#BB8FCE', '#85C1E9',
        '#F8B500', '#FF6F61', '#6B5B95', '#88B04B', '#F7CAC9',
        '#92A8D1', '#955251', '#B565A7', '#009B77', '#DD4124'
    ]
    
    segments_js = []
    for i, p in enumerate(display_participants):
        segments_js.append(f'{{"label": "{p}", "color": "{colors[i % len(colors)]}"}}')
    
    rotation_per_segment = 360 / num_segments
    target_rotation = 360 * 8 + (360 - (winner_index * rotation_per_segment + rotation_per_segment / 2))
    
    html = f'''
    <!DOCTYPE html>
    <html>
    <head>
        <style>
            * {{ margin: 0; padding: 0; box-sizing: border-box; }}
            body {{
                display: flex;
                flex-direction: column;
                align-items: center;
                justify-content: center;
                min-height: 100vh;
                background: transparent;
                font-family: 'Poppins', sans-serif;
            }}
            .wheel-container {{
                position: relative;
                width: {wheel_size}px;
                height: {wheel_size}px;
            }}
            .pointer {{
                position: absolute;
                top: -20px;
                left: 50%;
                transform: translateX(-50%);
                width: 0;
                height: 0;
                border-left: 20px solid transparent;
                border-right: 20px solid transparent;
                border-top: 40px solid #f5576c;
                z-index: 100;
                filter: drop-shadow(0 3px 6px rgba(0,0,0,0.3));
            }}
            .center-circle {{
                position: absolute;
                top: 50%;
                left: 50%;
                transform: translate(-50%, -50%);
                width: 80px;
                height: 80px;
                background: linear-gradient(145deg, #fff, #f0f0f0);
                border-radius: 50%;
                display: flex;
                align-items: center;
                justify-content: center;
                font-size: 2rem;
                box-shadow: 0 4px 15px rgba(0,0,0,0.2);
                z-index: 50;
            }}
            .winner-display {{
                margin-top: 20px;
                padding: 20px 40px;
                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                border-radius: 15px;
                color: white;
                font-size: 1.5rem;
                font-weight: 700;
                display: none;
                animation: popIn 0.5s ease;
            }}
            @keyframes popIn {{
                0% {{ transform: scale(0); opacity: 0; }}
                100% {{ transform: scale(1); opacity: 1; }}
            }}
        </style>
    </head>
    <body>
        <div class="wheel-container">
            <div class="pointer"></div>
            <canvas id="wheel" width="{wheel_size}" height="{wheel_size}"></canvas>
            <div class="center-circle">🎯</div>
        </div>
        <div class="winner-display" id="winnerDisplay">🎉 PEMENANG: <span id="winnerText"></span></div>
        <script>
            const segments = [{','.join(segments_js)}];
            const canvas = document.getElementById('wheel');
            const ctx = canvas.getContext('2d');
            const centerX = canvas.width / 2;
            const centerY = canvas.height / 2;
            const radius = canvas.width / 2 - 10;
            let currentRotation = 0;
            
            function drawWheel(rotation) {{
                ctx.clearRect(0, 0, canvas.width, canvas.height);
                ctx.save();
                ctx.translate(centerX, centerY);
                ctx.rotate(rotation * Math.PI / 180);
                
                const segmentAngle = (2 * Math.PI) / segments.length;
                
                segments.forEach((segment, index) => {{
                    ctx.beginPath();
                    ctx.moveTo(0, 0);
                    ctx.arc(0, 0, radius, index * segmentAngle, (index + 1) * segmentAngle);
                    ctx.closePath();
                    ctx.fillStyle = segment.color;
                    ctx.fill();
                    ctx.strokeStyle = '#fff';
                    ctx.lineWidth = 2;
                    ctx.stroke();
                    
                    ctx.save();
                    ctx.rotate(index * segmentAngle + segmentAngle / 2);
                    ctx.textAlign = 'right';
                    ctx.fillStyle = '#fff';
                    ctx.font = 'bold 12px Arial';
                    ctx.fillText(segment.label, radius - 15, 5);
                    ctx.restore();
                }});
                
                ctx.restore();
            }}
            
            drawWheel(0);
            
            const targetRotation = {target_rotation};
            const duration = 5000;
            const startTime = Date.now();
            
            function animate() {{
                const elapsed = Date.now() - startTime;
                const progress = Math.min(elapsed / duration, 1);
                const easeOut = 1 - Math.pow(1 - progress, 4);
                currentRotation = easeOut * targetRotation;
                drawWheel(currentRotation);
                
                if (progress < 1) {{
                    requestAnimationFrame(animate);
                }} else {{
                    document.getElementById('winnerDisplay').style.display = 'block';
                    document.getElementById('winnerText').textContent = '{winner}';
                }}
            }}
            
            setTimeout(animate, 500);
        </script>
    </body>
    </html>
    '''
    return html

def generate_pptx(results_df, prize_tiers):
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    
    for tier in prize_tiers:
        tier_winners = results_df[results_df["Hadiah"] == tier["name"]].copy()
        if len(tier_winners) == 0:
            continue
        
        tier_winners = tier_winners.sort_values(by="Nomor Undian", ascending=True).reset_index(drop=True)
        
        cols = 5
        rows_per_slide = 5
        winners_per_slide = cols * rows_per_slide
        total_winners = len(tier_winners)
        num_slides = (total_winners + winners_per_slide - 1) // winners_per_slide
        
        for slide_num in range(num_slides):
            slide_layout = prs.slide_layouts[6]
            slide = prs.slides.add_slide(slide_layout)
            
            background = slide.shapes.add_shape(1, Inches(0), Inches(0), prs.slide_width, prs.slide_height)
            background.fill.gradient()
            background.fill.gradient_stops[0].color.rgb = RGBColor(245, 87, 108)
            background.fill.gradient_stops[1].color.rgb = RGBColor(240, 147, 251)
            background.line.fill.background()
            
            title_text = f"{tier['icon']} {tier['name']}"
            if num_slides > 1:
                title_text += f" ({slide_num + 1}/{num_slides})"
            
            title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(12.33), Inches(0.8))
            tf = title_box.text_frame
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = title_text
            run.font.size = Pt(32)
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
            
            cell_width = Inches(2.4)
            cell_height = Inches(1.1)
            gap_x = Inches(0.1)
            gap_y = Inches(0.1)
            
            total_grid_width = cols * cell_width + (cols - 1) * gap_x
            start_x = (prs.slide_width - total_grid_width) / 2
            start_y = Inches(1.3)
            
            start_idx = slide_num * winners_per_slide
            end_idx = min(start_idx + winners_per_slide, total_winners)
            
            for idx, (_, row) in enumerate(tier_winners.iloc[start_idx:end_idx].iterrows()):
                row_num = idx // cols
                col_num = idx % cols
                
                left = start_x + col_num * (cell_width + gap_x)
                top = start_y + row_num * (cell_height + gap_y)
                
                shape = slide.shapes.add_shape(5, left, top, cell_width, cell_height)
                shape.fill.solid()
                shape.fill.fore_color.rgb = RGBColor(255, 255, 255)
                shape.line.color.rgb = RGBColor(245, 87, 108)
                shape.line.width = Pt(2)
                
                nomor = str(row["Nomor Undian"])
                nama_raw = row.get("Nama", "")
                nama = str(nama_raw) if pd.notna(nama_raw) else "-"
                if nama.lower() == "nan":
                    nama = "-"
                nama = nama[:12] if len(nama) > 12 else nama
                hp_raw = row.get("No HP", "")
                hp = mask_phone(hp_raw)
                
                tf = shape.text_frame
                tf.word_wrap = False
                p = tf.paragraphs[0]
                p.alignment = PP_ALIGN.CENTER
                p.space_before = Pt(4)
                p.space_after = Pt(0)
                run = p.add_run()
                run.text = nomor
                run.font.size = Pt(24)
                run.font.bold = True
                run.font.color.rgb = RGBColor(51, 51, 51)
                
                p2 = tf.add_paragraph()
                p2.alignment = PP_ALIGN.CENTER
                p2.space_before = Pt(2)
                p2.space_after = Pt(0)
                run2 = p2.add_run()
                run2.text = nama
                run2.font.size = Pt(12)
                run2.font.color.rgb = RGBColor(102, 102, 102)
                
                p3 = tf.add_paragraph()
                p3.alignment = PP_ALIGN.CENTER
                p3.space_before = Pt(0)
                run3 = p3.add_run()
                run3.text = hp
                run3.font.size = Pt(11)
                run3.font.color.rgb = RGBColor(136, 136, 136)
    
    pptx_buffer = BytesIO()
    prs.save(pptx_buffer)
    pptx_buffer.seek(0)
    return pptx_buffer.getvalue()

def generate_shuffle_pptx(winners_list, prize_name, name_lookup=None, phone_lookup=None):
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    
    if name_lookup is None:
        name_lookup = {}
    if phone_lookup is None:
        phone_lookup = {}
    
    sorted_winners = sorted(winners_list, key=lambda x: str(x))
    
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    
    background = slide.shapes.add_shape(1, Inches(0), Inches(0), prs.slide_width, prs.slide_height)
    background.fill.gradient()
    background.fill.gradient_stops[0].color.rgb = RGBColor(255, 152, 0)
    background.fill.gradient_stops[1].color.rgb = RGBColor(255, 87, 34)
    background.line.fill.background()
    
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.33), Inches(0.8))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = f"🎲 {prize_name}"
    run.font.size = Pt(36)
    run.font.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)
    
    cols = 5
    cell_width = Inches(2.4)
    cell_height = Inches(1.1)
    gap_x = Inches(0.1)
    gap_y = Inches(0.1)
    
    total_grid_width = cols * cell_width + (cols - 1) * gap_x
    start_x = (prs.slide_width - total_grid_width) / 2
    start_y = Inches(1.3)
    
    for idx, winner in enumerate(sorted_winners):
        row_num = idx // cols
        col_num = idx % cols
        
        left = start_x + col_num * (cell_width + gap_x)
        top = start_y + row_num * (cell_height + gap_y)
        
        shape = slide.shapes.add_shape(5, left, top, cell_width, cell_height)
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(255, 255, 255)
        shape.line.color.rgb = RGBColor(255, 152, 0)
        shape.line.width = Pt(2)
        
        nomor = str(winner)
        nama_raw = name_lookup.get(winner, "")
        nama = str(nama_raw) if pd.notna(nama_raw) else "-"
        if nama.lower() == "nan":
            nama = "-"
        nama = nama[:12] if len(nama) > 12 else nama
        hp = mask_phone(phone_lookup.get(winner, ""))
        
        tf = shape.text_frame
        tf.word_wrap = False
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.space_before = Pt(4)
        p.space_after = Pt(0)
        run = p.add_run()
        run.text = nomor
        run.font.size = Pt(24)
        run.font.bold = True
        run.font.color.rgb = RGBColor(51, 51, 51)
        
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        p2.space_before = Pt(2)
        p2.space_after = Pt(0)
        run2 = p2.add_run()
        run2.text = nama
        run2.font.size = Pt(12)
        run2.font.color.rgb = RGBColor(102, 102, 102)
        
        p3 = tf.add_paragraph()
        p3.alignment = PP_ALIGN.CENTER
        p3.space_before = Pt(0)
        run3 = p3.add_run()
        run3.text = hp
        run3.font.size = Pt(11)
        run3.font.color.rgb = RGBColor(136, 136, 136)
    
    pptx_buffer = BytesIO()
    prs.save(pptx_buffer)
    pptx_buffer.seek(0)
    return pptx_buffer.getvalue()

def generate_wheel_pptx(winners_list, prizes_list, name_lookup=None, phone_lookup=None):
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    
    if name_lookup is None:
        name_lookup = {}
    if phone_lookup is None:
        phone_lookup = {}
    
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    
    background = slide.shapes.add_shape(1, Inches(0), Inches(0), prs.slide_width, prs.slide_height)
    background.fill.gradient()
    background.fill.gradient_stops[0].color.rgb = RGBColor(233, 30, 99)
    background.fill.gradient_stops[1].color.rgb = RGBColor(156, 39, 176)
    background.line.fill.background()
    
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.33), Inches(1))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = "🎡 PEMENANG SPINNING WHEEL"
    run.font.size = Pt(36)
    run.font.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)
    
    start_y = Inches(1.5)
    
    for idx, (winner, prize) in enumerate(zip(winners_list, prizes_list)):
        nama_raw = name_lookup.get(winner, "")
        nama = str(nama_raw) if pd.notna(nama_raw) else "-"
        if nama.lower() == "nan":
            nama = "-"
        hp = mask_phone(phone_lookup.get(winner, ""))
        
        text_box = slide.shapes.add_textbox(Inches(0.5), start_y + Inches(idx * 0.55), Inches(12.33), Inches(0.5))
        tf = text_box.text_frame
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = f"#{idx+1}: {winner} - {nama} ({hp}) | {prize}"
        run.font.size = Pt(22)
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)
    
    pptx_buffer = BytesIO()
    prs.save(pptx_buffer)
    pptx_buffer.seek(0)
    return pptx_buffer.getvalue()

st.set_page_config(page_title="Undian Move & Groove", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
    .main-title { text-align: center; color: white; font-size: 2.5rem; font-weight: 800; margin: 0; }
    .subtitle { text-align: center; color: #ccc; font-size: 1.2rem; margin: 0; }
    .section-header { text-align: center; color: white; font-size: 1.5rem; font-weight: 700; margin: 1rem 0; }
    .stats-card { background: rgba(255,255,255,0.1); border-radius: 15px; padding: 1.5rem; text-align: center; }
    .stats-number { font-size: 2.5rem; font-weight: 800; color: #f5576c; }
    .stats-label { color: #ccc; font-size: 1rem; }
    .prize-header { background: linear-gradient(135deg, #f5576c, #f093fb); padding: 2rem; border-radius: 20px; text-align: center; color: white; margin: 1rem 0; }
    .winner-btn { background: linear-gradient(135deg, #4CAF50, #8BC34A); color: white; padding: 1rem; border-radius: 10px; text-align: center; margin: 0.5rem 0; cursor: pointer; }
</style>
""", unsafe_allow_html=True)

if "current_page" not in st.session_state:
    st.session_state["current_page"] = "home"
if "prize_tiers" not in st.session_state:
    saved_config = load_prize_config()
    st.session_state["prize_tiers"] = saved_config if saved_config else PRIZE_TIERS.copy()

if os.path.exists("attached_assets/Small Banner-01_1764081768006.png"):
    st.image("attached_assets/Small Banner-01_1764081768006.png", use_container_width=True)
st.markdown('<p class="main-title">🎉 UNDIAN MOVE & GROOVE 🎉</p>', unsafe_allow_html=True)
st.markdown('<p class="subtitle">7 Desember 2024</p>', unsafe_allow_html=True)

current_page = st.session_state.get("current_page", "home")

if current_page == "home":
    tab1, tab2 = st.tabs(["📁 Upload File CSV", "🔗 Google Sheets URL"])
    
    df = None
    
    with tab1:
        uploaded_file = st.file_uploader("Upload File CSV", type=["csv"], help="File CSV harus berisi kolom 'Nomor Undian'")
        if uploaded_file:
            try:
                uploaded_file.seek(0)
                file_content = uploaded_file.read()
                uploaded_file.seek(0)
                
                content_hash = hashlib.md5(file_content).hexdigest()
                if st.session_state.get("last_content_hash") != content_hash:
                    st.session_state["last_content_hash"] = content_hash
                    st.session_state["data_source_changed"] = True
                    st.session_state["evoucher_done"] = False
                    st.session_state["evoucher_results"] = None
                    st.session_state["shuffle_results"] = {}
                    st.session_state["shuffle_done"] = False
                    st.session_state["wheel_winners"] = []
                    st.session_state["wheel_prizes"] = []
                    st.session_state["wheel_done"] = False
                    if "sheets_df" in st.session_state:
                        del st.session_state["sheets_df"]
                    if "last_sheets_hash" in st.session_state:
                        del st.session_state["last_sheets_hash"]
                    if "remaining_pool" in st.session_state:
                        del st.session_state["remaining_pool"]
                
                df = pd.read_csv(uploaded_file, dtype=str, encoding='utf-8-sig')
                df.columns = df.columns.str.strip().str.replace('\ufeff', '')
            except Exception as e:
                st.error(f"Error: {e}")
    
    with tab2:
        sheets_url = st.text_input("Paste Google Sheets URL", placeholder="https://docs.google.com/spreadsheets/d/...")
        if sheets_url and st.button("📥 Ambil Data"):
            try:
                sheet_id_match = re.search(r'/d/([a-zA-Z0-9-_]+)', sheets_url)
                if sheet_id_match:
                    sheet_id = sheet_id_match.group(1)
                    csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
                    response = requests.get(csv_url, timeout=30)
                    response.raise_for_status()
                    
                    content_hash = hashlib.md5(response.content).hexdigest()
                    if st.session_state.get("last_sheets_hash") != content_hash:
                        st.session_state["last_sheets_hash"] = content_hash
                        st.session_state["data_source_changed"] = True
                        st.session_state["evoucher_done"] = False
                        st.session_state["evoucher_results"] = None
                        st.session_state["shuffle_results"] = {}
                        st.session_state["shuffle_done"] = False
                        st.session_state["wheel_winners"] = []
                        st.session_state["wheel_prizes"] = []
                        st.session_state["wheel_done"] = False
                        if "last_content_hash" in st.session_state:
                            del st.session_state["last_content_hash"]
                        if "remaining_pool" in st.session_state:
                            del st.session_state["remaining_pool"]
                    
                    df = pd.read_csv(StringIO(response.content.decode('utf-8-sig')), dtype=str)
                    df.columns = df.columns.str.strip().str.replace('\ufeff', '')
                    st.session_state["sheets_df"] = df
                    st.success(f"✅ Berhasil mengambil {len(df)} baris data!")
            except Exception as e:
                st.error(f"Error: {e}")
        
        if df is None and "sheets_df" in st.session_state:
            df = st.session_state["sheets_df"]
    
    if df is not None:
        undian_col = None
        for col in df.columns:
            if "undian" in col.lower():
                undian_col = col
                break
        
        if undian_col is None:
            st.error("❌ File harus memiliki kolom 'Nomor Undian'")
        else:
            df = df.rename(columns={undian_col: "Nomor Undian"})
            
            name_col = None
            for col in df.columns:
                if "nama" in col.lower():
                    name_col = col
                    break
            if name_col and name_col != "Nama":
                df = df.rename(columns={name_col: "Nama"})
            elif "Nama" not in df.columns:
                df["Nama"] = ""
            
            phone_col = None
            for col in df.columns:
                if "hp" in col.lower() or "phone" in col.lower() or "telepon" in col.lower():
                    phone_col = col
                    break
            if phone_col and phone_col != "No HP":
                df = df.rename(columns={phone_col: "No HP"})
            elif "No HP" not in df.columns:
                df["No HP"] = ""
            
            df["Nomor Undian"] = df["Nomor Undian"].astype(str).str.strip().str.zfill(4)
            df = df.dropna(subset=["Nomor Undian"])
            df = df[df["Nomor Undian"].str.len() > 0]
            
            df["Eligible"] = df.apply(lambda x: is_eligible_for_prize(x.get("Nama", ""), x.get("No HP", "")), axis=1)
            
            st.session_state["participant_data"] = df
            eligible_df = df[df["Eligible"] == True]
            st.session_state["eligible_participants"] = eligible_df["Nomor Undian"].tolist()
            
            if "remaining_pool" not in st.session_state or st.session_state.get("data_source_changed", False):
                st.session_state["remaining_pool"] = eligible_df.copy()
                st.session_state["data_source_changed"] = False
            
            total_all = len(df)
            total_eligible = len(eligible_df)
            total_excluded = total_all - total_eligible
            remaining_pool = st.session_state.get("remaining_pool", eligible_df)
            
            st.success(f"✅ Data: {total_all} peserta ({total_eligible} eligible, {total_excluded} VIP/F) | Sisa: {len(remaining_pool)}")
            
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("<p style='text-align:center; color:white; font-size:1.8rem; font-weight:bold;'>🎯 PILIH JENIS UNDIAN</p>", unsafe_allow_html=True)
            st.markdown("<br>", unsafe_allow_html=True)
            
            evoucher_done = st.session_state.get("evoucher_done", False)
            shuffle_done = st.session_state.get("shuffle_done", False)
            wheel_done = st.session_state.get("wheel_done", False)
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                prize_tiers = st.session_state.get("prize_tiers", PRIZE_TIERS)
                total_prizes = calculate_total_winners(prize_tiers)
                status_text = "✅ SELESAI" if evoucher_done else f"{total_prizes} hadiah"
                status_color = "#4CAF50" if evoucher_done else "white"
                st.markdown(f"""
                <div style="background: rgba(76,175,80,0.2); border-radius: 15px; padding: 1.5rem; text-align: center; border: 2px solid #4CAF50; min-height: 200px;">
                    <p style="color: #4CAF50; font-size: 1.8rem; font-weight: bold; margin: 0;">🎁 E-Voucher</p>
                    <p style="color: {status_color}; font-size: 1.1rem; margin: 0.5rem 0;">{status_text}</p>
                    <p style="color: #aaa; font-size: 0.9rem; margin: 0.5rem 0;">
                    4 Kategori Voucher<br>
                    Tokopedia, Indomaret, Bensin, SNL
                    </p>
                </div>
                """, unsafe_allow_html=True)
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("🎁 UNDIAN E-VOUCHER", key="btn_evoucher", use_container_width=True):
                    st.session_state["current_page"] = "evoucher_page"
                    st.rerun()
            
            with col2:
                shuffle_results = st.session_state.get("shuffle_results", {})
                completed_sessions = len(shuffle_results)
                status_text = f"✅ {completed_sessions}/3 Sesi" if completed_sessions > 0 else "3 Sesi x 30 hadiah"
                st.markdown(f"""
                <div style="background: rgba(255,152,0,0.2); border-radius: 15px; padding: 1.5rem; text-align: center; border: 2px solid #FF9800; min-height: 200px;">
                    <p style="color: #FF9800; font-size: 1.8rem; font-weight: bold; margin: 0;">🎲 Shuffle</p>
                    <p style="color: white; font-size: 1.1rem; margin: 0.5rem 0;">{status_text}</p>
                    <p style="color: #aaa; font-size: 0.9rem; margin: 0.5rem 0;">
                    Lucky Draw 3 Sesi<br>
                    Sisa: {len(remaining_pool)} peserta
                    </p>
                </div>
                """, unsafe_allow_html=True)
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("🎲 UNDIAN SHUFFLE", key="btn_shuffle", use_container_width=True):
                    st.session_state["current_page"] = "shuffle_page"
                    st.rerun()
            
            with col3:
                wheel_winners = st.session_state.get("wheel_winners", [])
                status_text = f"✅ {len(wheel_winners)}/10 Hadiah" if len(wheel_winners) > 0 else "10 Hadiah Utama"
                st.markdown(f"""
                <div style="background: rgba(233,30,99,0.2); border-radius: 15px; padding: 1.5rem; text-align: center; border: 2px solid #E91E63; min-height: 200px;">
                    <p style="color: #E91E63; font-size: 1.8rem; font-weight: bold; margin: 0;">🎡 Spinning Wheel</p>
                    <p style="color: white; font-size: 1.1rem; margin: 0.5rem 0;">{status_text}</p>
                    <p style="color: #aaa; font-size: 0.9rem; margin: 0.5rem 0;">
                    Grand Prize satu per satu<br>
                    Sisa: {len(remaining_pool)} peserta
                    </p>
                </div>
                """, unsafe_allow_html=True)
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("🎡 UNDIAN WHEEL", key="btn_wheel", use_container_width=True):
                    st.session_state["current_page"] = "wheel_page"
                    st.rerun()
            
            evoucher_results = st.session_state.get("evoucher_results")
            shuffle_results = st.session_state.get("shuffle_results", {})
            wheel_winners = st.session_state.get("wheel_winners", [])
            
            has_results = (evoucher_results is not None) or (len(shuffle_results) > 0) or (len(wheel_winners) > 0)
            
            if has_results:
                st.markdown("<br>", unsafe_allow_html=True)
                st.markdown("---")
                st.markdown("<p style='text-align:center; color:white; font-size:1.5rem; font-weight:bold;'>🏆 HASIL PEMENANG</p>", unsafe_allow_html=True)
                st.markdown("<br>", unsafe_allow_html=True)
                
                if evoucher_results is not None:
                    st.markdown("<p style='color:#4CAF50; font-size:1.2rem; font-weight:bold;'>🎁 E-Voucher (4 Kategori)</p>", unsafe_allow_html=True)
                    prize_tiers = st.session_state.get("prize_tiers", PRIZE_TIERS)
                    cols = st.columns(4)
                    for idx, tier in enumerate(prize_tiers):
                        with cols[idx]:
                            tier_winners = evoucher_results[evoucher_results["Hadiah"] == tier["name"]]
                            count = len(tier_winners)
                            if st.button(f"{tier['icon']} {tier['name'].split()[0]}\n({count})", key=f"home_ev_{idx}", use_container_width=True):
                                st.session_state["viewing_tier"] = tier
                                st.session_state["current_page"] = "evoucher_category"
                                st.rerun()
                    st.markdown("<br>", unsafe_allow_html=True)
                
                if len(shuffle_results) > 0:
                    st.markdown("<p style='color:#FF9800; font-size:1.2rem; font-weight:bold;'>🎲 Shuffle (3 Sesi)</p>", unsafe_allow_html=True)
                    cols = st.columns(3)
                    for i in range(3):
                        with cols[i]:
                            batch_key = f"shuffle_batch_{i}"
                            if batch_key in shuffle_results:
                                result = shuffle_results[batch_key]
                                prize_name = result.get("prize_name", f"Sesi {i+1}")
                                count = len(result.get("winners", []))
                                if st.button(f"🎲 Sesi {i+1}\n({count})", key=f"home_sh_{i}", use_container_width=True):
                                    st.session_state["viewing_shuffle_batch"] = i
                                    st.session_state["current_page"] = "shuffle_results"
                                    st.rerun()
                            else:
                                st.button(f"🎲 Sesi {i+1}\n(Belum)", key=f"home_sh_{i}", use_container_width=True, disabled=True)
                    st.markdown("<br>", unsafe_allow_html=True)
                
                if len(wheel_winners) > 0:
                    st.markdown("<p style='color:#E91E63; font-size:1.2rem; font-weight:bold;'>🎡 Spinning Wheel</p>", unsafe_allow_html=True)
                    if st.button(f"🎡 Grand Prize ({len(wheel_winners)} pemenang)", key="home_wheel", use_container_width=True):
                        st.session_state["current_page"] = "wheel_results"
                        st.rerun()
            
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("---")
            if st.button("🔄 RESET UNDIAN (Mulai dari Awal)", key="reset_all", use_container_width=True):
                keys_to_keep = ["prize_tiers", "participant_data", "eligible_participants"]
                for key in list(st.session_state.keys()):
                    if key not in keys_to_keep:
                        del st.session_state[key]
                st.session_state["remaining_pool"] = eligible_df.copy()
                st.session_state["current_page"] = "home"
                st.rerun()
    
    else:
        st.markdown("<br>", unsafe_allow_html=True)
        st.info("📁 Silakan upload file CSV atau paste URL Google Sheets untuk memulai undian.")

elif current_page == "evoucher_page":
    prize_tiers = st.session_state.get("prize_tiers", PRIZE_TIERS)
    total_prizes = calculate_total_winners(prize_tiers)
    evoucher_results = st.session_state.get("evoucher_results")
    
    if st.button("⬅️ KEMBALI KE MENU", key="back_to_home"):
        st.session_state["current_page"] = "home"
        st.rerun()
    
    st.markdown(f"""
    <div style="background: linear-gradient(135deg, #4CAF50, #8BC34A); padding: 2rem; border-radius: 15px; text-align: center; margin: 1rem 0;">
        <p style="color: white; font-size: 2.5rem; font-weight: bold; margin: 0;">🎁 UNDIAN E-VOUCHER</p>
        <p style="color: #fff; font-size: 1.2rem; margin: 0.5rem 0;">{total_prizes} Hadiah - {len(prize_tiers)} Kategori</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("<p style='text-align:center; color:white; font-size:1.3rem; font-weight:bold;'>⚙️ KONFIGURASI HADIAH</p>", unsafe_allow_html=True)
    
    with st.expander("📝 Edit Jenis & Jumlah Hadiah", expanded=evoucher_results is None):
        new_tiers = []
        cols = st.columns(4)
        for idx, tier in enumerate(prize_tiers):
            with cols[idx % 4]:
                st.markdown(f"**{tier['icon']} Kategori {idx+1}**")
                name = st.text_input(f"Nama", value=tier["name"], key=f"tier_name_{idx}")
                count = st.number_input(f"Jumlah", value=tier["count"], min_value=1, max_value=500, key=f"tier_count_{idx}")
                icon = st.text_input(f"Icon", value=tier["icon"], key=f"tier_icon_{idx}")
                new_tiers.append({"name": name, "icon": icon, "count": count})
        
        if st.button("💾 Simpan Konfigurasi", use_container_width=True):
            start = 1
            for tier in new_tiers:
                tier["start"] = start
                tier["end"] = start + tier["count"] - 1
                start = tier["end"] + 1
            st.session_state["prize_tiers"] = new_tiers
            save_prize_config(new_tiers)
            st.success("✅ Konfigurasi disimpan!")
            st.rerun()
    
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("<p style='text-align:center; color:white; font-size:1.3rem; font-weight:bold;'>📋 KATEGORI HADIAH</p>", unsafe_allow_html=True)
    
    cols = st.columns(4)
    for idx, tier in enumerate(prize_tiers):
        with cols[idx % 4]:
            st.markdown(f"""
            <div style="background: white; border-radius: 20px; padding: 1.5rem; text-align: center; margin-bottom: 1rem; box-shadow: 0 4px 15px rgba(0,0,0,0.2);">
                <div style="font-size: 3rem; margin-bottom: 0.5rem;">{tier['icon']}</div>
                <p style="color: #333; font-size: 1rem; font-weight: bold; margin: 0;">{tier['name']}</p>
                <p style="color: #f5576c; font-size: 1.2rem; font-weight: bold; margin: 0.3rem 0;">{tier['count']} Pemenang</p>
            </div>
            """, unsafe_allow_html=True)
    
    if evoucher_results is None:
        st.markdown("<br>", unsafe_allow_html=True)
        
        if st.button("🎲 MULAI UNDIAN E-VOUCHER", key="start_evoucher", use_container_width=True):
            eligible_participants = st.session_state.get("eligible_participants", [])
            
            if len(eligible_participants) < total_prizes:
                st.error(f"❌ Peserta eligible ({len(eligible_participants)}) kurang dari total hadiah ({total_prizes})")
            else:
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                for i in range(100):
                    progress_bar.progress(i + 1)
                    if i < 30:
                        status_text.markdown(f"<p style='text-align:center; font-size:1.5rem; color:white;'>🔄 Mengumpulkan data... {i+1}%</p>", unsafe_allow_html=True)
                    elif i < 70:
                        status_text.markdown(f"<p style='text-align:center; font-size:1.5rem; color:white;'>🎲 Mengacak peserta... {i+1}%</p>", unsafe_allow_html=True)
                    else:
                        status_text.markdown(f"<p style='text-align:center; font-size:1.5rem; color:white;'>🏆 Menentukan pemenang... {i+1}%</p>", unsafe_allow_html=True)
                    time.sleep(0.02)
                
                shuffled = secure_shuffle(eligible_participants)
                winners = shuffled[:total_prizes]
                
                participant_data = st.session_state.get("participant_data")
                name_lookup = dict(zip(participant_data["Nomor Undian"], participant_data["Nama"])) if participant_data is not None else {}
                phone_lookup = dict(zip(participant_data["Nomor Undian"], participant_data["No HP"])) if participant_data is not None else {}
                
                results = []
                for i, winner in enumerate(winners, 1):
                    results.append({
                        "Peringkat": i,
                        "Nomor Undian": winner,
                        "Nama": name_lookup.get(winner, ""),
                        "No HP": phone_lookup.get(winner, ""),
                        "Hadiah": get_prize_dynamic(i, prize_tiers)
                    })
                
                results_df = pd.DataFrame(results)
                st.session_state["evoucher_results"] = results_df
                st.session_state["evoucher_done"] = True
                
                remaining = [p for p in eligible_participants if p not in winners]
                remaining_df = participant_data[participant_data["Nomor Undian"].isin(remaining)].copy() if participant_data is not None else pd.DataFrame()
                st.session_state["remaining_pool"] = remaining_df
                
                progress_bar.empty()
                status_text.empty()
                st.balloons()
                st.rerun()
    
    else:
        st.markdown("<br>", unsafe_allow_html=True)
        st.success(f"✅ Undian E-Voucher selesai! {len(evoucher_results)} pemenang")
        
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown('<div class="section-header">🏆 LIHAT PEMENANG PER KATEGORI</div>', unsafe_allow_html=True)
        
        cols = st.columns(4)
        for idx, tier in enumerate(prize_tiers):
            with cols[idx % 4]:
                count = len(evoucher_results[evoucher_results["Hadiah"] == tier["name"]])
                if st.button(f"{tier['icon']} {tier['name']}\n({count} Pemenang)", key=f"view_tier_{idx}", use_container_width=True):
                    st.session_state["viewing_tier"] = tier
                    st.session_state["current_page"] = "evoucher_category"
                    st.rerun()
        
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("---")
        st.markdown("<p style='text-align:center; color:white; font-size:1.3rem; font-weight:600;'>📥 Download Hasil</p>", unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        with col1:
            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                evoucher_results.to_excel(writer, index=False, sheet_name='Hasil Undian')
            st.download_button(
                label="📊 Download Excel (.xlsx)",
                data=excel_buffer.getvalue(),
                file_name="hasil_evoucher.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        
        with col2:
            pptx_data = generate_pptx(evoucher_results, prize_tiers)
            st.download_button(
                label="📽️ Download PowerPoint (.pptx)",
                data=pptx_data,
                file_name="hasil_evoucher.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True
            )
        
        st.markdown("<br>", unsafe_allow_html=True)
        remaining_pool = st.session_state.get("remaining_pool", pd.DataFrame())
        
        with st.expander(f"📋 Nomor yang Belum Diundi ({len(remaining_pool)} peserta)", expanded=False):
            if len(remaining_pool) > 0:
                remaining_numbers = remaining_pool["Nomor Undian"].tolist()
                cols = st.columns(15)
                for idx, num in enumerate(remaining_numbers[:150]):
                    with cols[idx % 15]:
                        st.markdown(f"<div style='background:#333;color:white;padding:0.3rem;border-radius:5px;text-align:center;margin:2px;font-size:0.8rem;'>{num}</div>", unsafe_allow_html=True)
                if len(remaining_numbers) > 150:
                    st.info(f"... dan {len(remaining_numbers) - 150} nomor lainnya")
            else:
                st.info("Semua nomor sudah diundi")
        
        if st.button("📊 SISA NOMOR → KEMBALI KE MENU UTAMA", key="ev_to_home", use_container_width=True):
            st.session_state["current_page"] = "home"
            st.rerun()

elif current_page == "evoucher_category":
    tier = st.session_state.get("viewing_tier")
    results_df = st.session_state.get("evoucher_results")
    
    if tier is None or results_df is None:
        st.session_state["current_page"] = "evoucher_page"
        st.rerun()
    
    if st.button("⬅️ KEMBALI", key="back_to_results"):
        st.session_state["current_page"] = "evoucher_page"
        st.rerun()
    
    tier_winners = results_df[results_df["Hadiah"] == tier["name"]].copy()
    tier_winners = tier_winners.sort_values(by="Nomor Undian", ascending=True).reset_index(drop=True)
    
    st.markdown(f"""
    <div class="prize-header">
        <div style="font-size: 4rem;">{tier["icon"]}</div>
        <div style="font-size: 2.5rem; font-weight: 800;">{tier["name"]}</div>
        <div style="font-size: 1.3rem;">{len(tier_winners)} Pemenang</div>
    </div>
    """, unsafe_allow_html=True)
    
    cols = 7
    rows = (len(tier_winners) + cols - 1) // cols
    
    for row in range(rows):
        row_cols = st.columns(cols)
        for col in range(cols):
            idx = row * cols + col
            if idx < len(tier_winners):
                winner = tier_winners.iloc[idx]
                with row_cols[col]:
                    nomor = winner["Nomor Undian"]
                    nama_raw = winner.get("Nama", "")
                    nama = str(nama_raw) if pd.notna(nama_raw) else ""
                    hp = mask_phone(winner.get("No HP", ""))
                    display_nama = nama[:15] if nama and nama.lower() != "nan" else "-"
                    
                    st.markdown(f"""
                    <div style="background: linear-gradient(145deg, #fff, #f8f9fa); border-radius: 10px; padding: 0.6rem; text-align: center; border-left: 4px solid #f5576c; margin-bottom: 0.4rem; height: 75px; display: flex; flex-direction: column; justify-content: center;">
                        <div style="font-size: 1.1rem; font-weight: 800; color: #333; line-height: 1.3;">{nomor}</div>
                        <div style="font-size: 0.7rem; color: #666; line-height: 1.2; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">{display_nama}</div>
                        <div style="font-size: 0.65rem; color: #888; line-height: 1.2;">{hp}</div>
                    </div>
                    """, unsafe_allow_html=True)

elif current_page == "shuffle_page":
    remaining_pool = st.session_state.get("remaining_pool", pd.DataFrame())
    shuffle_results = st.session_state.get("shuffle_results", {})
    
    if st.button("⬅️ KEMBALI KE MENU", key="back_from_shuffle"):
        st.session_state["current_page"] = "home"
        st.rerun()
    
    st.markdown("""
    <div style="background: linear-gradient(135deg, #FF9800, #FF5722); padding: 2rem; border-radius: 15px; text-align: center; margin: 1rem 0;">
        <p style="color: white; font-size: 2.5rem; font-weight: bold; margin: 0;">🎲 UNDIAN SHUFFLE</p>
        <p style="color: #fff; font-size: 1.2rem; margin: 0.5rem 0;">Lucky Draw 3 Sesi</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown(f"""
    <div style="background: rgba(255,255,255,0.1); border-radius: 10px; padding: 1rem; margin: 1rem 0; text-align: center;">
        <p style="color: white; font-size: 1.2rem; margin: 0;">📊 Sisa Peserta: <strong>{len(remaining_pool)}</strong></p>
    </div>
    """, unsafe_allow_html=True)
    
    shuffle_batches = [
        {"name": "Sesi 1", "count": 30},
        {"name": "Sesi 2", "count": 30},
        {"name": "Sesi 3", "count": 30},
    ]
    
    for i, batch in enumerate(shuffle_batches):
        batch_key = f"shuffle_batch_{i}"
        batch_done = batch_key in shuffle_results
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        with st.expander(f"🎲 {batch['name']} ({batch['count']} pemenang)", expanded=not batch_done):
            if batch_done:
                winners = shuffle_results[batch_key]["winners"]
                prize_name = shuffle_results[batch_key]["prize_name"]
                
                st.success(f"✅ Selesai: {prize_name} - {len(winners)} pemenang")
                
                participant_data = st.session_state.get("participant_data")
                name_lookup = dict(zip(participant_data["Nomor Undian"], participant_data["Nama"])) if participant_data is not None else {}
                phone_lookup = dict(zip(participant_data["Nomor Undian"], participant_data["No HP"])) if participant_data is not None else {}
                
                cols = st.columns(10)
                for idx, w in enumerate(winners):
                    with cols[idx % 10]:
                        nama_raw = name_lookup.get(w, "")
                        nama = str(nama_raw) if pd.notna(nama_raw) else ""
                        display_nama = nama[:10] if nama and nama.lower() != "nan" else "-"
                        hp = mask_phone(phone_lookup.get(w, ""))
                        st.markdown(f"""
                        <div style='background:#4CAF50;color:white;padding:0.5rem;border-radius:8px;text-align:center;margin:2px;'>
                            <div style='font-weight:bold;'>{w}</div>
                            <div style='font-size:0.7rem;'>{display_nama}</div>
                            <div style='font-size:0.65rem;'>{hp}</div>
                        </div>
                        """, unsafe_allow_html=True)
                
                st.markdown("<br>", unsafe_allow_html=True)
                col1, col2 = st.columns(2)
                with col1:
                    df_batch = pd.DataFrame({
                        "Nomor Undian": winners,
                        "Nama": [name_lookup.get(w, "") for w in winners],
                        "No HP": [phone_lookup.get(w, "") for w in winners],
                        "Hadiah": prize_name
                    })
                    excel_buf = BytesIO()
                    with pd.ExcelWriter(excel_buf, engine='openpyxl') as writer:
                        df_batch.to_excel(writer, index=False)
                    st.download_button(f"📊 Download Excel {batch['name']}", excel_buf.getvalue(), f"shuffle_{i+1}.xlsx", use_container_width=True)
                with col2:
                    pptx_data = generate_shuffle_pptx(winners, prize_name, name_lookup, phone_lookup)
                    st.download_button(f"📽️ Download PPT {batch['name']}", pptx_data, f"shuffle_{i+1}.pptx", use_container_width=True)
            else:
                prize_name = st.text_input(f"Nama Hadiah {batch['name']}", placeholder="Contoh: Voucher Rp.500.000", key=f"prize_{batch_key}")
                
                remaining_count = len(remaining_pool)
                max_winners = min(batch['count'], remaining_count)
                
                if prize_name and remaining_count > 0:
                    if st.button(f"🎲 MULAI {batch['name']}", key=f"start_{batch_key}", use_container_width=True):
                        remaining_numbers = remaining_pool["Nomor Undian"].tolist()
                        batch_winners = []
                        temp_pool = remaining_numbers.copy()
                        
                        for _ in range(max_winners):
                            if len(temp_pool) == 0:
                                break
                            idx = secrets.randbelow(len(temp_pool))
                            batch_winners.append(temp_pool.pop(idx))
                        
                        shuffle_results[batch_key] = {
                            "winners": batch_winners,
                            "prize_name": prize_name
                        }
                        st.session_state["shuffle_results"] = shuffle_results
                        
                        new_pool = remaining_pool[~remaining_pool["Nomor Undian"].isin(batch_winners)]
                        st.session_state["remaining_pool"] = new_pool
                        
                        if len(shuffle_results) == 3:
                            st.session_state["shuffle_done"] = True
                        
                        st.rerun()
                elif remaining_count == 0:
                    st.warning("Tidak ada sisa peserta")
                else:
                    st.info("Masukkan nama hadiah terlebih dahulu")
    
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("---")
    
    with st.expander(f"📋 Nomor yang Belum Diundi ({len(remaining_pool)} peserta)", expanded=False):
        if len(remaining_pool) > 0:
            remaining_numbers = remaining_pool["Nomor Undian"].tolist()
            cols = st.columns(15)
            for idx, num in enumerate(remaining_numbers[:150]):
                with cols[idx % 15]:
                    st.markdown(f"<div style='background:#333;color:white;padding:0.3rem;border-radius:5px;text-align:center;margin:2px;font-size:0.8rem;'>{num}</div>", unsafe_allow_html=True)
            if len(remaining_numbers) > 150:
                st.info(f"... dan {len(remaining_numbers) - 150} nomor lainnya")
        else:
            st.info("Semua nomor sudah diundi")
    
    if st.button("📊 SISA NOMOR → KEMBALI KE MENU UTAMA", key="shuffle_done_btn", use_container_width=True):
        st.session_state["current_page"] = "home"
        st.rerun()

elif current_page == "shuffle_results":
    batch_idx = st.session_state.get("viewing_shuffle_batch", 0)
    shuffle_results = st.session_state.get("shuffle_results", {})
    batch_key = f"shuffle_batch_{batch_idx}"
    
    if batch_key not in shuffle_results:
        st.session_state["current_page"] = "home"
        st.rerun()
    
    if st.button("⬅️ KEMBALI", key="back_from_shuffle_results"):
        st.session_state["current_page"] = "home"
        st.rerun()
    
    result = shuffle_results[batch_key]
    winners = result["winners"]
    prize_name = result["prize_name"]
    
    st.markdown(f"""
    <div style="background: linear-gradient(135deg, #FF9800, #FF5722); padding: 2rem; border-radius: 15px; text-align: center; margin: 1rem 0;">
        <p style="color: white; font-size: 2.5rem; font-weight: bold; margin: 0;">🎲 {prize_name}</p>
        <p style="color: #fff; font-size: 1.2rem; margin: 0.5rem 0;">Sesi {batch_idx + 1} - {len(winners)} Pemenang</p>
    </div>
    """, unsafe_allow_html=True)
    
    participant_data = st.session_state.get("participant_data")
    name_lookup = dict(zip(participant_data["Nomor Undian"], participant_data["Nama"])) if participant_data is not None else {}
    phone_lookup = dict(zip(participant_data["Nomor Undian"], participant_data["No HP"])) if participant_data is not None else {}
    
    cols = 10
    rows = (len(winners) + cols - 1) // cols
    
    for row in range(rows):
        row_cols = st.columns(cols)
        for col in range(cols):
            idx = row * cols + col
            if idx < len(winners):
                w = winners[idx]
                with row_cols[col]:
                    nama_raw = name_lookup.get(w, "")
                    nama = str(nama_raw) if pd.notna(nama_raw) else ""
                    display_nama = nama[:15] if nama and nama.lower() != "nan" else "-"
                    hp = mask_phone(phone_lookup.get(w, ""))
                    st.markdown(f"""
                    <div style="background: linear-gradient(145deg, #fff, #f8f9fa); border-radius: 10px; padding: 0.8rem; text-align: center; border-left: 4px solid #FF9800; margin-bottom: 0.5rem;">
                        <div style="font-size: 1.2rem; font-weight: 800; color: #333;">{w}</div>
                        <div style="font-size: 0.7rem; color: #666;">{display_nama}</div>
                        <div style="font-size: 0.7rem; color: #888;">{hp}</div>
                    </div>
                    """, unsafe_allow_html=True)

elif current_page == "wheel_page":
    remaining_pool = st.session_state.get("remaining_pool", pd.DataFrame())
    wheel_winners = st.session_state.get("wheel_winners", [])
    wheel_prizes = st.session_state.get("wheel_prizes", [])
    
    if st.button("⬅️ KEMBALI KE MENU", key="back_from_wheel"):
        st.session_state["current_page"] = "home"
        st.rerun()
    
    st.markdown("""
    <div style="background: linear-gradient(135deg, #E91E63, #9C27B0); padding: 2rem; border-radius: 15px; text-align: center; margin: 1rem 0;">
        <p style="color: white; font-size: 2.5rem; font-weight: bold; margin: 0;">🎡 SPINNING WHEEL</p>
        <p style="color: #fff; font-size: 1.2rem; margin: 0.5rem 0;">10 Hadiah Utama</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown(f"""
    <div style="background: rgba(255,255,255,0.1); border-radius: 10px; padding: 1rem; margin: 1rem 0; text-align: center;">
        <p style="color: white; font-size: 1.2rem; margin: 0;">📊 Sisa Peserta: <strong>{len(remaining_pool)}</strong> | Pemenang: <strong>{len(wheel_winners)}/10</strong></p>
    </div>
    """, unsafe_allow_html=True)
    
    if "wheel_config" not in st.session_state:
        st.session_state["wheel_config"] = [{"name": f"Grand Prize {i+1}", "prize": ""} for i in range(10)]
    
    with st.expander("⚙️ Konfigurasi 10 Hadiah Utama", expanded=len(wheel_winners) == 0):
        wheel_config = st.session_state.get("wheel_config", [])
        new_config = []
        cols = st.columns(5)
        for i in range(10):
            with cols[i % 5]:
                prize = st.text_input(f"Hadiah #{i+1}", value=wheel_config[i]["prize"] if i < len(wheel_config) else "", key=f"wheel_prize_config_{i}", placeholder=f"Grand Prize {i+1}")
                new_config.append({"name": f"Hadiah #{i+1}", "prize": prize})
        st.session_state["wheel_config"] = new_config
    
    current_idx = len(wheel_winners)
    
    if current_idx < 10 and len(remaining_pool) > 0:
        st.markdown(f"<p style='text-align:center; color:white; font-size:1.5rem;'>🎯 Undian Hadiah ke-{current_idx + 1} dari 10</p>", unsafe_allow_html=True)
        
        wheel_config = st.session_state.get("wheel_config", [])
        prize_name = wheel_config[current_idx]["prize"] if current_idx < len(wheel_config) and wheel_config[current_idx]["prize"] else f"Grand Prize {current_idx + 1}"
        
        st.markdown(f"<p style='text-align:center; color:#E91E63; font-size:1.3rem;'>Hadiah: <strong>{prize_name}</strong></p>", unsafe_allow_html=True)
        
        if st.button("🎡 PUTAR WHEEL!", key=f"spin_wheel_{current_idx}", use_container_width=True):
            remaining_numbers = remaining_pool["Nomor Undian"].tolist()
            
            if len(remaining_numbers) > 0:
                winner_idx = secrets.randbelow(len(remaining_numbers))
                winner = remaining_numbers[winner_idx]
                
                wheel_html = create_spinning_wheel_html(remaining_numbers[:20], winner, 400)
                components.html(wheel_html, height=550)
                
                wheel_winners.append(winner)
                wheel_prizes.append(prize_name)
                st.session_state["wheel_winners"] = wheel_winners
                st.session_state["wheel_prizes"] = wheel_prizes
                
                new_pool = remaining_pool[remaining_pool["Nomor Undian"] != winner]
                st.session_state["remaining_pool"] = new_pool
                
                if len(wheel_winners) == 10:
                    st.session_state["wheel_done"] = True
                
                participant_data = st.session_state.get("participant_data")
                if participant_data is not None:
                    winner_row = participant_data[participant_data["Nomor Undian"] == winner]
                    if len(winner_row) > 0:
                        nama_raw = winner_row.iloc[0].get("Nama", "")
                        nama = str(nama_raw) if pd.notna(nama_raw) else "-"
                        if nama.lower() == "nan":
                            nama = "-"
                        hp = mask_phone(winner_row.iloc[0].get("No HP", ""))
                        st.success(f"🎉 Pemenang: {winner} - {nama} ({hp})")
                
                if len(wheel_winners) < 10:
                    if st.button("➡️ Lanjut ke Hadiah Berikutnya", key="next_wheel"):
                        st.rerun()
    
    if len(wheel_winners) > 0:
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("---")
        st.markdown("<p style='text-align:center; color:white; font-size:1.3rem; font-weight:600;'>🏆 Pemenang Wheel</p>", unsafe_allow_html=True)
        
        participant_data = st.session_state.get("participant_data")
        name_lookup = dict(zip(participant_data["Nomor Undian"], participant_data["Nama"])) if participant_data is not None else {}
        phone_lookup = dict(zip(participant_data["Nomor Undian"], participant_data["No HP"])) if participant_data is not None else {}
        
        for i, (w, p) in enumerate(zip(wheel_winners, wheel_prizes)):
            nama_raw = name_lookup.get(w, "")
            nama = str(nama_raw) if pd.notna(nama_raw) else "-"
            if nama.lower() == "nan":
                nama = "-"
            hp = mask_phone(phone_lookup.get(w, ""))
            st.markdown(f"""
            <div style="background: rgba(233,30,99,0.2); border-radius: 10px; padding: 0.8rem; margin: 0.5rem 0;">
                <p style='color:white; text-align:center; margin:0;'>
                    <strong>#{i+1}</strong>: {w} - {nama} ({hp}) | <span style='color:#E91E63;'>{p}</span>
                </p>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1:
            df_wheel = pd.DataFrame({
                "No": range(1, len(wheel_winners) + 1),
                "Nomor Undian": wheel_winners,
                "Nama": [name_lookup.get(w, "") for w in wheel_winners],
                "No HP": [phone_lookup.get(w, "") for w in wheel_winners],
                "Hadiah": wheel_prizes
            })
            excel_buf = BytesIO()
            with pd.ExcelWriter(excel_buf, engine='openpyxl') as writer:
                df_wheel.to_excel(writer, index=False)
            st.download_button("📊 Download Excel Wheel", excel_buf.getvalue(), "wheel_winners.xlsx", use_container_width=True)
        
        with col2:
            pptx_data = generate_wheel_pptx(wheel_winners, wheel_prizes, name_lookup, phone_lookup)
            st.download_button("📽️ Download PPT Wheel", pptx_data, "wheel_winners.pptx", use_container_width=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        with st.expander(f"📋 Nomor yang Belum Diundi ({len(remaining_pool)} peserta)", expanded=False):
            if len(remaining_pool) > 0:
                remaining_numbers = remaining_pool["Nomor Undian"].tolist()
                cols = st.columns(15)
                for idx, num in enumerate(remaining_numbers[:150]):
                    with cols[idx % 15]:
                        st.markdown(f"<div style='background:#333;color:white;padding:0.3rem;border-radius:5px;text-align:center;margin:2px;font-size:0.8rem;'>{num}</div>", unsafe_allow_html=True)
                if len(remaining_numbers) > 150:
                    st.info(f"... dan {len(remaining_numbers) - 150} nomor lainnya")
            else:
                st.info("Semua nomor sudah diundi")
        
        if st.button("📊 SISA NOMOR → KEMBALI KE MENU UTAMA", key="wheel_done_btn", use_container_width=True):
            st.session_state["current_page"] = "home"
            st.rerun()

elif current_page == "wheel_results":
    wheel_winners = st.session_state.get("wheel_winners", [])
    wheel_prizes = st.session_state.get("wheel_prizes", [])
    
    if len(wheel_winners) == 0:
        st.session_state["current_page"] = "home"
        st.rerun()
    
    if st.button("⬅️ KEMBALI", key="back_from_wheel_results"):
        st.session_state["current_page"] = "home"
        st.rerun()
    
    st.markdown("""
    <div style="background: linear-gradient(135deg, #E91E63, #9C27B0); padding: 2rem; border-radius: 15px; text-align: center; margin: 1rem 0;">
        <p style="color: white; font-size: 2.5rem; font-weight: bold; margin: 0;">🎡 PEMENANG SPINNING WHEEL</p>
        <p style="color: #fff; font-size: 1.2rem; margin: 0.5rem 0;">{len(wheel_winners)} Grand Prize Winners</p>
    </div>
    """, unsafe_allow_html=True)
    
    participant_data = st.session_state.get("participant_data")
    name_lookup = dict(zip(participant_data["Nomor Undian"], participant_data["Nama"])) if participant_data is not None else {}
    phone_lookup = dict(zip(participant_data["Nomor Undian"], participant_data["No HP"])) if participant_data is not None else {}
    
    cols = st.columns(5)
    for i, (w, p) in enumerate(zip(wheel_winners, wheel_prizes)):
        with cols[i % 5]:
            nama_raw = name_lookup.get(w, "")
            nama = str(nama_raw) if pd.notna(nama_raw) else ""
            display_nama = nama[:15] if nama and nama.lower() != "nan" else "-"
            hp = mask_phone(phone_lookup.get(w, ""))
            st.markdown(f"""
            <div style="background: linear-gradient(145deg, #fff, #f8f9fa); border-radius: 15px; padding: 1rem; text-align: center; border: 3px solid #E91E63; margin-bottom: 1rem;">
                <div style="font-size: 0.9rem; color: #E91E63; font-weight: bold;">#{i+1}</div>
                <div style="font-size: 1.5rem; font-weight: 800; color: #333;">{w}</div>
                <div style="font-size: 0.85rem; color: #666;">{display_nama}</div>
                <div style="font-size: 0.8rem; color: #888;">{hp}</div>
                <div style="font-size: 0.75rem; color: #E91E63; margin-top: 0.5rem;">{p}</div>
            </div>
            """, unsafe_allow_html=True)
