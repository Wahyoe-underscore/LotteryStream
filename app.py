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

def create_spinning_wheel_html(participants, winner, wheel_size=400):
    """Create HTML/JS for spinning wheel animation"""
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
            .spin-btn {{
                background: linear-gradient(135deg, #f5576c 0%, #f093fb 100%);
                color: white;
                border: none;
                padding: 15px 40px;
                font-size: 1.3rem;
                font-weight: 700;
                border-radius: 50px;
                cursor: pointer;
                margin-top: 30px;
                box-shadow: 0 8px 25px rgba(245, 87, 108, 0.4);
            }}
            .spin-btn:disabled {{ opacity: 0.6; cursor: not-allowed; }}
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
        <button class="spin-btn" id="spinBtn" onclick="spin()">🎡 PUTAR!</button>
        <div class="winner-display" id="winnerDisplay"></div>
        
        <script>
            const segments = [{','.join(segments_js)}];
            const canvas = document.getElementById('wheel');
            const ctx = canvas.getContext('2d');
            const centerX = canvas.width / 2;
            const centerY = canvas.height / 2;
            const radius = canvas.width / 2 - 5;
            
            let currentRotation = 0;
            let isSpinning = false;
            
            function drawWheel() {{
                const numSegments = segments.length;
                const arc = (2 * Math.PI) / numSegments;
                
                ctx.clearRect(0, 0, canvas.width, canvas.height);
                ctx.save();
                ctx.translate(centerX, centerY);
                ctx.rotate(currentRotation * Math.PI / 180);
                ctx.translate(-centerX, -centerY);
                
                for (let i = 0; i < numSegments; i++) {{
                    const startAngle = i * arc - Math.PI / 2;
                    const endAngle = startAngle + arc;
                    
                    ctx.beginPath();
                    ctx.moveTo(centerX, centerY);
                    ctx.arc(centerX, centerY, radius, startAngle, endAngle);
                    ctx.closePath();
                    ctx.fillStyle = segments[i].color;
                    ctx.fill();
                    ctx.strokeStyle = '#fff';
                    ctx.lineWidth = 2;
                    ctx.stroke();
                    
                    ctx.save();
                    ctx.translate(centerX, centerY);
                    ctx.rotate(startAngle + arc / 2 + Math.PI / 2);
                    ctx.textAlign = 'center';
                    ctx.fillStyle = '#fff';
                    ctx.font = 'bold 14px Poppins, sans-serif';
                    ctx.shadowColor = 'rgba(0,0,0,0.5)';
                    ctx.shadowBlur = 3;
                    ctx.fillText(segments[i].label, 0, -radius + 40);
                    ctx.restore();
                }}
                
                ctx.restore();
            }}
            
            function spin() {{
                if (isSpinning) return;
                isSpinning = true;
                document.getElementById('spinBtn').disabled = true;
                
                const targetRotation = {target_rotation};
                const duration = 5000;
                const startTime = performance.now();
                const startRotation = currentRotation;
                
                function animate(currentTime) {{
                    const elapsed = currentTime - startTime;
                    const progress = Math.min(elapsed / duration, 1);
                    
                    const easeOut = 1 - Math.pow(1 - progress, 4);
                    currentRotation = startRotation + (targetRotation - startRotation) * easeOut;
                    
                    drawWheel();
                    
                    if (progress < 1) {{
                        requestAnimationFrame(animate);
                    }} else {{
                        isSpinning = false;
                        const winnerDisplay = document.getElementById('winnerDisplay');
                        winnerDisplay.innerHTML = '🎉 PEMENANG: <strong>{winner}</strong> 🎉';
                        winnerDisplay.style.display = 'block';
                    }}
                }}
                
                requestAnimationFrame(animate);
            }}
            
            drawWheel();
        </script>
    </body>
    </html>
    '''
    return html

def create_shuffle_reveal_html(winners, prize_name):
    """Create HTML/JS for shuffle and reveal animation"""
    winners_js = ','.join([f'"{w}"' for w in winners])
    num_winners = len(winners)
    
    html = f'''
    <!DOCTYPE html>
    <html>
    <head>
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;600;700;800&display=swap');
            * {{ margin: 0; padding: 0; box-sizing: border-box; }}
            body {{
                font-family: 'Poppins', sans-serif;
                background: transparent;
                min-height: 100vh;
                display: flex;
                flex-direction: column;
                align-items: center;
                padding: 20px;
            }}
            .title {{
                color: white;
                font-size: 1.8rem;
                font-weight: 700;
                margin-bottom: 20px;
                text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
            }}
            .shuffle-container {{
                display: flex;
                flex-wrap: wrap;
                justify-content: center;
                gap: 10px;
                max-width: 900px;
                margin-bottom: 20px;
            }}
            .number-card {{
                width: 80px;
                height: 80px;
                background: linear-gradient(145deg, #fff, #f0f0f0);
                border-radius: 15px;
                display: flex;
                align-items: center;
                justify-content: center;
                font-size: 1.3rem;
                font-weight: 800;
                color: #333;
                box-shadow: 0 5px 15px rgba(0,0,0,0.2);
                opacity: 0;
                transform: scale(0);
            }}
            .number-card.shuffling {{
                animation: shuffle 0.1s infinite;
                opacity: 1;
                transform: scale(1);
            }}
            .number-card.revealed {{
                opacity: 1;
                transform: scale(1);
                background: linear-gradient(145deg, #f5576c, #f093fb);
                color: white;
                animation: popIn 0.4s ease;
            }}
            @keyframes shuffle {{
                0%, 100% {{ transform: scale(1); }}
                50% {{ transform: scale(0.95); }}
            }}
            @keyframes popIn {{
                0% {{ transform: scale(0) rotate(-10deg); opacity: 0; }}
                50% {{ transform: scale(1.1) rotate(5deg); }}
                100% {{ transform: scale(1) rotate(0); opacity: 1; }}
            }}
            .shuffle-btn {{
                background: linear-gradient(135deg, #f5576c 0%, #f093fb 100%);
                color: white;
                border: none;
                padding: 15px 50px;
                font-size: 1.3rem;
                font-weight: 700;
                border-radius: 50px;
                cursor: pointer;
                box-shadow: 0 8px 25px rgba(245, 87, 108, 0.4);
            }}
            .shuffle-btn:disabled {{ opacity: 0.6; cursor: not-allowed; }}
            .status {{
                color: white;
                font-size: 1.2rem;
                margin-top: 15px;
                font-weight: 600;
            }}
        </style>
    </head>
    <body>
        <div class="title">🎁 {prize_name} - {num_winners} Pemenang</div>
        <div class="shuffle-container" id="container"></div>
        <button class="shuffle-btn" id="shuffleBtn" onclick="startShuffle()">🎲 ACAK & TAMPILKAN!</button>
        <div class="status" id="status"></div>
        
        <script>
            const winners = [{winners_js}];
            const container = document.getElementById('container');
            const statusEl = document.getElementById('status');
            let cards = [];
            
            for (let i = 0; i < winners.length; i++) {{
                const card = document.createElement('div');
                card.className = 'number-card';
                card.textContent = '????';
                container.appendChild(card);
                cards.push(card);
            }}
            
            async function startShuffle() {{
                document.getElementById('shuffleBtn').disabled = true;
                
                const randomNums = [];
                for (let i = 0; i < winners.length; i++) {{
                    randomNums.push(String(Math.floor(Math.random() * 9000) + 1000).padStart(4, '0'));
                }}
                
                cards.forEach((card, i) => {{
                    card.classList.add('shuffling');
                    card.textContent = randomNums[i];
                }});
                
                statusEl.textContent = '🔄 Mengacak nomor...';
                
                const shuffleInterval = setInterval(() => {{
                    cards.forEach(card => {{
                        card.textContent = String(Math.floor(Math.random() * 9000) + 1000).padStart(4, '0');
                    }});
                }}, 50);
                
                await new Promise(r => setTimeout(r, 2000));
                clearInterval(shuffleInterval);
                
                statusEl.textContent = '🎉 Menampilkan pemenang...';
                
                for (let i = 0; i < winners.length; i++) {{
                    await new Promise(r => setTimeout(r, 100));
                    cards[i].classList.remove('shuffling');
                    cards[i].classList.add('revealed');
                    cards[i].textContent = winners[i];
                }}
                
                statusEl.textContent = '✅ Selesai! ' + winners.length + ' pemenang terpilih!';
            }}
        </script>
    </body>
    </html>
    '''
    return html

def save_prize_config(prize_tiers):
    try:
        with open(PRIZE_CONFIG_FILE, 'w') as f:
            json.dump(prize_tiers, f, indent=2)
        return True
    except Exception as e:
        st.error(f"Gagal menyimpan konfigurasi: {e}")
        return False

def load_prize_config():
    try:
        if os.path.exists(PRIZE_CONFIG_FILE):
            with open(PRIZE_CONFIG_FILE, 'r') as f:
                return json.load(f)
    except Exception as e:
        pass
    return None

st.set_page_config(
    page_title="Sistem Undian Move & Groove",
    page_icon="🎉",
    layout="wide"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;600;700;800&display=swap');
    
    .stApp {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 50%, #f093fb 100%);
        font-family: 'Poppins', sans-serif;
    }
    
    .main-title {
        text-align: center;
        font-size: 3.5rem;
        font-weight: 800;
        color: white;
        text-shadow: 3px 3px 6px rgba(0,0,0,0.3);
        margin-bottom: 0.5rem;
    }
    
    .subtitle {
        text-align: center;
        font-size: 1.5rem;
        color: #fff9c4;
        margin-bottom: 2rem;
        font-weight: 600;
    }
    
    .stats-card {
        background: linear-gradient(145deg, #ffffff 0%, #f0f0f0 100%);
        border-radius: 20px;
        padding: 2rem;
        text-align: center;
        box-shadow: 0 10px 40px rgba(0,0,0,0.2);
        margin: 1rem 0;
    }
    
    .stats-number {
        font-size: 4rem;
        font-weight: 800;
        background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
    }
    
    .stats-label {
        font-size: 1.3rem;
        color: #666;
        font-weight: 600;
        margin-top: 0.5rem;
    }
    
    .section-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 1.5rem 2rem;
        border-radius: 15px;
        text-align: center;
        font-size: 2rem;
        font-weight: 700;
        margin: 2rem 0 1.5rem 0;
        box-shadow: 0 8px 30px rgba(102, 126, 234, 0.4);
    }
    
    .stButton > button {
        background: linear-gradient(135deg, #f5576c 0%, #f093fb 100%) !important;
        color: white !important;
        font-size: 1.5rem !important;
        font-weight: 700 !important;
        padding: 1rem 3rem !important;
        border-radius: 50px !important;
        border: none !important;
        box-shadow: 0 8px 30px rgba(245, 87, 108, 0.4) !important;
    }
    
    .stDownloadButton > button {
        background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%) !important;
        color: white !important;
        font-size: 1.3rem !important;
        font-weight: 700 !important;
        padding: 1rem 2rem !important;
        border-radius: 50px !important;
        border: none !important;
    }
    
    .prize-header {
        background: linear-gradient(135deg, #f5576c 0%, #f093fb 100%);
        color: white;
        padding: 2rem;
        border-radius: 20px;
        text-align: center;
        margin-bottom: 1.5rem;
        box-shadow: 0 10px 40px rgba(245, 87, 108, 0.4);
    }
</style>
""", unsafe_allow_html=True)

PRIZE_TIERS = [
    {"name": "Tokopedia Rp.100.000,-", "start": 1, "end": 175, "icon": "🛒", "color": "#4CAF50", "count": 175},
    {"name": "Indomaret Rp.100.000,-", "start": 176, "end": 350, "icon": "🏪", "color": "#FF9800", "count": 175},
    {"name": "Bensin Rp.100.000,-", "start": 351, "end": 525, "icon": "⛽", "color": "#2196F3", "count": 175},
    {"name": "SNL Rp.100.000,-", "start": 526, "end": 700, "icon": "🎁", "color": "#E91E63", "count": 175},
]

def get_prize_dynamic(rank, prize_tiers):
    for tier in prize_tiers:
        if tier["start"] <= rank <= tier["end"]:
            return tier["name"]
    return "Tidak Ada Hadiah"

def calculate_total_winners(prize_tiers):
    return sum(tier["count"] for tier in prize_tiers)

def is_eligible_for_prize(name, phone):
    name_str = str(name).strip().upper() if name and str(name) != "nan" else ""
    phone_str = str(phone).strip().upper() if phone and str(phone) != "nan" else ""
    
    excluded_codes = ["VIP", "F"]
    
    for code in excluded_codes:
        if name_str == code or phone_str == code:
            return False
        if name_str.startswith(code + " ") or name_str.endswith(" " + code):
            return False
        if phone_str.startswith(code + " ") or phone_str.endswith(" " + code):
            return False
    
    return True

def secure_shuffle(participants):
    shuffled = participants.copy()
    n = len(shuffled)
    for i in range(n - 1, 0, -1):
        j = secrets.randbelow(i + 1)
        shuffled[i], shuffled[j] = shuffled[j], shuffled[i]
    return shuffled

def generate_pptx(results_df, prize_tiers, title="Hasil Undian"):
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    
    for tier in prize_tiers:
        tier_winners = results_df[results_df["Hadiah"] == tier["name"]].copy()
        if len(tier_winners) == 0:
            continue
        
        tier_winners = tier_winners.drop_duplicates(subset=["Nomor Undian"])
        tier_winners = tier_winners.sort_values(by="Nomor Undian", ascending=True).reset_index(drop=True)
        
        slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(slide_layout)
        
        background = slide.shapes.add_shape(1, Inches(0), Inches(0), prs.slide_width, prs.slide_height)
        background.fill.gradient()
        background.fill.gradient_stops[0].color.rgb = RGBColor(102, 126, 234)
        background.fill.gradient_stops[1].color.rgb = RGBColor(118, 75, 162)
        background.line.fill.background()
        
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.33), Inches(0.8))
        tf = title_box.text_frame
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = f"{tier['icon']} {tier['name']}"
        run.font.size = Pt(36)
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)
        
        subtitle_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.1), Inches(12.33), Inches(0.4))
        tf_sub = subtitle_box.text_frame
        p_sub = tf_sub.paragraphs[0]
        p_sub.alignment = PP_ALIGN.CENTER
        run_sub = p_sub.add_run()
        run_sub.text = f"{len(tier_winners)} Pemenang"
        run_sub.font.size = Pt(20)
        run_sub.font.color.rgb = RGBColor(255, 255, 255)
        
        cols = 10
        cell_width = Inches(1.15)
        cell_height = Inches(0.6)
        gap_x = Inches(0.08)
        gap_y = Inches(0.08)
        
        total_grid_width = cols * cell_width + (cols - 1) * gap_x
        start_x = (prs.slide_width - total_grid_width) / 2
        start_y = Inches(1.8)
        
        for idx, (_, row) in enumerate(tier_winners.iterrows()):
            if idx >= 50:
                break
            row_num = idx // cols
            col_num = idx % cols
            
            left = start_x + col_num * (cell_width + gap_x)
            top = start_y + row_num * (cell_height + gap_y)
            
            shape = slide.shapes.add_shape(5, left, top, cell_width, cell_height)
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor(255, 255, 255)
            shape.line.color.rgb = RGBColor(245, 87, 108)
            shape.line.width = Pt(2)
            
            shape.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
            shape.text_frame.paragraphs[0].space_before = Pt(0)
            shape.text_frame.paragraphs[0].space_after = Pt(0)
            
            run = shape.text_frame.paragraphs[0].add_run()
            run.text = str(row["Nomor Undian"])
            run.font.size = Pt(16)
            run.font.bold = True
            run.font.color.rgb = RGBColor(51, 51, 51)
    
    pptx_buffer = BytesIO()
    prs.save(pptx_buffer)
    pptx_buffer.seek(0)
    return pptx_buffer.getvalue()

def generate_shuffle_pptx(winners_list, prize_name):
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    
    background = slide.shapes.add_shape(1, Inches(0), Inches(0), prs.slide_width, prs.slide_height)
    background.fill.gradient()
    background.fill.gradient_stops[0].color.rgb = RGBColor(255, 152, 0)
    background.fill.gradient_stops[1].color.rgb = RGBColor(255, 87, 34)
    background.line.fill.background()
    
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(12.33), Inches(1))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = f"🎲 {prize_name}"
    run.font.size = Pt(40)
    run.font.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)
    
    cols = 10
    cell_width = Inches(1.1)
    cell_height = Inches(0.7)
    gap_x = Inches(0.1)
    gap_y = Inches(0.1)
    
    total_grid_width = cols * cell_width + (cols - 1) * gap_x
    start_x = (prs.slide_width - total_grid_width) / 2
    start_y = Inches(2.0)
    
    for idx, winner in enumerate(winners_list):
        row_num = idx // cols
        col_num = idx % cols
        
        left = start_x + col_num * (cell_width + gap_x)
        top = start_y + row_num * (cell_height + gap_y)
        
        shape = slide.shapes.add_shape(5, left, top, cell_width, cell_height)
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(255, 255, 255)
        
        shape.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        run = shape.text_frame.paragraphs[0].add_run()
        run.text = str(winner)
        run.font.size = Pt(18)
        run.font.bold = True
        run.font.color.rgb = RGBColor(51, 51, 51)
    
    pptx_buffer = BytesIO()
    prs.save(pptx_buffer)
    pptx_buffer.seek(0)
    return pptx_buffer.getvalue()

if "current_page" not in st.session_state:
    st.session_state["current_page"] = "home"
if "prize_tiers" not in st.session_state:
    saved_config = load_prize_config()
    st.session_state["prize_tiers"] = saved_config if saved_config else PRIZE_TIERS.copy()

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
                    st.session_state["wheel_winners"] = []
                    st.session_state["wheel_prizes"] = []
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
                        st.session_state["wheel_winners"] = []
                        st.session_state["wheel_prizes"] = []
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
            
            st.success(f"✅ Data berhasil dimuat: {total_all} peserta ({total_eligible} eligible, {total_excluded} VIP/F)")
            
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("<p style='text-align:center; color:white; font-size:1.8rem; font-weight:bold;'>🎯 PILIH JENIS UNDIAN</p>", unsafe_allow_html=True)
            st.markdown("<br>", unsafe_allow_html=True)
            
            evoucher_done = st.session_state.get("evoucher_done", False)
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                prize_tiers = st.session_state.get("prize_tiers", PRIZE_TIERS)
                total_prizes = calculate_total_winners(prize_tiers)
                st.markdown(f"""
                <div style="background: rgba(76,175,80,0.2); border-radius: 15px; padding: 1.5rem; text-align: center; border: 2px solid #4CAF50; min-height: 250px;">
                    <p style="color: #4CAF50; font-size: 1.8rem; font-weight: bold; margin: 0;">🎁 E-Voucher</p>
                    <p style="color: white; font-size: 1.1rem; margin: 0.5rem 0;">{total_prizes} hadiah, {len(prize_tiers)} kategori</p>
                    <p style="color: #aaa; font-size: 0.9rem; margin: 0.5rem 0;">
                    Tokopedia, Indomaret, Bensin, SNL<br>
                    VIP & F tidak diundi
                    </p>
                </div>
                """, unsafe_allow_html=True)
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("🎁 UNDIAN E-VOUCHER", key="btn_evoucher", use_container_width=True):
                    st.session_state["current_page"] = "evoucher_preview"
                    st.rerun()
            
            with col2:
                remaining_pool = st.session_state.get("remaining_pool", eligible_df)
                remaining_count = len(remaining_pool)
                st.markdown(f"""
                <div style="background: rgba(255,152,0,0.2); border-radius: 15px; padding: 1.5rem; text-align: center; border: 2px solid #FF9800; min-height: 250px;">
                    <p style="color: #FF9800; font-size: 1.8rem; font-weight: bold; margin: 0;">🎲 Shuffle</p>
                    <p style="color: white; font-size: 1.1rem; margin: 0.5rem 0;">Lucky Draw 3 Sesi</p>
                    <p style="color: #aaa; font-size: 0.9rem; margin: 0.5rem 0;">
                    Sesi 1, 2, 3 (masing-masing 30 hadiah)<br>
                    Sisa: {remaining_count} peserta
                    </p>
                </div>
                """, unsafe_allow_html=True)
                st.markdown("<br>", unsafe_allow_html=True)
                shuffle_disabled = not evoucher_done
                if st.button("🎲 UNDIAN SHUFFLE", key="btn_shuffle", use_container_width=True, disabled=shuffle_disabled):
                    st.session_state["current_page"] = "shuffle_page"
                    st.rerun()
                if shuffle_disabled:
                    st.caption("⚠️ Selesaikan E-Voucher terlebih dahulu")
            
            with col3:
                st.markdown(f"""
                <div style="background: rgba(233,30,99,0.2); border-radius: 15px; padding: 1.5rem; text-align: center; border: 2px solid #E91E63; min-height: 250px;">
                    <p style="color: #E91E63; font-size: 1.8rem; font-weight: bold; margin: 0;">🎡 Spinning Wheel</p>
                    <p style="color: white; font-size: 1.1rem; margin: 0.5rem 0;">10 Hadiah Utama</p>
                    <p style="color: #aaa; font-size: 0.9rem; margin: 0.5rem 0;">
                    Grand Prize satu per satu<br>
                    Sisa: {remaining_count} peserta
                    </p>
                </div>
                """, unsafe_allow_html=True)
                st.markdown("<br>", unsafe_allow_html=True)
                wheel_disabled = not evoucher_done
                if st.button("🎡 UNDIAN WHEEL", key="btn_wheel", use_container_width=True, disabled=wheel_disabled):
                    st.session_state["current_page"] = "wheel_page"
                    st.rerun()
                if wheel_disabled:
                    st.caption("⚠️ Selesaikan E-Voucher terlebih dahulu")
            
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
        
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("<p style='text-align:center; color:white; font-size:1.5rem; font-weight:bold;'>🎯 3 MODE UNDIAN TERSEDIA</p>", unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown("""
            <div style="background: rgba(76,175,80,0.2); border-radius: 15px; padding: 1.5rem; text-align: center; border: 2px solid #4CAF50;">
                <p style="color: #4CAF50; font-size: 1.5rem; font-weight: bold; margin: 0;">🎁 E-Voucher</p>
                <p style="color: white; font-size: 1rem; margin: 0.5rem 0;">700 hadiah, 4 kategori</p>
                <p style="color: #aaa; font-size: 0.85rem;">Tokopedia, Indomaret, Bensin, SNL</p>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown("""
            <div style="background: rgba(255,152,0,0.2); border-radius: 15px; padding: 1.5rem; text-align: center; border: 2px solid #FF9800;">
                <p style="color: #FF9800; font-size: 1.5rem; font-weight: bold; margin: 0;">🎲 Shuffle</p>
                <p style="color: white; font-size: 1rem; margin: 0.5rem 0;">Lucky Draw 3 Sesi</p>
                <p style="color: #aaa; font-size: 0.85rem;">Masing-masing 30 hadiah</p>
            </div>
            """, unsafe_allow_html=True)
        
        with col3:
            st.markdown("""
            <div style="background: rgba(233,30,99,0.2); border-radius: 15px; padding: 1.5rem; text-align: center; border: 2px solid #E91E63;">
                <p style="color: #E91E63; font-size: 1.5rem; font-weight: bold; margin: 0;">🎡 Spinning Wheel</p>
                <p style="color: white; font-size: 1rem; margin: 0.5rem 0;">10 Hadiah Utama</p>
                <p style="color: #aaa; font-size: 0.85rem;">Grand Prize dramatis</p>
            </div>
            """, unsafe_allow_html=True)

elif current_page == "evoucher_preview":
    prize_tiers = st.session_state.get("prize_tiers", PRIZE_TIERS)
    total_prizes = calculate_total_winners(prize_tiers)
    
    if st.button("⬅️ KEMBALI", key="back_to_home"):
        st.session_state["current_page"] = "home"
        st.rerun()
    
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(f"""
    <div style="background: linear-gradient(135deg, #4CAF50, #8BC34A); padding: 2rem; border-radius: 15px; text-align: center; margin: 1rem 0;">
        <p style="color: white; font-size: 2.5rem; font-weight: bold; margin: 0;">🎁 UNDIAN E-VOUCHER</p>
        <p style="color: #fff; font-size: 1.2rem; margin: 0.5rem 0;">{total_prizes} Hadiah - {len(prize_tiers)} Kategori</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("<p style='text-align:center; color:white; font-size:1.5rem; font-weight:bold;'>📋 KATEGORI HADIAH</p>", unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)
    
    cols = st.columns(4)
    for idx, tier in enumerate(prize_tiers):
        with cols[idx % 4]:
            st.markdown(f"""
            <div style="background: white; border-radius: 20px; padding: 1.5rem; text-align: center; margin-bottom: 1rem; box-shadow: 0 4px 15px rgba(0,0,0,0.2);">
                <div style="font-size: 3rem; margin-bottom: 0.5rem;">{tier['icon']}</div>
                <p style="color: #333; font-size: 1.1rem; font-weight: bold; margin: 0;">{tier['name']}</p>
                <p style="color: #666; font-size: 0.9rem; margin: 0.3rem 0;">Peringkat {tier['start']}-{tier['end']}</p>
                <p style="color: #f5576c; font-size: 1.2rem; font-weight: bold; margin: 0;">{tier['count']} Pemenang</p>
            </div>
            """, unsafe_allow_html=True)
    
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
            
            remaining = [p for p in eligible_participants if p not in winners]
            remaining_df = participant_data[participant_data["Nomor Undian"].isin(remaining)].copy() if participant_data is not None else pd.DataFrame()
            st.session_state["remaining_pool"] = remaining_df
            
            progress_bar.empty()
            status_text.empty()
            st.balloons()
            
            st.session_state["current_page"] = "evoucher_results"
            st.rerun()

elif current_page == "evoucher_results":
    results_df = st.session_state.get("evoucher_results")
    prize_tiers = st.session_state.get("prize_tiers", PRIZE_TIERS)
    
    if results_df is None:
        st.session_state["current_page"] = "home"
        st.rerun()
    
    st.markdown("""
    <div style="background: linear-gradient(135deg, #4CAF50, #8BC34A); padding: 2rem; border-radius: 15px; text-align: center; margin: 1rem 0;">
        <p style="color: white; font-size: 2.5rem; font-weight: bold; margin: 0;">🎉 UNDIAN E-VOUCHER SELESAI! 🎉</p>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown(f"""
        <div class="stats-card">
            <div class="stats-number">{len(st.session_state.get('eligible_participants', [])):,}</div>
            <div class="stats-label">👥 Total Peserta</div>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown(f"""
        <div class="stats-card">
            <div class="stats-number">{len(results_df)}</div>
            <div class="stats-label">🏆 Total Pemenang</div>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        remaining_pool = st.session_state.get("remaining_pool", pd.DataFrame())
        st.markdown(f"""
        <div class="stats-card">
            <div class="stats-number">{len(remaining_pool)}</div>
            <div class="stats-label">📊 Sisa Peserta</div>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="section-header">🏆 PILIH KATEGORI UNTUK LIHAT PEMENANG</div>', unsafe_allow_html=True)
    
    cols = st.columns(4)
    for idx, tier in enumerate(prize_tiers):
        with cols[idx % 4]:
            count = len(results_df[results_df["Hadiah"] == tier["name"]])
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
            results_df.to_excel(writer, index=False, sheet_name='Hasil Undian')
        
        def mark_excel():
            st.session_state["evoucher_excel_done"] = True
        
        st.download_button(
            label="📊 Download Excel (.xlsx)",
            data=excel_buffer.getvalue(),
            file_name="hasil_evoucher.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            on_click=mark_excel
        )
    
    with col2:
        pptx_data = generate_pptx(results_df, prize_tiers)
        
        def mark_pptx():
            st.session_state["evoucher_pptx_done"] = True
        
        st.download_button(
            label="📽️ Download PowerPoint (.pptx)",
            data=pptx_data,
            file_name="hasil_evoucher.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
            on_click=mark_pptx
        )
    
    excel_done = st.session_state.get("evoucher_excel_done", False)
    pptx_done = st.session_state.get("evoucher_pptx_done", False)
    
    if excel_done:
        st.markdown("<p style='text-align:center; color:#90EE90;'>✅ Excel sudah di-download</p>", unsafe_allow_html=True)
    if pptx_done:
        st.markdown("<p style='text-align:center; color:#90EE90;'>✅ PowerPoint sudah di-download</p>", unsafe_allow_html=True)
    
    if excel_done and pptx_done:
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("---")
        
        if st.button("📊 SISA NOMOR → LANJUT KE SHUFFLE/WHEEL", key="continue_to_shuffle", use_container_width=True):
            st.session_state["evoucher_done"] = True
            st.session_state["current_page"] = "home"
            st.rerun()
    else:
        st.markdown("<p style='text-align:center; color:#ffeb3b;'>⚠️ Download Excel dan PPT terlebih dahulu</p>", unsafe_allow_html=True)

elif current_page == "evoucher_category":
    tier = st.session_state.get("viewing_tier")
    results_df = st.session_state.get("evoucher_results")
    
    if tier is None or results_df is None:
        st.session_state["current_page"] = "evoucher_results"
        st.rerun()
    
    if st.button("⬅️ KEMBALI", key="back_to_results"):
        st.session_state["current_page"] = "evoucher_results"
        st.rerun()
    
    tier_winners = results_df[results_df["Hadiah"] == tier["name"]].copy()
    tier_winners = tier_winners.sort_values(by="Nomor Undian", ascending=True).reset_index(drop=True)
    
    st.markdown(f"""
    <div class="prize-header">
        <div style="font-size: 4rem;">{tier["icon"]}</div>
        <div style="font-size: 2.5rem; font-weight: 800;">{tier["name"]}</div>
        <div style="font-size: 1.3rem;">Peringkat {tier["start"]} - {tier["end"]} | {len(tier_winners)} Pemenang</div>
    </div>
    """, unsafe_allow_html=True)
    
    participant_data = st.session_state.get("participant_data")
    
    cols = 10
    rows = (len(tier_winners) + cols - 1) // cols
    
    for row in range(rows):
        row_cols = st.columns(cols)
        for col in range(cols):
            idx = row * cols + col
            if idx < len(tier_winners):
                winner = tier_winners.iloc[idx]
                with row_cols[col]:
                    nomor = winner["Nomor Undian"]
                    nama = winner.get("Nama", "")
                    hp = str(winner.get("No HP", ""))
                    hp_masked = f"****{hp[-4:]}" if len(hp) >= 4 else ""
                    
                    st.markdown(f"""
                    <div style="background: linear-gradient(145deg, #fff, #f8f9fa); border-radius: 10px; padding: 0.8rem; text-align: center; border-left: 4px solid #f5576c; margin-bottom: 0.5rem;">
                        <div style="font-size: 0.75rem; color: #888;">#{winner["Peringkat"]}</div>
                        <div style="font-size: 1.2rem; font-weight: 800; color: #333;">{nomor}</div>
                    </div>
                    """, unsafe_allow_html=True)

elif current_page == "shuffle_page":
    remaining_pool = st.session_state.get("remaining_pool", pd.DataFrame())
    
    if len(remaining_pool) == 0:
        st.warning("Tidak ada sisa peserta untuk diundi")
        if st.button("⬅️ KEMBALI"):
            st.session_state["current_page"] = "home"
            st.rerun()
    else:
        if st.button("⬅️ KEMBALI", key="back_from_shuffle"):
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
            {"name": "Sesi 1 - Hadiah Pertama", "count": 30},
            {"name": "Sesi 2 - Hadiah Kedua", "count": 30},
            {"name": "Sesi 3 - Hadiah Ketiga", "count": 30},
        ]
        
        shuffle_results = st.session_state.get("shuffle_results", {})
        
        for i, batch in enumerate(shuffle_batches):
            batch_key = f"shuffle_batch_{i}"
            batch_done = batch_key in shuffle_results
            
            st.markdown("<br>", unsafe_allow_html=True)
            
            with st.expander(f"🎲 {batch['name']} ({batch['count']} pemenang)", expanded=not batch_done):
                if batch_done:
                    winners = shuffle_results[batch_key]["winners"]
                    prize_name = shuffle_results[batch_key]["prize_name"]
                    
                    st.success(f"✅ Selesai: {prize_name} - {len(winners)} pemenang")
                    
                    cols = st.columns(10)
                    for idx, w in enumerate(winners):
                        with cols[idx % 10]:
                            st.markdown(f"<div style='background:#4CAF50;color:white;padding:0.5rem;border-radius:8px;text-align:center;margin:2px;font-weight:bold;'>{w}</div>", unsafe_allow_html=True)
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        excel_buf = BytesIO()
                        df_batch = pd.DataFrame({"Nomor Undian": winners, "Hadiah": prize_name})
                        with pd.ExcelWriter(excel_buf, engine='openpyxl') as writer:
                            df_batch.to_excel(writer, index=False)
                        st.download_button(f"📊 Download Excel {batch['name']}", excel_buf.getvalue(), f"shuffle_{i+1}.xlsx", use_container_width=True)
                    with col2:
                        pptx_data = generate_shuffle_pptx(winners, prize_name)
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
                            
                            st.rerun()
                    elif remaining_count == 0:
                        st.warning("Tidak ada sisa peserta")
                    else:
                        st.info("Masukkan nama hadiah terlebih dahulu")
        
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("---")
        
        if len(shuffle_results) == 3:
            if st.button("📊 SISA NOMOR → KEMBALI KE MENU UTAMA", key="shuffle_done", use_container_width=True):
                st.session_state["current_page"] = "home"
                st.rerun()

elif current_page == "wheel_page":
    remaining_pool = st.session_state.get("remaining_pool", pd.DataFrame())
    
    if len(remaining_pool) == 0:
        st.warning("Tidak ada sisa peserta untuk diundi")
        if st.button("⬅️ KEMBALI"):
            st.session_state["current_page"] = "home"
            st.rerun()
    else:
        if st.button("⬅️ KEMBALI", key="back_from_wheel"):
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
            <p style="color: white; font-size: 1.2rem; margin: 0;">📊 Sisa Peserta: <strong>{len(remaining_pool)}</strong></p>
        </div>
        """, unsafe_allow_html=True)
        
        wheel_winners = st.session_state.get("wheel_winners", [])
        wheel_prizes = st.session_state.get("wheel_prizes", [])
        current_wheel_idx = len(wheel_winners)
        
        if current_wheel_idx < 10:
            st.markdown(f"<p style='text-align:center; color:white; font-size:1.5rem;'>🎯 Hadiah ke-{current_wheel_idx + 1} dari 10</p>", unsafe_allow_html=True)
            
            prize_name = st.text_input("Nama Hadiah", placeholder="Contoh: Grand Prize - TV 50 inch", key=f"wheel_prize_{current_wheel_idx}")
            
            if prize_name:
                if st.button("🎡 PUTAR WHEEL!", key=f"spin_wheel_{current_wheel_idx}", use_container_width=True):
                    remaining_numbers = remaining_pool["Nomor Undian"].tolist()
                    
                    if len(remaining_numbers) > 0:
                        winner_idx = secrets.randbelow(len(remaining_numbers))
                        winner = remaining_numbers[winner_idx]
                        
                        wheel_html = create_spinning_wheel_html(remaining_numbers[:20], winner, 400)
                        components.html(wheel_html, height=600)
                        
                        wheel_winners.append(winner)
                        wheel_prizes.append(prize_name)
                        st.session_state["wheel_winners"] = wheel_winners
                        st.session_state["wheel_prizes"] = wheel_prizes
                        
                        new_pool = remaining_pool[remaining_pool["Nomor Undian"] != winner]
                        st.session_state["remaining_pool"] = new_pool
                        
                        participant_data = st.session_state.get("participant_data")
                        if participant_data is not None:
                            winner_row = participant_data[participant_data["Nomor Undian"] == winner]
                            if len(winner_row) > 0:
                                nama = winner_row.iloc[0].get("Nama", "")
                                hp = str(winner_row.iloc[0].get("No HP", ""))
                                hp_masked = f"****{hp[-4:]}" if len(hp) >= 4 else ""
                                st.success(f"🎉 Pemenang: {winner} - {nama} ({hp_masked})")
                        
                        if st.button("➡️ Lanjut ke Hadiah Berikutnya", key="next_wheel"):
                            st.rerun()
            else:
                st.info("Masukkan nama hadiah terlebih dahulu")
        
        if len(wheel_winners) > 0:
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("---")
            st.markdown("<p style='text-align:center; color:white; font-size:1.3rem; font-weight:600;'>🏆 Pemenang Wheel</p>", unsafe_allow_html=True)
            
            for i, (w, p) in enumerate(zip(wheel_winners, wheel_prizes)):
                st.markdown(f"<p style='color:white; text-align:center;'>#{i+1}: <strong>{w}</strong> - {p}</p>", unsafe_allow_html=True)
            
            if len(wheel_winners) == 10:
                col1, col2 = st.columns(2)
                with col1:
                    df_wheel = pd.DataFrame({"Nomor Undian": wheel_winners, "Hadiah": wheel_prizes})
                    excel_buf = BytesIO()
                    with pd.ExcelWriter(excel_buf, engine='openpyxl') as writer:
                        df_wheel.to_excel(writer, index=False)
                    st.download_button("📊 Download Excel Wheel", excel_buf.getvalue(), "wheel_winners.xlsx", use_container_width=True)
                
                with col2:
                    if st.button("📊 SISA NOMOR → KEMBALI KE MENU UTAMA", key="wheel_done", use_container_width=True):
                        st.session_state["current_page"] = "home"
                        st.rerun()
