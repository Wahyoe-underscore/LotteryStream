import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import secrets
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
            .wheel {{
                width: 100%;
                height: 100%;
                border-radius: 50%;
                position: relative;
                overflow: hidden;
                box-shadow: 0 0 30px rgba(0,0,0,0.3);
                transition: transform 5s cubic-bezier(0.17, 0.67, 0.12, 0.99);
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
                cursor: pointer;
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
                transition: transform 0.2s;
            }}
            .spin-btn:hover {{ transform: scale(1.05); }}
            .spin-btn:disabled {{ 
                opacity: 0.6; 
                cursor: not-allowed;
                transform: none;
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
    """Create HTML/JS for shuffle and reveal animation for multiple winners"""
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
                transition: all 0.3s ease;
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
                transition: transform 0.2s;
            }}
            .shuffle-btn:hover {{ transform: scale(1.05); }}
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
            
            // Create cards
            for (let i = 0; i < winners.length; i++) {{
                const card = document.createElement('div');
                card.className = 'number-card';
                card.textContent = '????';
                container.appendChild(card);
                cards.push(card);
            }}
            
            async function startShuffle() {{
                document.getElementById('shuffleBtn').disabled = true;
                
                // Start shuffling animation
                const randomNums = [];
                for (let i = 0; i < winners.length; i++) {{
                    randomNums.push(String(Math.floor(Math.random() * 9000) + 1000).padStart(4, '0'));
                }}
                
                cards.forEach((card, i) => {{
                    card.classList.add('shuffling');
                    card.textContent = randomNums[i];
                }});
                
                statusEl.textContent = '🔄 Mengacak nomor...';
                
                // Shuffle for 2 seconds
                const shuffleInterval = setInterval(() => {{
                    cards.forEach(card => {{
                        card.textContent = String(Math.floor(Math.random() * 9000) + 1000).padStart(4, '0');
                    }});
                }}, 50);
                
                await new Promise(r => setTimeout(r, 2000));
                clearInterval(shuffleInterval);
                
                statusEl.textContent = '🎉 Menampilkan pemenang...';
                
                // Reveal one by one
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
    """Save prize configuration to JSON file"""
    try:
        with open(PRIZE_CONFIG_FILE, 'w') as f:
            json.dump(prize_tiers, f, indent=2)
        return True
    except Exception as e:
        st.error(f"Gagal menyimpan konfigurasi: {e}")
        return False

def load_prize_config():
    """Load prize configuration from JSON file"""
    try:
        if os.path.exists(PRIZE_CONFIG_FILE):
            with open(PRIZE_CONFIG_FILE, 'r') as f:
                return json.load(f)
    except Exception as e:
        st.warning(f"Menggunakan konfigurasi default: {e}")
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
        animation: pulse 2s infinite;
    }
    
    @keyframes pulse {
        0%, 100% { transform: scale(1); }
        50% { transform: scale(1.02); }
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
        transition: all 0.3s ease !important;
    }
    
    .stButton > button:hover {
        transform: translateY(-3px) !important;
        box-shadow: 0 12px 40px rgba(245, 87, 108, 0.5) !important;
    }
    
    .stDownloadButton > button {
        background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%) !important;
        color: white !important;
        font-size: 1.3rem !important;
        font-weight: 700 !important;
        padding: 1rem 2rem !important;
        border-radius: 50px !important;
        border: none !important;
        box-shadow: 0 8px 30px rgba(17, 153, 142, 0.4) !important;
    }
    
    .info-box {
        background: rgba(255,255,255,0.9);
        border-radius: 15px;
        padding: 1.5rem;
        margin: 1rem 0;
        border-left: 5px solid #667eea;
    }
    
    .prize-card-clickable {
        background: white;
        border-radius: 20px;
        padding: 1.5rem;
        margin: 0.8rem;
        box-shadow: 0 8px 25px rgba(0,0,0,0.15);
        text-align: center;
        cursor: pointer;
        transition: all 0.3s ease;
        border: 3px solid transparent;
    }
    
    .prize-card-clickable:hover {
        transform: translateY(-8px);
        box-shadow: 0 15px 40px rgba(0,0,0,0.25);
        border-color: #f5576c;
    }
    
    .winner-grid {
        display: grid;
        grid-template-columns: repeat(10, 1fr);
        gap: 8px;
        padding: 1rem;
    }
    
    .winner-cell {
        background: linear-gradient(145deg, #ffffff 0%, #f8f9fa 100%);
        border-radius: 10px;
        padding: 0.8rem 0.5rem;
        text-align: center;
        box-shadow: 0 3px 10px rgba(0,0,0,0.1);
        border-left: 4px solid #f5576c;
    }
    
    .winner-rank {
        font-size: 0.75rem;
        color: #888;
        font-weight: 600;
    }
    
    .winner-number {
        font-size: 1.1rem;
        font-weight: 800;
        color: #333;
        margin-top: 2px;
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
    
    .prize-header-title {
        font-size: 2.5rem;
        font-weight: 800;
        margin-bottom: 0.5rem;
    }
    
    .prize-header-subtitle {
        font-size: 1.3rem;
        opacity: 0.9;
    }
    
    .back-button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%) !important;
    }
    
    .slide-container {
        background: white;
        border-radius: 20px;
        padding: 1.5rem;
        box-shadow: 0 10px 40px rgba(0,0,0,0.2);
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

PRIZE_TIERS = [
    {"name": "Bensin Rp.100.000,-", "start": 1, "end": 75, "icon": "⛽", "color": "#FF6B6B", "count": 75},
    {"name": "Top100 Rp.100.000,-", "start": 76, "end": 175, "icon": "💳", "color": "#4ECDC4", "count": 100},
    {"name": "SNL Rp.100.000,-", "start": 176, "end": 250, "icon": "🎁", "color": "#45B7D1", "count": 75},
    {"name": "Bensin Rp.150.000,-", "start": 251, "end": 325, "icon": "⛽", "color": "#96CEB4", "count": 75},
    {"name": "Top100 Rp.150.000,-", "start": 326, "end": 400, "icon": "💳", "color": "#FFEAA7", "count": 75},
    {"name": "SNL Rp.150.000,-", "start": 401, "end": 500, "icon": "🎁", "color": "#DDA0DD", "count": 100},
    {"name": "Bensin Rp.200.000,-", "start": 501, "end": 600, "icon": "⛽", "color": "#98D8C8", "count": 100},
    {"name": "Top100 Rp.200.000,-", "start": 601, "end": 700, "icon": "💳", "color": "#F7DC6F", "count": 100},
    {"name": "SNL Rp.200.000,-", "start": 701, "end": 800, "icon": "🎁", "color": "#BB8FCE", "count": 100},
]

TOTAL_WINNERS = 800

def get_prize(rank):
    for tier in PRIZE_TIERS:
        if tier["start"] <= rank <= tier["end"]:
            return tier["name"]
    return "Tidak Ada Hadiah"

def get_prize_dynamic(rank, prize_tiers):
    for tier in prize_tiers:
        if tier["start"] <= rank <= tier["end"]:
            return tier["name"]
    return "Tidak Ada Hadiah"

def calculate_total_winners(prize_tiers):
    return sum(tier["count"] for tier in prize_tiers)

def is_eligible_for_prize(name, phone):
    """Check if participant is eligible for prize (not marked as VIP or F)"""
    name_str = str(name).strip().upper() if name and str(name) != "nan" else ""
    phone_str = str(phone).strip().upper() if phone and str(phone) != "nan" else ""
    
    excluded_codes = ["VIP", "F"]
    
    for code in excluded_codes:
        if name_str == code or phone_str == code:
            return False
        if code in name_str or code in phone_str:
            if name_str == code or phone_str == code:
                return False
            if name_str.startswith(code + " ") or name_str.endswith(" " + code) or (" " + code + " ") in name_str:
                return False
            if phone_str.startswith(code + " ") or phone_str.endswith(" " + code) or (" " + code + " ") in phone_str:
                return False
    
    if name_str == "VIP" or name_str == "F" or phone_str == "VIP" or phone_str == "F":
        return False
    
    return True

def secure_shuffle(data_list):
    shuffled = data_list.copy()
    n = len(shuffled)
    for i in range(n - 1, 0, -1):
        j = secrets.randbelow(i + 1)
        shuffled[i], shuffled[j] = shuffled[j], shuffled[i]
    return shuffled

def create_gradient_background(slide, prs):
    background = slide.shapes.add_shape(
        1, Inches(0), Inches(0), prs.slide_width, prs.slide_height
    )
    background.shadow.inherit = False
    fill = background.fill
    fill.gradient()
    fill.gradient_angle = 135
    fill.gradient_stops[0].color.rgb = RGBColor(102, 126, 234)
    fill.gradient_stops[1].color.rgb = RGBColor(240, 147, 251)
    background.line.fill.background()
    
    spTree = slide.shapes._spTree
    sp = background._element
    spTree.remove(sp)
    spTree.insert(2, sp)

def add_winner_cell(slide, left, top, width, height, rank, number):
    shape = slide.shapes.add_shape(1, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(255, 255, 255)
    shape.line.color.rgb = RGBColor(245, 87, 108)
    shape.line.width = Pt(2)
    
    tf = shape.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    
    p1 = tf.paragraphs[0]
    p1.alignment = PP_ALIGN.CENTER
    run1 = p1.add_run()
    run1.text = f"#{rank}"
    run1.font.size = Pt(9)
    run1.font.color.rgb = RGBColor(136, 136, 136)
    run1.font.bold = True
    
    p2 = tf.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    run2 = p2.add_run()
    run2.text = str(number)
    run2.font.size = Pt(18)
    run2.font.color.rgb = RGBColor(51, 51, 51)
    run2.font.bold = True

def add_header(slide, tier, page_num=None, total_pages=None):
    header = slide.shapes.add_shape(
        1, Inches(0.5), Inches(0.3), Inches(12.33), Inches(1.5)
    )
    header.fill.solid()
    header.fill.fore_color.rgb = RGBColor(245, 87, 108)
    header.line.fill.background()
    
    icon_box = slide.shapes.add_textbox(Inches(5.5), Inches(0.35), Inches(2), Inches(0.5))
    tf_icon = icon_box.text_frame
    p_icon = tf_icon.paragraphs[0]
    p_icon.alignment = PP_ALIGN.CENTER
    run_icon = p_icon.add_run()
    run_icon.text = tier["icon"]
    run_icon.font.size = Pt(36)
    
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.85), Inches(12.33), Inches(0.6))
    tf_title = title_box.text_frame
    p_title = tf_title.paragraphs[0]
    p_title.alignment = PP_ALIGN.CENTER
    run_title = p_title.add_run()
    run_title.text = tier["name"]
    run_title.font.size = Pt(32)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(255, 255, 255)
    
    subtitle_text = f"Peringkat {tier['start']} - {tier['end']} | {tier['count']} Pemenang"
    if page_num and total_pages:
        subtitle_text += f" (Slide {page_num}/{total_pages})"
    
    subtitle_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.4), Inches(12.33), Inches(0.4))
    tf_sub = subtitle_box.text_frame
    p_sub = tf_sub.paragraphs[0]
    p_sub.alignment = PP_ALIGN.CENTER
    run_sub = p_sub.add_run()
    run_sub.text = subtitle_text
    run_sub.font.size = Pt(16)
    run_sub.font.color.rgb = RGBColor(255, 255, 255)

def create_winners_slide(prs, tier, winners_data, page_num=None, total_pages=None):
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    
    create_gradient_background(slide, prs)
    add_header(slide, tier, page_num, total_pages)
    
    num_winners = len(winners_data)
    cols = 10
    rows = (num_winners + cols - 1) // cols
    
    cell_width = Inches(1.15)
    cell_height = Inches(0.7)
    gap_x = Inches(0.08)
    gap_y = Inches(0.08)
    
    total_grid_width = cols * cell_width + (cols - 1) * gap_x
    start_x = (prs.slide_width - total_grid_width) / 2
    start_y = Inches(2.0)
    
    for idx, (_, row) in enumerate(winners_data.iterrows()):
        row_num = idx // cols
        col_num = idx % cols
        
        left = start_x + col_num * (cell_width + gap_x)
        top = start_y + row_num * (cell_height + gap_y)
        
        add_winner_cell(slide, left, top, cell_width, cell_height, row["Peringkat"], row["Nomor Undian"])
    
    return slide

def generate_pptx(results_df, prize_tiers=None):
    if prize_tiers is None:
        prize_tiers = PRIZE_TIERS
    
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    
    for tier in prize_tiers:
        tier_winners = results_df[results_df["Hadiah"] == tier["name"]].copy()
        tier_winners = tier_winners.drop_duplicates(subset=["Nomor Undian"])
        tier_winners = tier_winners.sort_values(by="Nomor Undian", ascending=True).reset_index(drop=True)
        
        if len(tier_winners) <= 50:
            create_winners_slide(prs, tier, tier_winners)
        else:
            first_half = tier_winners.iloc[:50]
            second_half = tier_winners.iloc[50:]
            
            create_winners_slide(prs, tier, first_half, 1, 2)
            create_winners_slide(prs, tier, second_half, 2, 2)
    
    pptx_buffer = BytesIO()
    prs.save(pptx_buffer)
    pptx_buffer.seek(0)
    return pptx_buffer.getvalue()

if "selected_prize" not in st.session_state:
    st.session_state["selected_prize"] = None

if "prize_tiers" not in st.session_state:
    saved_config = load_prize_config()
    if saved_config:
        st.session_state["prize_tiers"] = saved_config
    else:
        st.session_state["prize_tiers"] = PRIZE_TIERS.copy()

DEFAULT_ICONS = ["🎁", "⛽", "💳", "🏆", "💰", "🎯", "⭐", "🎊", "🎉", "💎"]

with st.sidebar:
    st.markdown("### ⚙️ Pengaturan Hadiah")
    
    if not st.session_state.get("lottery_done", False):
        st.markdown("---")
        
        if "new_prizes" not in st.session_state:
            st.session_state["new_prizes"] = []
            for tier in st.session_state["prize_tiers"]:
                st.session_state["new_prizes"].append({
                    "name": tier["name"],
                    "count": tier["count"],
                    "icon": tier["icon"]
                })
        
        st.markdown("**Daftar Hadiah:**")
        
        prizes_to_remove = []
        for idx, prize in enumerate(st.session_state["new_prizes"]):
            col1, col2, col3 = st.columns([3, 2, 1])
            with col1:
                new_name = st.text_input(
                    "Nama",
                    value=prize["name"],
                    key=f"prize_name_{idx}",
                    label_visibility="collapsed"
                )
                if new_name != prize["name"]:
                    st.session_state["new_prizes"][idx]["name"] = new_name
            with col2:
                new_count = st.number_input(
                    "Jumlah",
                    min_value=1,
                    value=int(prize["count"]),
                    key=f"prize_count_{idx}",
                    label_visibility="collapsed"
                )
                if new_count != prize["count"]:
                    st.session_state["new_prizes"][idx]["count"] = int(new_count)
            with col3:
                if st.button("🗑️", key=f"remove_{idx}"):
                    prizes_to_remove.append(idx)
        
        for idx in sorted(prizes_to_remove, reverse=True):
            st.session_state["new_prizes"].pop(idx)
            st.rerun()
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button("➕ Tambah Hadiah"):
                icon = DEFAULT_ICONS[len(st.session_state["new_prizes"]) % len(DEFAULT_ICONS)]
                st.session_state["new_prizes"].append({
                    "name": "Hadiah Baru",
                    "count": 10,
                    "icon": icon
                })
                st.rerun()
        
        with col2:
            if st.button("💾 Simpan Hadiah", type="primary"):
                new_tiers = []
                current_start = 1
                for idx in range(len(st.session_state["new_prizes"])):
                    name_key = f"prize_name_{idx}"
                    count_key = f"prize_count_{idx}"
                    name = st.session_state.get(name_key, st.session_state["new_prizes"][idx]["name"])
                    count = int(st.session_state.get(count_key, st.session_state["new_prizes"][idx]["count"]))
                    icon = st.session_state["new_prizes"][idx].get("icon", DEFAULT_ICONS[idx % len(DEFAULT_ICONS)])
                    
                    new_tiers.append({
                        "name": name,
                        "start": current_start,
                        "end": current_start + count - 1,
                        "icon": icon,
                        "color": "#FF6B6B",
                        "count": count
                    })
                    current_start += count
                
                st.session_state["prize_tiers"] = new_tiers
                save_prize_config(new_tiers)
                del st.session_state["new_prizes"]
                st.success("✅ Hadiah tersimpan secara permanen!")
                st.rerun()
        
        total = sum(int(st.session_state.get(f"prize_count_{idx}", p["count"])) for idx, p in enumerate(st.session_state["new_prizes"]))
        st.markdown(f"**Total Pemenang: {total}**")
        
        st.markdown("---")
        st.caption("⚠️ Klik 'Simpan Hadiah' untuk menerapkan perubahan")
    else:
        st.info("Undian sudah selesai. Reset untuk mengubah hadiah.")

if st.session_state.get("continue_lottery_mode", False):
    remaining_pool = st.session_state.get("remaining_pool")
    
    if remaining_pool is not None and len(remaining_pool) > 0:
        st.image("attached_assets/Small Banner-01_1764081768006.png", use_container_width=True)
        
        st.markdown('<p class="main-title">🎡 UNDIAN SISA PESERTA 🎡</p>', unsafe_allow_html=True)
        st.markdown('<p class="subtitle">Spinning Wheel & Shuffle Mode</p>', unsafe_allow_html=True)
        
        all_ws_winners = st.session_state.get("all_wheel_shuffle_winners", [])
        
        st.markdown(f"""
        <div style="background: rgba(255,255,255,0.1); border-radius: 10px; padding: 1rem; margin: 1rem 0;">
            <p style="color: white; text-align: center; margin: 0; font-size: 1.2rem;">
                📊 <strong>Sisa Peserta Eligible:</strong> {len(remaining_pool)} | 
                <strong>Sudah Diundi:</strong> {len(all_ws_winners)}
            </p>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("---")
        st.markdown("<p style='text-align:center; color:white; font-size:1.5rem; font-weight:600;'>🎯 Batch Undian Baru</p>", unsafe_allow_html=True)
        
        col1, col2 = st.columns([1, 2])
        with col1:
            max_winners = len(remaining_pool)
            num_batch_winners = st.number_input(
                "Jumlah pemenang batch ini",
                min_value=1,
                max_value=max_winners,
                value=min(10, max_winners),
                key="batch_winner_count"
            )
        with col2:
            batch_prize_name = st.text_input(
                "Nama hadiah untuk batch ini",
                value="",
                placeholder="Contoh: Voucher Belanja Rp.500.000,-",
                key="batch_prize_name"
            )
        
        if not batch_prize_name:
            st.warning("⚠️ Masukkan nama hadiah terlebih dahulu")
        
        col_w1, col_w2 = st.columns(2)
        with col_w1:
            wheel_disabled = not batch_prize_name
            if st.button("🎡 SPINNING WHEEL", use_container_width=True, type="primary", disabled=wheel_disabled):
                remaining_numbers = remaining_pool["Nomor Undian"].tolist()
                batch_winners = []
                temp_pool = remaining_numbers.copy()
                for _ in range(min(num_batch_winners, len(temp_pool))):
                    idx = secrets.randbelow(len(temp_pool))
                    batch_winners.append(temp_pool.pop(idx))
                
                st.session_state["wheel_mode_active"] = True
                st.session_state["wheel_winners"] = batch_winners
                st.session_state["wheel_current_index"] = 0
                st.session_state["wheel_prize_display"] = batch_prize_name
                st.session_state["wheel_batch_confirmed"] = False
                st.rerun()
        
        with col_w2:
            shuffle_disabled = not batch_prize_name
            if st.button("🎲 SHUFFLE MODE", use_container_width=True, disabled=shuffle_disabled):
                remaining_numbers = remaining_pool["Nomor Undian"].tolist()
                batch_winners = []
                temp_pool = remaining_numbers.copy()
                for _ in range(min(num_batch_winners, len(temp_pool))):
                    idx = secrets.randbelow(len(temp_pool))
                    batch_winners.append(temp_pool.pop(idx))
                
                st.session_state["shuffle_mode_active"] = True
                st.session_state["shuffle_winners"] = batch_winners
                st.session_state["shuffle_prize_display"] = batch_prize_name
                st.session_state["shuffle_batch_confirmed"] = False
                st.rerun()
        
        if st.session_state.get("wheel_mode_active"):
            wheel_winners = st.session_state.get("wheel_winners", [])
            wheel_idx = st.session_state.get("wheel_current_index", 0)
            prize_display = st.session_state.get("wheel_prize_display", "Hadiah")
            participant_data = st.session_state.get("participant_data")
            
            st.markdown("---")
            
            if wheel_idx < len(wheel_winners):
                current_winner = wheel_winners[wheel_idx]
                
                winner_name = ""
                winner_phone = ""
                if participant_data is not None:
                    winner_row = participant_data[participant_data["Nomor Undian"] == current_winner]
                    if len(winner_row) > 0:
                        winner_name = winner_row.iloc[0].get("Nama", "")
                        phone = str(winner_row.iloc[0].get("No HP", ""))
                        if len(phone) >= 4:
                            winner_phone = f"****{phone[-4:]}"
                
                st.markdown(f"""
                <div style="text-align:center; color:white; font-size:1.8rem; margin: 1rem 0;">
                    🎯 {prize_display}
                </div>
                <div style="text-align:center; color:#fff9c4; font-size:1.2rem; margin-bottom: 1rem;">
                    Pemenang ke-{wheel_idx + 1} dari {len(wheel_winners)}
                </div>
                """, unsafe_allow_html=True)
                
                wheel_html = create_spinning_wheel_html(wheel_winners, current_winner, wheel_size=450)
                components.html(wheel_html, height=650)
                
                if winner_name or winner_phone:
                    st.markdown(f"""
                    <div style="text-align:center; background: rgba(255,255,255,0.1); padding: 0.5rem; border-radius: 10px; margin-top: 0.5rem;">
                        <span style="color: white; font-size: 1rem;">👤 {winner_name} | 📱 {winner_phone}</span>
                    </div>
                    """, unsafe_allow_html=True)
                
                col_nav1, col_nav2, col_nav3 = st.columns(3)
                with col_nav1:
                    if wheel_idx > 0:
                        if st.button("⬅️ Sebelumnya", key="wheel_prev_cont", use_container_width=True):
                            st.session_state["wheel_current_index"] = wheel_idx - 1
                            st.rerun()
                with col_nav2:
                    if st.button("❌ Batalkan", key="wheel_cancel_cont", use_container_width=True):
                        st.session_state["wheel_mode_active"] = False
                        st.rerun()
                with col_nav3:
                    if wheel_idx < len(wheel_winners) - 1:
                        if st.button("➡️ Selanjutnya", key="wheel_next_cont", use_container_width=True):
                            st.session_state["wheel_current_index"] = wheel_idx + 1
                            st.rerun()
                    else:
                        if st.button("✅ Konfirmasi Semua", key="wheel_confirm_cont", use_container_width=True, type="primary"):
                            for w in wheel_winners:
                                all_ws_winners.append({"Nomor Undian": w, "Hadiah": prize_display})
                            st.session_state["all_wheel_shuffle_winners"] = all_ws_winners
                            
                            new_pool = remaining_pool[~remaining_pool["Nomor Undian"].isin(wheel_winners)]
                            st.session_state["remaining_pool"] = new_pool
                            st.session_state["wheel_mode_active"] = False
                            st.success(f"✅ {len(wheel_winners)} pemenang dikonfirmasi!")
                            st.rerun()
            
        if st.session_state.get("shuffle_mode_active"):
            shuffle_winners = st.session_state.get("shuffle_winners", [])
            prize_display = st.session_state.get("shuffle_prize_display", "Hadiah")
            
            st.markdown("---")
            
            st.markdown(f"""
            <div style="text-align:center; color:white; font-size:1.8rem; margin: 1rem 0;">
                🎲 {prize_display}
            </div>
            <div style="text-align:center; color:#fff9c4; font-size:1.2rem; margin-bottom: 1rem;">
                {len(shuffle_winners)} Pemenang
            </div>
            """, unsafe_allow_html=True)
            
            shuffle_html = create_shuffle_reveal_html(shuffle_winners, prize_display)
            components.html(shuffle_html, height=500)
            
            col_s1, col_s2 = st.columns(2)
            with col_s1:
                if st.button("❌ Batalkan", key="shuffle_cancel_cont", use_container_width=True):
                    st.session_state["shuffle_mode_active"] = False
                    st.rerun()
            with col_s2:
                if st.button("✅ Konfirmasi Pemenang", key="shuffle_confirm_cont", use_container_width=True, type="primary"):
                    for w in shuffle_winners:
                        all_ws_winners.append({"Nomor Undian": w, "Hadiah": prize_display})
                    st.session_state["all_wheel_shuffle_winners"] = all_ws_winners
                    
                    new_pool = remaining_pool[~remaining_pool["Nomor Undian"].isin(shuffle_winners)]
                    st.session_state["remaining_pool"] = new_pool
                    st.session_state["shuffle_mode_active"] = False
                    st.success(f"✅ {len(shuffle_winners)} pemenang dikonfirmasi!")
                    st.rerun()
        
        if len(all_ws_winners) > 0 and not st.session_state.get("wheel_mode_active") and not st.session_state.get("shuffle_mode_active"):
            st.markdown("---")
            st.markdown("<p style='text-align:center; color:white; font-size:1.3rem; font-weight:600;'>📋 Daftar Pemenang Batch</p>", unsafe_allow_html=True)
            
            ws_df = pd.DataFrame(all_ws_winners)
            participant_data = st.session_state.get("participant_data")
            if participant_data is not None:
                ws_df = ws_df.merge(
                    participant_data[["Nomor Undian", "Nama", "No HP"]], 
                    on="Nomor Undian", 
                    how="left"
                )
                ws_df["No HP Display"] = ws_df["No HP"].apply(lambda x: f"****{str(x)[-4:]}" if pd.notna(x) and len(str(x)) >= 4 else "")
            
            by_prize = ws_df.groupby("Hadiah")
            for prize_name, group in by_prize:
                st.markdown(f"**{prize_name}** ({len(group)} pemenang)")
                display_cols = ["Nomor Undian", "Nama"]
                if "No HP Display" in group.columns:
                    display_cols.append("No HP Display")
                st.dataframe(group[display_cols].reset_index(drop=True), use_container_width=True, hide_index=True)
            
            st.markdown("<br>", unsafe_allow_html=True)
            
            col_exp1, col_exp2 = st.columns(2)
            with col_exp1:
                export_df = ws_df[["Nomor Undian", "Hadiah"]].copy()
                if "Nama" in ws_df.columns:
                    export_df["Nama"] = ws_df["Nama"]
                if "No HP" in ws_df.columns:
                    export_df["No HP"] = ws_df["No HP"]
                
                excel_buffer = BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    export_df.to_excel(writer, index=False, sheet_name='Pemenang Batch')
                excel_data = excel_buffer.getvalue()
                
                st.download_button(
                    label="📊 Download Excel Batch",
                    data=excel_data,
                    file_name="pemenang_batch_wheel_shuffle.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            with col_exp2:
                remaining_csv = remaining_pool[["Nomor Undian", "Nama", "No HP"]].to_csv(index=False)
                st.download_button(
                    label=f"📥 Download Sisa ({len(remaining_pool)} peserta)",
                    data=remaining_csv,
                    file_name="sisa_peserta_eligible.csv",
                    mime="text/csv",
                    use_container_width=True
                )
        
        st.markdown("---")
        col_back1, col_back2 = st.columns(2)
        with col_back1:
            if st.button("⬅️ KEMBALI KE HASIL UNDIAN", use_container_width=True):
                st.session_state["continue_lottery_mode"] = False
                st.session_state["wheel_mode_active"] = False
                st.session_state["shuffle_mode_active"] = False
                st.rerun()
        with col_back2:
            if st.button("🔄 RESET SEMUA (Undian Baru)", use_container_width=True):
                st.session_state["lottery_done"] = False
                st.session_state["results_df"] = None
                st.session_state["continue_lottery_mode"] = False
                st.session_state["wheel_mode_active"] = False
                st.session_state["shuffle_mode_active"] = False
                st.session_state["all_wheel_shuffle_winners"] = []
                st.session_state["excel_downloaded"] = False
                st.session_state["pptx_downloaded"] = False
                st.rerun()
    else:
        st.warning("Tidak ada sisa peserta eligible. Kembali ke halaman utama.")
        if st.button("⬅️ Kembali"):
            st.session_state["continue_lottery_mode"] = False
            st.rerun()

elif st.session_state.get("selected_prize") is not None and st.session_state.get("lottery_done", False):
    selected_tier = st.session_state["selected_prize"]
    results_df = st.session_state["results_df"]
    
    tier_winners = results_df[results_df["Hadiah"] == selected_tier["name"]].copy()
    tier_winners = tier_winners.drop_duplicates(subset=["Nomor Undian"])
    tier_winners = tier_winners.sort_values(by="Nomor Undian", ascending=True).reset_index(drop=True)
    
    col1, col2, col3 = st.columns([1, 6, 1])
    with col1:
        if st.button("⬅️ KEMBALI", use_container_width=True):
            st.session_state["selected_prize"] = None
            st.rerun()
    
    winner_count = selected_tier["end"] - selected_tier["start"] + 1
    st.markdown(f"""
    <div class="prize-header">
        <div style="font-size: 4rem;">{selected_tier["icon"]}</div>
        <div class="prize-header-title">{selected_tier["name"]}</div>
        <div class="prize-header-subtitle">Peringkat {selected_tier["start"]} - {selected_tier["end"]} | {winner_count} Pemenang</div>
    </div>
    """, unsafe_allow_html=True)
    
    num_winners = len(tier_winners)
    
    col_mode1, col_mode2, col_mode3 = st.columns(3)
    with col_mode1:
        grid_mode = st.button("📊 Mode Grid", use_container_width=True, 
                              type="primary" if st.session_state.get("display_mode", "grid") == "grid" else "secondary")
        if grid_mode:
            st.session_state["display_mode"] = "grid"
            st.session_state["wheel_index"] = 0
            st.rerun()
    with col_mode2:
        shuffle_mode = st.button("🎲 Mode Shuffle", use_container_width=True,
                                type="primary" if st.session_state.get("display_mode") == "shuffle" else "secondary")
        if shuffle_mode:
            st.session_state["display_mode"] = "shuffle"
            st.rerun()
    with col_mode3:
        wheel_mode = st.button("🎡 Mode Wheel", use_container_width=True,
                              type="primary" if st.session_state.get("display_mode") == "wheel" else "secondary")
        if wheel_mode:
            st.session_state["display_mode"] = "wheel"
            st.session_state["wheel_index"] = 0
            st.rerun()
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    current_mode = st.session_state.get("display_mode", "grid")
    
    if current_mode == "wheel":
        wheel_index = st.session_state.get("wheel_index", 0)
        
        if wheel_index < num_winners:
            current_winner = tier_winners.iloc[wheel_index]
            winner_number = current_winner["Nomor Undian"]
            
            all_numbers = tier_winners["Nomor Undian"].tolist()
            
            st.markdown(f"""
            <div style="text-align:center; color:white; font-size:1.3rem; margin-bottom:1rem;">
                🎯 Pemenang ke-{wheel_index + 1} dari {num_winners}
            </div>
            """, unsafe_allow_html=True)
            
            wheel_html = create_spinning_wheel_html(all_numbers, winner_number, wheel_size=450)
            components.html(wheel_html, height=650)
            
            col_prev, col_next = st.columns(2)
            with col_prev:
                if wheel_index > 0:
                    if st.button("⬅️ Pemenang Sebelumnya", use_container_width=True):
                        st.session_state["wheel_index"] = wheel_index - 1
                        st.rerun()
            with col_next:
                if wheel_index < num_winners - 1:
                    if st.button("➡️ Pemenang Selanjutnya", use_container_width=True):
                        st.session_state["wheel_index"] = wheel_index + 1
                        st.rerun()
        else:
            st.success("✅ Semua pemenang sudah ditampilkan!")
            if st.button("🔄 Mulai Ulang", use_container_width=True):
                st.session_state["wheel_index"] = 0
                st.rerun()
                
    elif current_mode == "shuffle":
        winner_numbers = tier_winners["Nomor Undian"].tolist()
        shuffle_html = create_shuffle_reveal_html(winner_numbers, selected_tier["name"])
        components.html(shuffle_html, height=600)
        
    else:
        num_cols = 10 if num_winners >= 10 else num_winners
        has_phone_data = "No HP" in tier_winners.columns and tier_winners["No HP"].notna().any()
        has_name_data = "Nama" in tier_winners.columns and tier_winners["Nama"].notna().any()
        cell_height = 120 if (has_phone_data or has_name_data) else 85
        grid_height = ((num_winners + num_cols - 1) // num_cols) * cell_height + 40
        
        has_phone = "No HP" in tier_winners.columns
        has_name = "Nama" in tier_winners.columns
        
        winners_html = f'''
        <html>
        <head>
            <style>
                @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;600;700;800&display=swap');
                body {{
                    margin: 0;
                    padding: 0;
                    font-family: 'Poppins', sans-serif;
                    background: transparent;
                }}
                .winner-grid {{
                    display: grid;
                    grid-template-columns: repeat({num_cols}, 1fr);
                    gap: 8px;
                    padding: 1rem;
                }}
                .winner-cell {{
                    background: linear-gradient(145deg, #ffffff 0%, #f8f9fa 100%);
                    border-radius: 10px;
                    padding: 0.8rem 0.5rem;
                    text-align: center;
                    box-shadow: 0 3px 10px rgba(0,0,0,0.1);
                    border-left: 4px solid #f5576c;
                }}
                .winner-rank {{
                    font-size: 0.75rem;
                    color: #888;
                    font-weight: 600;
                }}
                .winner-number {{
                    font-size: 1.3rem;
                    font-weight: 800;
                    color: #333;
                    margin-top: 0.2rem;
                }}
                .winner-name {{
                    font-size: 0.75rem;
                    color: #555;
                    margin-top: 0.2rem;
                    font-weight: 600;
                    white-space: nowrap;
                    overflow: hidden;
                    text-overflow: ellipsis;
                }}
                .winner-phone {{
                    font-size: 0.7rem;
                    color: #666;
                    margin-top: 0.2rem;
                }}
            </style>
        </head>
        <body>
            <div class="winner-grid">
        '''
        for _, row in tier_winners.iterrows():
            name = row.get("Nama", "") if has_name else ""
            name_html = f'<div class="winner-name">{name}</div>' if name and str(name) != "nan" else ""
            
            phone = row.get("No HP", "") if has_phone else ""
            if phone and str(phone) != "nan" and str(phone).strip():
                phone_display = str(phone).strip()
            else:
                phone_display = ""
            phone_html = f'<div class="winner-phone">📱 {phone_display}</div>' if phone_display else ""
            
            winners_html += f'''
                <div class="winner-cell">
                    <div class="winner-rank">#{row["Peringkat"]}</div>
                    <div class="winner-number">{row["Nomor Undian"]}</div>
                    {name_html}
                    {phone_html}
                </div>
            '''
        winners_html += '''
            </div>
        </body>
        </html>
        '''
        
        components.html(winners_html, height=grid_height, scrolling=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        tier_excel_buffer = BytesIO()
        with pd.ExcelWriter(tier_excel_buffer, engine='openpyxl') as writer:
            tier_winners.to_excel(writer, index=False, sheet_name='Pemenang')
        tier_excel_data = tier_excel_buffer.getvalue()
        
        st.download_button(
            label=f"📥 Download Pemenang {selected_tier['name']}",
            data=tier_excel_data,
            file_name=f"pemenang_{selected_tier['name'].replace(' ', '_').replace('.', '').replace(',', '')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

else:
    st.image("attached_assets/Small Banner-01_1764081768006.png", use_container_width=True)
    
    st.markdown('<p class="main-title">🎉 Sistem Undian Move & Groove 🎉</p>', unsafe_allow_html=True)
    st.markdown('<p class="subtitle">Awareness of Moving The Body for Bone Health</p>', unsafe_allow_html=True)
    
    st.markdown("---")
    
    has_remaining_pool = st.session_state.get("remaining_pool") is not None and len(st.session_state.get("remaining_pool", [])) > 0
    
    if has_remaining_pool and not st.session_state.get("lottery_done", False):
        remaining_pool = st.session_state.get("remaining_pool")
        participant_data = st.session_state.get("participant_data")
        
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, #1a5a1a 0%, #2d7d2d 100%); border-radius: 15px; padding: 1.5rem; margin: 1rem 0; text-align: center;">
            <p style="color: #90EE90; font-size: 1.5rem; font-weight: bold; margin: 0;">📊 Data Sisa Peserta Tersedia</p>
            <p style="color: white; font-size: 1.2rem; margin: 0.5rem 0;">Peserta Eligible: <strong>{len(remaining_pool)}</strong></p>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown("""
            <div style="background: rgba(255,255,255,0.1); border-radius: 15px; padding: 1rem; text-align: center; border: 2px solid #4CAF50; height: 100%;">
                <p style="color: #4CAF50; font-size: 1.3rem; font-weight: bold;">🎁 Undian E-Voucher</p>
                <p style="color: white; font-size: 0.9rem;">700 hadiah, 4 kategori<br>Dikirim lewat WA</p>
            </div>
            """, unsafe_allow_html=True)
            
            st.markdown("<br>", unsafe_allow_html=True)
            
            st.markdown("""
            <p style="color: #aaa; font-size: 0.85rem; text-align: center;">
            • Data: 5000 nomor<br>
            • VIP & F tidak diundi<br>
            • Ada animasi saat undian<br>
            • Download PPT, XLS, Sisa
            </p>
            """, unsafe_allow_html=True)
            
            if st.button("🎲 MULAI UNDIAN UTAMA", key="start_main_lottery", use_container_width=True, type="primary"):
                st.session_state["lottery_mode"] = "main"
                st.rerun()
        
        with col2:
            st.markdown("""
            <div style="background: rgba(255,255,255,0.1); border-radius: 15px; padding: 1rem; text-align: center; border: 2px solid #FF9800; height: 100%;">
                <p style="color: #FF9800; font-size: 1.3rem; font-weight: bold;">🎲 Shuffle</p>
                <p style="color: white; font-size: 0.9rem;">Luckydraw Sesi 1, 2, 3<br>Hadiah per batch</p>
            </div>
            """, unsafe_allow_html=True)
            
            st.markdown("<br>", unsafe_allow_html=True)
            
            st.markdown(f"""
            <p style="color: #aaa; font-size: 0.85rem; text-align: center;">
            • Data: {len(remaining_pool)} sisa peserta<br>
            • Ada nama & nomor HP<br>
            • 3x undian per sesi<br>
            • Update otomatis tiap batch
            </p>
            """, unsafe_allow_html=True)
            
            shuffle_num = st.number_input("Jumlah pemenang", min_value=1, max_value=len(remaining_pool), value=min(50, len(remaining_pool)), key="shuffle_count_home")
            shuffle_prize = st.text_input("Nama hadiah", placeholder="Contoh: Voucher Rp.100.000", key="shuffle_prize_home")
            
            if st.button("🎲 MULAI SHUFFLE", key="start_shuffle_home", use_container_width=True, disabled=not shuffle_prize):
                remaining_numbers = remaining_pool["Nomor Undian"].tolist()
                batch_winners = []
                temp_pool = remaining_numbers.copy()
                for _ in range(min(shuffle_num, len(temp_pool))):
                    idx = secrets.randbelow(len(temp_pool))
                    batch_winners.append(temp_pool.pop(idx))
                
                st.session_state["shuffle_mode_active"] = True
                st.session_state["shuffle_winners"] = batch_winners
                st.session_state["shuffle_prize_display"] = shuffle_prize
                st.session_state["home_shuffle_mode"] = True
                st.rerun()
        
        with col3:
            st.markdown("""
            <div style="background: rgba(255,255,255,0.1); border-radius: 15px; padding: 1rem; text-align: center; border: 2px solid #E91E63; height: 100%;">
                <p style="color: #E91E63; font-size: 1.3rem; font-weight: bold;">🎡 Spinning Wheel</p>
                <p style="color: white; font-size: 0.9rem;">Hadiah Utama<br>10 pemenang</p>
            </div>
            """, unsafe_allow_html=True)
            
            st.markdown("<br>", unsafe_allow_html=True)
            
            st.markdown(f"""
            <p style="color: #aaa; font-size: 0.85rem; text-align: center;">
            • Data: {len(remaining_pool)} sisa peserta<br>
            • Ada nama & nomor HP<br>
            • Diputar per hadiah<br>
            • Efek dramatis
            </p>
            """, unsafe_allow_html=True)
            
            wheel_num = st.number_input("Jumlah hadiah", min_value=1, max_value=min(50, len(remaining_pool)), value=min(10, len(remaining_pool)), key="wheel_count_home")
            wheel_prize = st.text_input("Nama hadiah", placeholder="Contoh: Grand Prize", key="wheel_prize_home")
            
            if st.button("🎡 MULAI WHEEL", key="start_wheel_home", use_container_width=True, type="primary", disabled=not wheel_prize):
                remaining_numbers = remaining_pool["Nomor Undian"].tolist()
                batch_winners = []
                temp_pool = remaining_numbers.copy()
                for _ in range(min(wheel_num, len(temp_pool))):
                    idx = secrets.randbelow(len(temp_pool))
                    batch_winners.append(temp_pool.pop(idx))
                
                st.session_state["wheel_mode_active"] = True
                st.session_state["wheel_winners"] = batch_winners
                st.session_state["wheel_current_index"] = 0
                st.session_state["wheel_prize_display"] = wheel_prize
                st.session_state["home_wheel_mode"] = True
                st.rerun()
        
        if st.session_state.get("shuffle_mode_active") and st.session_state.get("home_shuffle_mode"):
            shuffle_winners = st.session_state.get("shuffle_winners", [])
            prize_display = st.session_state.get("shuffle_prize_display", "Hadiah")
            
            st.markdown("---")
            st.markdown(f"""
            <div style="text-align:center; color:white; font-size:1.8rem; margin: 1rem 0;">
                🎲 {prize_display} - {len(shuffle_winners)} Pemenang
            </div>
            """, unsafe_allow_html=True)
            
            shuffle_html = create_shuffle_reveal_html(shuffle_winners, prize_display)
            components.html(shuffle_html, height=500)
            
            if participant_data is not None:
                st.markdown("**Detail Pemenang:**")
                winner_details = []
                for w in shuffle_winners:
                    row = participant_data[participant_data["Nomor Undian"] == w]
                    if len(row) > 0:
                        nama = row.iloc[0].get("Nama", "")
                        hp = str(row.iloc[0].get("No HP", ""))
                        hp_masked = f"****{hp[-4:]}" if len(hp) >= 4 else ""
                        winner_details.append({"Nomor Undian": w, "Nama": nama, "No HP": hp_masked})
                if winner_details:
                    st.dataframe(pd.DataFrame(winner_details), use_container_width=True, hide_index=True)
            
            col_s1, col_s2 = st.columns(2)
            with col_s1:
                if st.button("❌ Batalkan", key="shuffle_cancel_home", use_container_width=True):
                    st.session_state["shuffle_mode_active"] = False
                    st.session_state["home_shuffle_mode"] = False
                    st.rerun()
            with col_s2:
                if st.button("✅ Konfirmasi & Update Sisa", key="shuffle_confirm_home", use_container_width=True, type="primary"):
                    new_pool = remaining_pool[~remaining_pool["Nomor Undian"].isin(shuffle_winners)]
                    st.session_state["remaining_pool"] = new_pool
                    
                    all_ws = st.session_state.get("all_wheel_shuffle_winners", [])
                    for w in shuffle_winners:
                        all_ws.append({"Nomor Undian": w, "Hadiah": prize_display})
                    st.session_state["all_wheel_shuffle_winners"] = all_ws
                    
                    st.session_state["shuffle_mode_active"] = False
                    st.session_state["home_shuffle_mode"] = False
                    st.success(f"✅ {len(shuffle_winners)} pemenang dikonfirmasi! Sisa peserta: {len(new_pool)}")
                    st.rerun()
        
        if st.session_state.get("wheel_mode_active") and st.session_state.get("home_wheel_mode"):
            wheel_winners = st.session_state.get("wheel_winners", [])
            wheel_idx = st.session_state.get("wheel_current_index", 0)
            prize_display = st.session_state.get("wheel_prize_display", "Hadiah")
            
            st.markdown("---")
            
            if wheel_idx < len(wheel_winners):
                current_winner = wheel_winners[wheel_idx]
                
                winner_name = ""
                winner_phone = ""
                if participant_data is not None:
                    winner_row = participant_data[participant_data["Nomor Undian"] == current_winner]
                    if len(winner_row) > 0:
                        winner_name = winner_row.iloc[0].get("Nama", "")
                        phone = str(winner_row.iloc[0].get("No HP", ""))
                        if len(phone) >= 4:
                            winner_phone = f"****{phone[-4:]}"
                
                st.markdown(f"""
                <div style="text-align:center; color:white; font-size:1.8rem; margin: 1rem 0;">
                    🎯 {prize_display}
                </div>
                <div style="text-align:center; color:#fff9c4; font-size:1.2rem; margin-bottom: 1rem;">
                    Pemenang ke-{wheel_idx + 1} dari {len(wheel_winners)}
                </div>
                """, unsafe_allow_html=True)
                
                wheel_html = create_spinning_wheel_html(wheel_winners, current_winner, wheel_size=450)
                components.html(wheel_html, height=650)
                
                if winner_name or winner_phone:
                    st.markdown(f"""
                    <div style="text-align:center; background: rgba(255,255,255,0.1); padding: 0.5rem; border-radius: 10px; margin-top: 0.5rem;">
                        <span style="color: white; font-size: 1rem;">👤 {winner_name} | 📱 {winner_phone}</span>
                    </div>
                    """, unsafe_allow_html=True)
                
                col_nav1, col_nav2, col_nav3 = st.columns(3)
                with col_nav1:
                    if wheel_idx > 0:
                        if st.button("⬅️ Sebelumnya", key="wheel_prev_home", use_container_width=True):
                            st.session_state["wheel_current_index"] = wheel_idx - 1
                            st.rerun()
                with col_nav2:
                    if st.button("❌ Batalkan", key="wheel_cancel_home", use_container_width=True):
                        st.session_state["wheel_mode_active"] = False
                        st.session_state["home_wheel_mode"] = False
                        st.rerun()
                with col_nav3:
                    if wheel_idx < len(wheel_winners) - 1:
                        if st.button("➡️ Selanjutnya", key="wheel_next_home", use_container_width=True):
                            st.session_state["wheel_current_index"] = wheel_idx + 1
                            st.rerun()
                    else:
                        if st.button("✅ Konfirmasi Semua", key="wheel_confirm_home", use_container_width=True, type="primary"):
                            new_pool = remaining_pool[~remaining_pool["Nomor Undian"].isin(wheel_winners)]
                            st.session_state["remaining_pool"] = new_pool
                            
                            all_ws = st.session_state.get("all_wheel_shuffle_winners", [])
                            for w in wheel_winners:
                                all_ws.append({"Nomor Undian": w, "Hadiah": prize_display})
                            st.session_state["all_wheel_shuffle_winners"] = all_ws
                            
                            st.session_state["wheel_mode_active"] = False
                            st.session_state["home_wheel_mode"] = False
                            st.success(f"✅ {len(wheel_winners)} pemenang dikonfirmasi! Sisa peserta: {len(new_pool)}")
                            st.rerun()
        
        all_ws_winners = st.session_state.get("all_wheel_shuffle_winners", [])
        if len(all_ws_winners) > 0 and not st.session_state.get("wheel_mode_active") and not st.session_state.get("shuffle_mode_active"):
            st.markdown("---")
            st.markdown("<p style='text-align:center; color:white; font-size:1.3rem; font-weight:600;'>📋 Rekap Pemenang Wheel/Shuffle</p>", unsafe_allow_html=True)
            
            ws_df = pd.DataFrame(all_ws_winners)
            if participant_data is not None:
                ws_df = ws_df.merge(
                    participant_data[["Nomor Undian", "Nama", "No HP"]], 
                    on="Nomor Undian", 
                    how="left"
                )
            
            by_prize = ws_df.groupby("Hadiah")
            for prize_name, group in by_prize:
                st.markdown(f"**{prize_name}** ({len(group)} pemenang)")
            
            col_exp1, col_exp2 = st.columns(2)
            with col_exp1:
                excel_buffer = BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    ws_df.to_excel(writer, index=False, sheet_name='Pemenang')
                st.download_button(
                    label="📊 Download Excel Pemenang",
                    data=excel_buffer.getvalue(),
                    file_name="pemenang_wheel_shuffle.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            with col_exp2:
                remaining_csv = remaining_pool[["Nomor Undian", "Nama", "No HP"]].to_csv(index=False)
                st.download_button(
                    label=f"📥 Download Sisa ({len(remaining_pool)})",
                    data=remaining_csv,
                    file_name="sisa_peserta.csv",
                    mime="text/csv",
                    use_container_width=True
                )
        
        st.markdown("---")
        if st.button("🔄 Reset Semua (Upload Data Baru)", use_container_width=True):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()
    
    elif st.session_state.get("lottery_done", False) and st.session_state.get("results_df") is not None:
        results_df = st.session_state["results_df"]
        total_winners = len(results_df)
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown(f"""
            <div class="stats-card">
                <div class="stats-number">{st.session_state.get('total_participants', total_winners):,}</div>
                <div class="stats-label">👥 Total Peserta</div>
            </div>
            """, unsafe_allow_html=True)
        with col2:
            st.markdown(f"""
            <div class="stats-card">
                <div class="stats-number">{total_winners}</div>
                <div class="stats-label">🏆 Total Pemenang</div>
            </div>
            """, unsafe_allow_html=True)
        prize_tiers = st.session_state.get("prize_tiers", PRIZE_TIERS)
        num_categories = len(prize_tiers)
        
        with col3:
            st.markdown(f"""
            <div class="stats-card">
                <div class="stats-number">{num_categories}</div>
                <div class="stats-label">🎁 Kategori Hadiah</div>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown('<div class="section-header">🏆 PILIH KATEGORI HADIAH 🏆</div>', unsafe_allow_html=True)
        st.markdown("<p style='text-align:center; color:white; font-size:1.2rem;'>Klik pada kategori hadiah untuk melihat pemenang dalam satu layar</p>", unsafe_allow_html=True)
        
        cols = st.columns(3)
        for idx, tier in enumerate(prize_tiers):
            col_idx = idx % 3
            with cols[col_idx]:
                count = len(results_df[results_df["Hadiah"] == tier["name"]])
                if st.button(
                    f"{tier['icon']} {tier['name']}\n({count} Pemenang)",
                    key=f"prize_main_{idx}",
                    use_container_width=True
                ):
                    st.session_state["selected_prize"] = tier
                    st.rerun()
        
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("---")
        
        st.markdown("<p style='text-align:center; color:white; font-size:1.3rem; font-weight:600;'>📥 Download Hasil Undian</p>", unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        with col1:
            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                results_df.to_excel(writer, index=False, sheet_name='Hasil Undian')
            excel_data = excel_buffer.getvalue()
            
            def mark_excel_downloaded():
                st.session_state["excel_downloaded"] = True
            
            st.download_button(
                label="📊 Download Excel (.xlsx)",
                data=excel_data,
                file_name="hasil_undian_move_groove.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                on_click=mark_excel_downloaded
            )
        
        with col2:
            def mark_pptx_downloaded():
                st.session_state["pptx_downloaded"] = True
            
            prize_tiers = st.session_state.get("prize_tiers", PRIZE_TIERS)
            pptx_data = generate_pptx(results_df, prize_tiers)
            
            st.download_button(
                label="📽️ Download PowerPoint (.pptx)",
                data=pptx_data,
                file_name="hasil_undian_move_groove.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
                on_click=mark_pptx_downloaded
            )
        
        excel_done = st.session_state.get("excel_downloaded", False)
        pptx_done = st.session_state.get("pptx_downloaded", False)
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        if excel_done:
            st.markdown("<p style='text-align:center; color:#90EE90;'>✅ Excel sudah di-download</p>", unsafe_allow_html=True)
        if pptx_done:
            st.markdown("<p style='text-align:center; color:#90EE90;'>✅ PowerPoint sudah di-download</p>", unsafe_allow_html=True)
        
        if excel_done and pptx_done:
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("---")
            st.markdown("<p style='text-align:center; color:white; font-size:1.3rem; font-weight:600;'>🔄 Opsi Setelah Undian</p>", unsafe_allow_html=True)
            
            participant_data = st.session_state.get("participant_data")
            if participant_data is not None:
                winner_numbers = results_df["Nomor Undian"].tolist()
                remaining_data = participant_data[~participant_data["Nomor Undian"].isin(winner_numbers)].copy()
                remaining_eligible = remaining_data[remaining_data["Eligible"] == True]
                
                st.markdown(f"""
                <div style="background: rgba(255,255,255,0.1); border-radius: 10px; padding: 1rem; margin: 1rem 0;">
                    <p style="color: white; text-align: center; margin: 0;">
                        📊 <strong>Pemenang:</strong> {len(winner_numbers)} | 
                        <strong>Sisa Peserta:</strong> {len(remaining_data)} | 
                        <strong>Sisa Eligible:</strong> {len(remaining_eligible)}
                    </p>
                </div>
                """, unsafe_allow_html=True)
                
                remaining_csv = remaining_data[["Nomor Undian", "Nama", "No HP"]].to_csv(index=False)
                
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.download_button(
                        label="📥 Download Sisa Peserta (CSV)",
                        data=remaining_csv,
                        file_name="sisa_peserta_belum_menang.csv",
                        mime="text/csv",
                        use_container_width=True
                    )
                
                with col2:
                    if st.button("🔄 UNDIAN BARU (Reset)", use_container_width=True):
                        st.session_state["lottery_done"] = False
                        st.session_state["results_df"] = None
                        st.session_state["selected_prize"] = None
                        st.session_state["excel_downloaded"] = False
                        st.session_state["pptx_downloaded"] = False
                        st.rerun()
                
                with col3:
                    if st.button("➡️ LANJUT UNDIAN SISA", use_container_width=True, type="primary"):
                        st.session_state["remaining_pool"] = remaining_eligible.copy()
                        st.session_state["continue_lottery_mode"] = True
                        st.session_state["all_wheel_shuffle_winners"] = []
                        st.rerun()
                
                st.caption("💡 'Download Sisa Peserta' = file baru berisi nomor yang belum menang. Data awal tetap utuh.")
                
        elif not excel_done or not pptx_done:
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("<p style='text-align:center; color:#ffeb3b; font-size:1rem;'>⚠️ Download Excel dan PowerPoint terlebih dahulu sebelum memulai undian baru</p>", unsafe_allow_html=True)
    
    else:
        tab1, tab2 = st.tabs(["📁 Upload File CSV", "🔗 Google Sheets URL"])
        
        df = None
        data_source = None
        
        with tab1:
            uploaded_file = st.file_uploader(
                "Upload File CSV (harus memiliki kolom 'Nomor Undian' dan 'No HP')",
                type=["csv"],
                help="File CSV harus berisi kolom 'Nomor Undian' dan 'No HP'"
            )
            
            if uploaded_file is not None:
                current_file_name = uploaded_file.name
                if st.session_state.get("last_uploaded_file") != current_file_name:
                    st.session_state["last_uploaded_file"] = current_file_name
                    st.session_state["lottery_done"] = False
                    st.session_state["results_df"] = None
                
                try:
                    uploaded_file.seek(0)
                    content = uploaded_file.read()
                    uploaded_file.seek(0)
                    
                    try:
                        first_line = content.decode('utf-8-sig').split('\n')[0]
                    except:
                        first_line = content.decode('utf-8').split('\n')[0]
                    
                    uploaded_file.seek(0)
                    
                    if ';' in first_line:
                        df = pd.read_csv(uploaded_file, dtype=str, sep=';', encoding='utf-8-sig')
                    else:
                        df = pd.read_csv(uploaded_file, dtype=str, encoding='utf-8-sig')
                    
                    df.columns = df.columns.str.strip().str.replace('\ufeff', '')
                    data_source = "csv"
                except Exception as e:
                    st.error(f"❌ Error membaca file: {str(e)}")
        
        with tab2:
            sheets_url = st.text_input(
                "Paste Google Sheets URL",
                placeholder="https://docs.google.com/spreadsheets/d/...",
                help="Pastikan Google Sheets sudah di-share sebagai 'Anyone with the link can view'"
            )
            
            if sheets_url and st.button("📥 Ambil Data dari Google Sheets", use_container_width=True):
                try:
                    sheet_id_match = re.search(r'/d/([a-zA-Z0-9-_]+)', sheets_url)
                    if sheet_id_match:
                        sheet_id = sheet_id_match.group(1)
                        csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
                        
                        with st.spinner("Mengambil data dari Google Sheets..."):
                            response = requests.get(csv_url, timeout=30)
                            response.raise_for_status()
                            
                            csv_content = response.content.decode('utf-8-sig')
                            
                            if ';' in csv_content.split('\n')[0]:
                                df = pd.read_csv(StringIO(csv_content), dtype=str, sep=';')
                            else:
                                df = pd.read_csv(StringIO(csv_content), dtype=str)
                            
                            df.columns = df.columns.str.strip().str.replace('\ufeff', '')
                            data_source = "sheets"
                            st.session_state["sheets_df"] = df
                            st.session_state["last_uploaded_file"] = sheets_url
                            st.success(f"✅ Berhasil mengambil {len(df)} baris data!")
                    else:
                        st.error("❌ URL tidak valid. Pastikan URL dari Google Sheets.")
                except requests.exceptions.RequestException as e:
                    st.error("❌ Gagal mengambil data. Pastikan Google Sheets sudah di-share sebagai 'Anyone with the link'")
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")
            
            if "sheets_df" in st.session_state and st.session_state.get("last_uploaded_file") == sheets_url:
                df = st.session_state["sheets_df"]
                data_source = "sheets"
        
        if df is not None:
            undian_col = None
            for col in df.columns:
                if "undian" in col.lower():
                    undian_col = col
                    break
            
            name_col = None
            for col in df.columns:
                col_lower = col.lower()
                if "nama" in col_lower or "name" in col_lower:
                    name_col = col
                    break
            
            phone_col = None
            phone_candidates_priority = []
            phone_candidates_secondary = []
            
            for col in df.columns:
                col_lower = col.lower()
                non_empty_count = df[col].notna().sum() - (df[col].astype(str).str.strip() == "").sum() - (df[col].astype(str).str.lower() == "nan").sum()
                
                if "hp" in col_lower or "phone" in col_lower or "telp" in col_lower or "telepon" in col_lower:
                    phone_candidates_priority.append((col, non_empty_count))
                elif col_lower == "wa" or col_lower.endswith(" wa") or "nomor wa" in col_lower or "(wa)" in col_lower:
                    phone_candidates_secondary.append((col, non_empty_count))
            
            if phone_candidates_priority:
                phone_candidates_priority.sort(key=lambda x: x[1], reverse=True)
                phone_col = phone_candidates_priority[0][0]
            elif phone_candidates_secondary:
                phone_candidates_secondary.sort(key=lambda x: x[1], reverse=True)
                phone_col = phone_candidates_secondary[0][0]
            
            if undian_col is None:
                st.error("❌ Error: File harus memiliki kolom 'Nomor Undian'")
                st.info("Kolom yang ditemukan: " + ", ".join(df.columns.tolist()))
            else:
                df["Nomor Undian"] = df[undian_col].astype(str).str.strip()
                
                if name_col:
                    df["Nama"] = df[name_col].astype(str).str.strip()
                else:
                    df["Nama"] = ""
                
                if phone_col:
                    df["No HP"] = df[phone_col].astype(str).str.strip()
                else:
                    df["No HP"] = ""
                
                df = df[df["Nomor Undian"].notna() & (df["Nomor Undian"] != "") & (df["Nomor Undian"] != "nan")]
                
                df["Nomor Undian"] = df["Nomor Undian"].apply(lambda x: str(x).zfill(4))
                
                participant_data = df[["Nomor Undian", "Nama", "No HP"]].copy()
                participant_data = participant_data.drop_duplicates(subset=["Nomor Undian"])
                
                participant_data["Eligible"] = participant_data.apply(
                    lambda row: is_eligible_for_prize(row["Nama"], row["No HP"]), 
                    axis=1
                )
                
                st.session_state["participant_data"] = participant_data
                
                all_participants = participant_data["Nomor Undian"].tolist()
                total_participants = len(all_participants)
                
                eligible_data = participant_data[participant_data["Eligible"] == True]
                eligible_participants = eligible_data["Nomor Undian"].tolist()
                total_eligible = len(eligible_participants)
                total_excluded = total_participants - total_eligible
                
                st.session_state["total_participants"] = total_participants
                st.session_state["eligible_participants"] = eligible_participants
                st.session_state["total_eligible"] = total_eligible
                
                if "remaining_pool" not in st.session_state or st.session_state.get("remaining_pool") is None:
                    st.session_state["remaining_pool"] = eligible_data.copy()
                
                prize_tiers = st.session_state.get("prize_tiers", PRIZE_TIERS)
                total_winners = calculate_total_winners(prize_tiers)
                num_categories = len(prize_tiers)
                
                remaining_pool = st.session_state.get("remaining_pool")
                remaining_count = len(remaining_pool) if remaining_pool is not None else total_eligible
                
                st.markdown(f"""
                <div style="background: linear-gradient(135deg, #1a1a5a 0%, #2d2d7d 100%); border-radius: 15px; padding: 1.5rem; margin: 1rem 0;">
                    <div style="display: flex; justify-content: space-around; text-align: center; flex-wrap: wrap;">
                        <div style="padding: 0.5rem 1rem;">
                            <div style="color: #90EE90; font-size: 2rem; font-weight: bold;">{total_participants:,}</div>
                            <div style="color: white; font-size: 0.9rem;">🎟️ Total Data</div>
                        </div>
                        <div style="padding: 0.5rem 1rem;">
                            <div style="color: #FFD700; font-size: 2rem; font-weight: bold;">{remaining_count:,}</div>
                            <div style="color: white; font-size: 0.9rem;">✅ Eligible</div>
                        </div>
                        <div style="padding: 0.5rem 1rem;">
                            <div style="color: #FF6B6B; font-size: 2rem; font-weight: bold;">{total_excluded:,}</div>
                            <div style="color: white; font-size: 0.9rem;">🚫 VIP/F</div>
                        </div>
                    </div>
                </div>
                """, unsafe_allow_html=True)
                
                st.markdown("<br>", unsafe_allow_html=True)
                st.markdown("<p style='text-align:center; color:white; font-size:1.5rem; font-weight:bold;'>🎯 PILIH MODE UNDIAN</p>", unsafe_allow_html=True)
                st.markdown("<br>", unsafe_allow_html=True)
                
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.markdown(f"""
                    <div style="background: rgba(76,175,80,0.2); border-radius: 15px; padding: 1.5rem; text-align: center; border: 2px solid #4CAF50; min-height: 200px;">
                        <p style="color: #4CAF50; font-size: 1.5rem; font-weight: bold; margin: 0;">🎁 Undian E-Voucher</p>
                        <p style="color: white; font-size: 1rem; margin: 0.5rem 0;">{total_winners} hadiah, {num_categories} kategori</p>
                        <p style="color: #aaa; font-size: 0.85rem; margin: 0.5rem 0;">
                        Dikirim lewat WA<br>
                        Ada animasi saat undian<br>
                        Download PPT, XLS, Sisa
                        </p>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    st.markdown("<br>", unsafe_allow_html=True)
                    if st.button(
                        "🎲 LIHAT KATEGORI E-VOUCHER",
                        key="show_evoucher_categories",
                        use_container_width=True,
                        type="primary"
                    ):
                        st.session_state["show_evoucher_preview"] = True
                        st.rerun()
                
                start_lottery = False
                
                if st.session_state.get("show_evoucher_preview"):
                    st.markdown("---")
                    st.markdown("<p style='text-align:center; color:white; font-size:1.8rem; font-weight:bold;'>🎁 KATEGORI HADIAH E-VOUCHER</p>", unsafe_allow_html=True)
                    st.markdown("<br>", unsafe_allow_html=True)
                    
                    prize_cols = st.columns(3)
                    for idx, tier in enumerate(prize_tiers):
                        col_idx = idx % 3
                        with prize_cols[col_idx]:
                            st.markdown(f"""
                            <div style="background: white; border-radius: 20px; padding: 1.5rem; text-align: center; margin-bottom: 1rem; box-shadow: 0 4px 15px rgba(0,0,0,0.2);">
                                <div style="font-size: 2.5rem; margin-bottom: 0.5rem;">{tier['icon']}</div>
                                <p style="color: #333; font-size: 1.1rem; font-weight: bold; margin: 0;">{tier['name']}</p>
                                <p style="color: #666; font-size: 0.9rem; margin: 0.3rem 0;">Peringkat {tier['start']}-{tier['end']}</p>
                                <p style="color: #FF6B6B; font-size: 1.1rem; font-weight: bold; margin: 0;">{tier['count']} Pemenang</p>
                            </div>
                            """, unsafe_allow_html=True)
                    
                    st.markdown("<br>", unsafe_allow_html=True)
                    
                    col_back, col_start = st.columns(2)
                    with col_back:
                        if st.button("⬅️ KEMBALI", key="back_from_preview", use_container_width=True):
                            st.session_state["show_evoucher_preview"] = False
                            st.rerun()
                    with col_start:
                        start_lottery = st.button(
                            "🎲 MULAI UNDIAN SEKARANG",
                            key="start_evoucher_now",
                            use_container_width=True,
                            type="primary"
                        )
                
                with col2:
                    st.markdown(f"""
                    <div style="background: rgba(255,152,0,0.2); border-radius: 15px; padding: 1.5rem; text-align: center; border: 2px solid #FF9800; min-height: 200px;">
                        <p style="color: #FF9800; font-size: 1.5rem; font-weight: bold; margin: 0;">🎲 Shuffle</p>
                        <p style="color: white; font-size: 1rem; margin: 0.5rem 0;">Luckydraw Sesi 1, 2, 3</p>
                        <p style="color: #aaa; font-size: 0.85rem; margin: 0.5rem 0;">
                        Sisa: {remaining_count} peserta<br>
                        Undian per batch<br>
                        Update otomatis
                        </p>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    st.markdown("<br>", unsafe_allow_html=True)
                    shuffle_num = st.number_input("Jumlah pemenang", min_value=1, max_value=remaining_count, value=min(50, remaining_count), key="shuffle_num_main")
                    shuffle_prize = st.text_input("Nama hadiah", placeholder="Contoh: Voucher Rp.100.000", key="shuffle_prize_main")
                    
                    shuffle_disabled = not shuffle_prize
                    if st.button("🎲 MULAI SHUFFLE", key="start_shuffle_main", use_container_width=True, disabled=shuffle_disabled):
                        remaining_numbers = remaining_pool["Nomor Undian"].tolist()
                        batch_winners = []
                        temp_pool = remaining_numbers.copy()
                        for _ in range(min(shuffle_num, len(temp_pool))):
                            idx = secrets.randbelow(len(temp_pool))
                            batch_winners.append(temp_pool.pop(idx))
                        
                        st.session_state["shuffle_mode_active"] = True
                        st.session_state["shuffle_winners"] = batch_winners
                        st.session_state["shuffle_prize_display"] = shuffle_prize
                        st.session_state["main_shuffle_mode"] = True
                        st.rerun()
                
                with col3:
                    st.markdown(f"""
                    <div style="background: rgba(233,30,99,0.2); border-radius: 15px; padding: 1.5rem; text-align: center; border: 2px solid #E91E63; min-height: 200px;">
                        <p style="color: #E91E63; font-size: 1.5rem; font-weight: bold; margin: 0;">🎡 Spinning Wheel</p>
                        <p style="color: white; font-size: 1rem; margin: 0.5rem 0;">Hadiah Utama</p>
                        <p style="color: #aaa; font-size: 0.85rem; margin: 0.5rem 0;">
                        Sisa: {remaining_count} peserta<br>
                        Diputar per hadiah<br>
                        Efek dramatis
                        </p>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    st.markdown("<br>", unsafe_allow_html=True)
                    wheel_num = st.number_input("Jumlah hadiah", min_value=1, max_value=min(50, remaining_count), value=min(10, remaining_count), key="wheel_num_main")
                    wheel_prize = st.text_input("Nama hadiah", placeholder="Contoh: Grand Prize", key="wheel_prize_main")
                    
                    wheel_disabled = not wheel_prize
                    if st.button("🎡 MULAI WHEEL", key="start_wheel_main", use_container_width=True, type="primary", disabled=wheel_disabled):
                        remaining_numbers = remaining_pool["Nomor Undian"].tolist()
                        batch_winners = []
                        temp_pool = remaining_numbers.copy()
                        for _ in range(min(wheel_num, len(temp_pool))):
                            idx = secrets.randbelow(len(temp_pool))
                            batch_winners.append(temp_pool.pop(idx))
                        
                        st.session_state["wheel_mode_active"] = True
                        st.session_state["wheel_winners"] = batch_winners
                        st.session_state["wheel_current_index"] = 0
                        st.session_state["wheel_prize_display"] = wheel_prize
                        st.session_state["main_wheel_mode"] = True
                        st.rerun()
                
                if st.session_state.get("shuffle_mode_active") and st.session_state.get("main_shuffle_mode"):
                    shuffle_winners = st.session_state.get("shuffle_winners", [])
                    prize_display = st.session_state.get("shuffle_prize_display", "Hadiah")
                    
                    st.markdown("---")
                    st.markdown(f"""
                    <div style="text-align:center; color:white; font-size:1.8rem; margin: 1rem 0;">
                        🎲 {prize_display} - {len(shuffle_winners)} Pemenang
                    </div>
                    """, unsafe_allow_html=True)
                    
                    shuffle_html = create_shuffle_reveal_html(shuffle_winners, prize_display)
                    components.html(shuffle_html, height=500)
                    
                    st.markdown("**Detail Pemenang:**")
                    winner_details = []
                    for w in shuffle_winners:
                        row = participant_data[participant_data["Nomor Undian"] == w]
                        if len(row) > 0:
                            nama = row.iloc[0].get("Nama", "")
                            hp = str(row.iloc[0].get("No HP", ""))
                            hp_masked = f"****{hp[-4:]}" if len(hp) >= 4 else ""
                            winner_details.append({"Nomor Undian": w, "Nama": nama, "No HP": hp_masked})
                    if winner_details:
                        st.dataframe(pd.DataFrame(winner_details), use_container_width=True, hide_index=True)
                    
                    col_s1, col_s2 = st.columns(2)
                    with col_s1:
                        if st.button("❌ Batalkan", key="shuffle_cancel_main", use_container_width=True):
                            st.session_state["shuffle_mode_active"] = False
                            st.session_state["main_shuffle_mode"] = False
                            st.rerun()
                    with col_s2:
                        if st.button("✅ Konfirmasi & Update Sisa", key="shuffle_confirm_main", use_container_width=True, type="primary"):
                            new_pool = remaining_pool[~remaining_pool["Nomor Undian"].isin(shuffle_winners)]
                            st.session_state["remaining_pool"] = new_pool
                            
                            all_ws = st.session_state.get("all_wheel_shuffle_winners", [])
                            for w in shuffle_winners:
                                all_ws.append({"Nomor Undian": w, "Hadiah": prize_display})
                            st.session_state["all_wheel_shuffle_winners"] = all_ws
                            
                            st.session_state["shuffle_mode_active"] = False
                            st.session_state["main_shuffle_mode"] = False
                            st.success(f"✅ {len(shuffle_winners)} pemenang dikonfirmasi! Sisa peserta: {len(new_pool)}")
                            st.rerun()
                
                if st.session_state.get("wheel_mode_active") and st.session_state.get("main_wheel_mode"):
                    wheel_winners = st.session_state.get("wheel_winners", [])
                    wheel_idx = st.session_state.get("wheel_current_index", 0)
                    prize_display = st.session_state.get("wheel_prize_display", "Hadiah")
                    
                    st.markdown("---")
                    
                    if wheel_idx < len(wheel_winners):
                        current_winner = wheel_winners[wheel_idx]
                        
                        winner_name = ""
                        winner_phone = ""
                        winner_row = participant_data[participant_data["Nomor Undian"] == current_winner]
                        if len(winner_row) > 0:
                            winner_name = winner_row.iloc[0].get("Nama", "")
                            phone = str(winner_row.iloc[0].get("No HP", ""))
                            if len(phone) >= 4:
                                winner_phone = f"****{phone[-4:]}"
                        
                        st.markdown(f"""
                        <div style="text-align:center; color:white; font-size:1.8rem; margin: 1rem 0;">
                            🎯 {prize_display}
                        </div>
                        <div style="text-align:center; color:#fff9c4; font-size:1.2rem; margin-bottom: 1rem;">
                            Pemenang ke-{wheel_idx + 1} dari {len(wheel_winners)}
                        </div>
                        """, unsafe_allow_html=True)
                        
                        wheel_html = create_spinning_wheel_html(wheel_winners, current_winner, wheel_size=450)
                        components.html(wheel_html, height=650)
                        
                        if winner_name or winner_phone:
                            st.markdown(f"""
                            <div style="text-align:center; background: rgba(255,255,255,0.1); padding: 0.5rem; border-radius: 10px; margin-top: 0.5rem;">
                                <span style="color: white; font-size: 1rem;">👤 {winner_name} | 📱 {winner_phone}</span>
                            </div>
                            """, unsafe_allow_html=True)
                        
                        col_nav1, col_nav2, col_nav3 = st.columns(3)
                        with col_nav1:
                            if wheel_idx > 0:
                                if st.button("⬅️ Sebelumnya", key="wheel_prev_main", use_container_width=True):
                                    st.session_state["wheel_current_index"] = wheel_idx - 1
                                    st.rerun()
                        with col_nav2:
                            if st.button("❌ Batalkan", key="wheel_cancel_main", use_container_width=True):
                                st.session_state["wheel_mode_active"] = False
                                st.session_state["main_wheel_mode"] = False
                                st.rerun()
                        with col_nav3:
                            if wheel_idx < len(wheel_winners) - 1:
                                if st.button("➡️ Selanjutnya", key="wheel_next_main", use_container_width=True):
                                    st.session_state["wheel_current_index"] = wheel_idx + 1
                                    st.rerun()
                            else:
                                if st.button("✅ Konfirmasi Semua", key="wheel_confirm_main", use_container_width=True, type="primary"):
                                    new_pool = remaining_pool[~remaining_pool["Nomor Undian"].isin(wheel_winners)]
                                    st.session_state["remaining_pool"] = new_pool
                                    
                                    all_ws = st.session_state.get("all_wheel_shuffle_winners", [])
                                    for w in wheel_winners:
                                        all_ws.append({"Nomor Undian": w, "Hadiah": prize_display})
                                    st.session_state["all_wheel_shuffle_winners"] = all_ws
                                    
                                    st.session_state["wheel_mode_active"] = False
                                    st.session_state["main_wheel_mode"] = False
                                    st.success(f"✅ {len(wheel_winners)} pemenang dikonfirmasi! Sisa peserta: {len(new_pool)}")
                                    st.rerun()
                
                all_ws_winners = st.session_state.get("all_wheel_shuffle_winners", [])
                if len(all_ws_winners) > 0 and not st.session_state.get("wheel_mode_active") and not st.session_state.get("shuffle_mode_active"):
                    st.markdown("---")
                    st.markdown("<p style='text-align:center; color:white; font-size:1.3rem; font-weight:600;'>📋 Rekap Pemenang Shuffle/Wheel</p>", unsafe_allow_html=True)
                    
                    ws_df = pd.DataFrame(all_ws_winners)
                    ws_df = ws_df.merge(
                        participant_data[["Nomor Undian", "Nama", "No HP"]], 
                        on="Nomor Undian", 
                        how="left"
                    )
                    
                    by_prize = ws_df.groupby("Hadiah")
                    for prize_name_g, group in by_prize:
                        st.markdown(f"**{prize_name_g}** ({len(group)} pemenang)")
                    
                    col_exp1, col_exp2 = st.columns(2)
                    with col_exp1:
                        excel_buffer = BytesIO()
                        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                            ws_df.to_excel(writer, index=False, sheet_name='Pemenang')
                        st.download_button(
                            label="📊 Download Excel Pemenang",
                            data=excel_buffer.getvalue(),
                            file_name="pemenang_shuffle_wheel.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    with col_exp2:
                        remaining_csv = remaining_pool[["Nomor Undian", "Nama", "No HP"]].to_csv(index=False)
                        st.download_button(
                            label=f"📥 Download Sisa ({len(remaining_pool)})",
                            data=remaining_csv,
                            file_name="sisa_peserta.csv",
                            mime="text/csv",
                            use_container_width=True
                        )
                
                if start_lottery:
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    for i in range(100):
                        progress_bar.progress(i + 1)
                        if i < 30:
                            status_text.markdown(f"<p style='text-align:center; font-size:1.5rem; color:white;'>🔄 Mengumpulkan data peserta... {i+1}%</p>", unsafe_allow_html=True)
                        elif i < 70:
                            status_text.markdown(f"<p style='text-align:center; font-size:1.5rem; color:white;'>🎲 Mengacak peserta secara acak... {i+1}%</p>", unsafe_allow_html=True)
                        else:
                            status_text.markdown(f"<p style='text-align:center; font-size:1.5rem; color:white;'>🏆 Menentukan pemenang... {i+1}%</p>", unsafe_allow_html=True)
                        time.sleep(0.03)
                    
                    eligible_participants = st.session_state.get("eligible_participants", eligible_participants)
                    shuffled_participants = secure_shuffle(eligible_participants)
                    
                    prize_tiers = st.session_state.get("prize_tiers", PRIZE_TIERS)
                    total_winners_needed = calculate_total_winners(prize_tiers)
                    num_winners = min(total_winners_needed, len(shuffled_participants))
                    winners = shuffled_participants[:num_winners]
                    
                    participant_data = st.session_state.get("participant_data")
                    if participant_data is not None:
                        name_lookup = dict(zip(participant_data["Nomor Undian"], participant_data["Nama"]))
                        phone_lookup = dict(zip(participant_data["Nomor Undian"], participant_data["No HP"]))
                    else:
                        name_lookup = {}
                        phone_lookup = {}
                    
                    prize_tiers = st.session_state.get("prize_tiers", PRIZE_TIERS)
                    
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
                    
                    st.session_state["results_df"] = results_df
                    st.session_state["lottery_done"] = True
                    
                    progress_bar.empty()
                    status_text.empty()
                    
                    st.balloons()
                    st.rerun()
        
        if df is None:
            st.markdown("### 🎁 Daftar Hadiah")
            prize_tiers = st.session_state.get("prize_tiers", PRIZE_TIERS)
            cols = st.columns(3)
            for idx, tier in enumerate(prize_tiers):
                col_idx = idx % 3
                with cols[col_idx]:
                    st.markdown(f"""
                    <div class="prize-card-clickable">
                        <div style="font-size: 2.5rem;">{tier["icon"]}</div>
                        <div style="font-weight: 700; color: #333; font-size: 1.1rem; margin-top: 0.5rem;">{tier["name"]}</div>
                        <div style="font-size: 0.9rem; color: #666; margin-top: 0.3rem;">Peringkat {tier["start"]}-{tier["end"]}</div>
                        <div style="font-size: 1rem; color: #f5576c; font-weight: 600; margin-top: 0.5rem;">{tier["count"]} Pemenang</div>
                    </div>
                    """, unsafe_allow_html=True)
