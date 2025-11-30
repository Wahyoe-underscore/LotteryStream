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

if st.session_state.get("selected_prize") is not None and st.session_state.get("lottery_done", False):
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
    
    if st.session_state.get("lottery_done", False) and st.session_state.get("results_df") is not None:
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
            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                st.success("✅ Semua hasil undian sudah di-download!")
                if st.button("🔄 UNDIAN BARU", use_container_width=True):
                    st.session_state["lottery_done"] = False
                    st.session_state["results_df"] = None
                    st.session_state["selected_prize"] = None
                    st.session_state["excel_downloaded"] = False
                    st.session_state["pptx_downloaded"] = False
                    st.rerun()
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
            for col in df.columns:
                col_lower = col.lower()
                if "nomor wa" in col_lower or col_lower == "wa":
                    phone_col = col
                    break
            
            if phone_col is None:
                for col in df.columns:
                    col_lower = col.lower()
                    if "hp" in col_lower or "phone" in col_lower or "telp" in col_lower:
                        phone_col = col
                        break
            
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
                
                prize_tiers = st.session_state.get("prize_tiers", PRIZE_TIERS)
                total_winners = calculate_total_winners(prize_tiers)
                num_categories = len(prize_tiers)
                
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.markdown(f"""
                    <div class="stats-card">
                        <div class="stats-number">{total_participants:,}</div>
                        <div class="stats-label">👥 Total Peserta</div>
                    </div>
                    """, unsafe_allow_html=True)
                with col2:
                    st.markdown(f"""
                    <div class="stats-card">
                        <div class="stats-number" style="color: #27ae60;">{total_eligible:,}</div>
                        <div class="stats-label">✅ Berhak Hadiah</div>
                    </div>
                    """, unsafe_allow_html=True)
                with col3:
                    st.markdown(f"""
                    <div class="stats-card">
                        <div class="stats-number">{total_winners}</div>
                        <div class="stats-label">🏆 Total Pemenang</div>
                    </div>
                    """, unsafe_allow_html=True)
                with col4:
                    st.markdown(f"""
                    <div class="stats-card">
                        <div class="stats-number" style="color: #e74c3c;">{total_excluded}</div>
                        <div class="stats-label">🚫 VIP/F (Dikecualikan)</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                if total_excluded > 0:
                    st.info(f"ℹ️ {total_excluded} peserta ditandai VIP/F dan tidak akan ikut undian hadiah.")
                
                if total_eligible < total_winners:
                    st.warning(f"⚠️ Peringatan: Jumlah peserta eligible ({total_eligible}) kurang dari {total_winners}. Semua peserta eligible akan menjadi pemenang.")
                
                st.markdown("<br>", unsafe_allow_html=True)
                
                col1, col2, col3 = st.columns([1, 2, 1])
                with col2:
                    start_lottery = st.button(
                        "🎲 MULAI UNDIAN 🎲",
                        use_container_width=True,
                        type="primary"
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
