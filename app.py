import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import secrets
from io import BytesIO
import time

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

def secure_shuffle(data_list):
    shuffled = data_list.copy()
    n = len(shuffled)
    for i in range(n - 1, 0, -1):
        j = secrets.randbelow(i + 1)
        shuffled[i], shuffled[j] = shuffled[j], shuffled[i]
    return shuffled

if "selected_prize" not in st.session_state:
    st.session_state["selected_prize"] = None

if st.session_state.get("selected_prize") is not None and st.session_state.get("lottery_done", False):
    selected_tier = st.session_state["selected_prize"]
    results_df = st.session_state["results_df"]
    
    tier_winners = results_df[results_df["Hadiah"] == selected_tier["name"]]
    
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
    grid_height = ((num_winners + num_cols - 1) // num_cols) * 85 + 40
    
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
                font-size: 1.4rem;
                font-weight: 800;
                color: #333;
                margin-top: 0.2rem;
            }}
        </style>
    </head>
    <body>
        <div class="winner-grid">
    '''
    for _, row in tier_winners.iterrows():
        winners_html += f'''
            <div class="winner-cell">
                <div class="winner-rank">#{row["Peringkat"]}</div>
                <div class="winner-number">{row["Nomor Undian"]}</div>
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
        with col3:
            st.markdown(f"""
            <div class="stats-card">
                <div class="stats-number">9</div>
                <div class="stats-label">🎁 Kategori Hadiah</div>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown('<div class="section-header">🏆 PILIH KATEGORI HADIAH 🏆</div>', unsafe_allow_html=True)
        st.markdown("<p style='text-align:center; color:white; font-size:1.2rem;'>Klik pada kategori hadiah untuk melihat pemenang dalam satu layar</p>", unsafe_allow_html=True)
        
        cols = st.columns(3)
        for idx, tier in enumerate(PRIZE_TIERS):
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
        
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                results_df.to_excel(writer, index=False, sheet_name='Hasil Undian')
            excel_data = excel_buffer.getvalue()
            
            st.download_button(
                label="📥 DOWNLOAD SEMUA HASIL UNDIAN (EXCEL)",
                data=excel_data,
                file_name="hasil_undian_move_groove.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary"
            )
        
        st.markdown("<br>", unsafe_allow_html=True)
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            if st.button("🔄 UNDIAN BARU", use_container_width=True):
                st.session_state["lottery_done"] = False
                st.session_state["results_df"] = None
                st.session_state["selected_prize"] = None
                st.rerun()
    
    else:
        uploaded_file = st.file_uploader(
            "📁 Upload File CSV (harus memiliki kolom 'Nomor Undian')",
            type=["csv"],
            help="File CSV harus berisi kolom dengan nama 'Nomor Undian'"
        )
        
        if uploaded_file is not None:
            current_file_name = uploaded_file.name
            if st.session_state.get("last_uploaded_file") != current_file_name:
                st.session_state["last_uploaded_file"] = current_file_name
                st.session_state["lottery_done"] = False
                st.session_state["results_df"] = None
            
            try:
                df = pd.read_csv(uploaded_file, dtype=str)
                
                if "Nomor Undian" not in df.columns:
                    st.error("❌ Error: File CSV harus memiliki kolom 'Nomor Undian'")
                    st.info("Kolom yang ditemukan: " + ", ".join(df.columns.tolist()))
                else:
                    participants = df["Nomor Undian"].dropna().tolist()
                    participants = [str(p).zfill(4) for p in participants]
                    total_participants = len(participants)
                    
                    st.session_state["total_participants"] = total_participants
                    
                    col1, col2, col3 = st.columns(3)
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
                            <div class="stats-number">{TOTAL_WINNERS}</div>
                            <div class="stats-label">🏆 Total Pemenang</div>
                        </div>
                        """, unsafe_allow_html=True)
                    with col3:
                        st.markdown(f"""
                        <div class="stats-card">
                            <div class="stats-number">9</div>
                            <div class="stats-label">🎁 Kategori Hadiah</div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    if total_participants < TOTAL_WINNERS:
                        st.warning(f"⚠️ Peringatan: Jumlah peserta ({total_participants}) kurang dari {TOTAL_WINNERS}. Semua peserta akan menjadi pemenang.")
                    
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
                        
                        shuffled_participants = secure_shuffle(participants)
                        
                        num_winners = min(TOTAL_WINNERS, len(shuffled_participants))
                        winners = shuffled_participants[:num_winners]
                        
                        results = []
                        for i, winner in enumerate(winners, 1):
                            results.append({
                                "Peringkat": i,
                                "Nomor Undian": winner,
                                "Hadiah": get_prize(i)
                            })
                        
                        results_df = pd.DataFrame(results)
                        
                        st.session_state["results_df"] = results_df
                        st.session_state["lottery_done"] = True
                        
                        progress_bar.empty()
                        status_text.empty()
                        
                        st.balloons()
                        st.rerun()
                    
            except Exception as e:
                st.error(f"❌ Error membaca file: {str(e)}")
        
        else:
            st.markdown("### 🎁 Daftar Hadiah")
            cols = st.columns(3)
            for idx, tier in enumerate(PRIZE_TIERS):
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
