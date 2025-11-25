import streamlit as st
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
    
    .winner-card {
        background: linear-gradient(145deg, #fff9c4 0%, #ffecb3 100%);
        border-radius: 15px;
        padding: 1.5rem;
        margin: 0.5rem 0;
        box-shadow: 0 5px 20px rgba(0,0,0,0.15);
        border-left: 5px solid #ffd700;
    }
    
    .prize-badge {
        background: linear-gradient(135deg, #f5576c 0%, #f093fb 100%);
        color: white;
        padding: 0.5rem 1.5rem;
        border-radius: 25px;
        font-weight: 700;
        display: inline-block;
        font-size: 1.1rem;
        box-shadow: 0 4px 15px rgba(245, 87, 108, 0.4);
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
    
    .upload-section {
        background: rgba(255,255,255,0.95);
        border-radius: 20px;
        padding: 2rem;
        box-shadow: 0 10px 40px rgba(0,0,0,0.2);
        margin: 1rem 0;
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
    
    .dataframe {
        font-size: 1.1rem !important;
    }
    
    div[data-testid="stDataFrame"] {
        background: white;
        border-radius: 15px;
        padding: 1rem;
        box-shadow: 0 5px 20px rgba(0,0,0,0.1);
    }
    
    .confetti {
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        pointer-events: none;
        z-index: 9999;
    }
    
    .info-box {
        background: rgba(255,255,255,0.9);
        border-radius: 15px;
        padding: 1.5rem;
        margin: 1rem 0;
        border-left: 5px solid #667eea;
    }
    
    .prize-summary-card {
        background: white;
        border-radius: 15px;
        padding: 1rem 1.5rem;
        margin: 0.5rem;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        text-align: center;
        transition: transform 0.3s ease;
    }
    
    .prize-summary-card:hover {
        transform: translateY(-5px);
    }
</style>
""", unsafe_allow_html=True)

PRIZE_TIERS = [
    {"name": "Bensin Rp.100.000,-", "start": 1, "end": 100, "icon": "⛽"},
    {"name": "Top100 Rp.100.000,-", "start": 101, "end": 200, "icon": "💳"},
    {"name": "SNL Rp.100.000,-", "start": 201, "end": 300, "icon": "🎁"},
    {"name": "Bensin Rp.150.000,-", "start": 301, "end": 400, "icon": "⛽"},
    {"name": "Top100 Rp.150.000,-", "start": 401, "end": 500, "icon": "💳"},
    {"name": "SNL Rp.150.000,-", "start": 501, "end": 600, "icon": "🎁"},
    {"name": "Bensin Rp.200.000,-", "start": 601, "end": 700, "icon": "⛽"},
    {"name": "Top100 Rp.200.000,-", "start": 701, "end": 800, "icon": "💳"},
    {"name": "SNL Rp.200.000,-", "start": 801, "end": 900, "icon": "🎁"},
]

def get_prize(rank):
    for tier in PRIZE_TIERS:
        if tier["start"] <= rank <= tier["end"]:
            return tier["name"]
    return "Tidak Ada Hadiah"

def get_prize_icon(rank):
    for tier in PRIZE_TIERS:
        if tier["start"] <= rank <= tier["end"]:
            return tier["icon"]
    return "🎁"

def secure_shuffle(data_list):
    shuffled = data_list.copy()
    n = len(shuffled)
    for i in range(n - 1, 0, -1):
        j = secrets.randbelow(i + 1)
        shuffled[i], shuffled[j] = shuffled[j], shuffled[i]
    return shuffled

st.image("attached_assets/Small Banner-01_1764081768006.png", use_container_width=True)

st.markdown('<p class="main-title">🎉 Sistem Undian Move & Groove 🎉</p>', unsafe_allow_html=True)
st.markdown('<p class="subtitle">Awareness of Moving The Body for Bone Health</p>', unsafe_allow_html=True)

st.markdown("---")

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
                    <div class="stats-number">900</div>
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
            
            if total_participants < 900:
                st.warning(f"⚠️ Peringatan: Jumlah peserta ({total_participants}) kurang dari 900. Semua peserta akan menjadi pemenang.")
            
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
                
                num_winners = min(900, len(shuffled_participants))
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
            
            if st.session_state.get("lottery_done", False):
                results_df = st.session_state["results_df"]
                
                st.markdown('<div class="section-header">🏆 DAFTAR PEMENANG 🏆</div>', unsafe_allow_html=True)
                
                st.markdown("### 📊 Ringkasan Hadiah")
                
                cols = st.columns(3)
                for idx, tier in enumerate(PRIZE_TIERS):
                    col_idx = idx % 3
                    with cols[col_idx]:
                        count = len(results_df[results_df["Hadiah"] == tier["name"]])
                        st.markdown(f"""
                        <div class="prize-summary-card">
                            <div style="font-size: 2rem;">{tier["icon"]}</div>
                            <div style="font-weight: 700; color: #333; font-size: 1rem;">{tier["name"]}</div>
                            <div style="font-size: 1.5rem; font-weight: 800; color: #f5576c;">{count} Pemenang</div>
                            <div style="font-size: 0.9rem; color: #666;">Peringkat {tier["start"]}-{tier["end"]}</div>
                        </div>
                        """, unsafe_allow_html=True)
                
                st.markdown("<br>", unsafe_allow_html=True)
                st.markdown("### 📋 Daftar Lengkap Pemenang")
                
                styled_df = results_df.copy()
                st.dataframe(
                    styled_df,
                    use_container_width=True,
                    hide_index=True,
                    height=600,
                    column_config={
                        "Peringkat": st.column_config.NumberColumn(
                            "🏅 Peringkat",
                            format="%d"
                        ),
                        "Nomor Undian": st.column_config.TextColumn(
                            "🎫 Nomor Undian"
                        ),
                        "Hadiah": st.column_config.TextColumn(
                            "🎁 Hadiah"
                        )
                    }
                )
                
                excel_buffer = BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    results_df.to_excel(writer, index=False, sheet_name='Hasil Undian')
                excel_data = excel_buffer.getvalue()
                
                st.markdown("<br>", unsafe_allow_html=True)
                col1, col2, col3 = st.columns([1, 2, 1])
                with col2:
                    st.download_button(
                        label="📥 DOWNLOAD HASIL UNDIAN (EXCEL)",
                        data=excel_data,
                        file_name="hasil_undian_move_groove.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        type="primary"
                    )
                
    except Exception as e:
        st.error(f"❌ Error membaca file: {str(e)}")

else:
    st.markdown("""
    <div class="info-box">
        <h3 style="color: #667eea; margin-bottom: 1rem;">📝 Cara Penggunaan</h3>
        <ol style="font-size: 1.1rem; line-height: 2;">
            <li>Upload file CSV yang berisi kolom <strong>'Nomor Undian'</strong></li>
            <li>Klik tombol <strong>'MULAI UNDIAN'</strong></li>
            <li>Sistem akan mengacak dan memilih <strong>900 pemenang</strong></li>
            <li>Download hasil undian dalam format <strong>Excel</strong></li>
        </ol>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("### 📄 Contoh Format CSV")
    example_df = pd.DataFrame({
        "Nomor Undian": ["0001", "0002", "0003", "0004", "0005"]
    })
    st.dataframe(example_df, use_container_width=False, hide_index=True)
    
    st.markdown("### 🎁 Daftar Hadiah")
    cols = st.columns(3)
    for idx, tier in enumerate(PRIZE_TIERS):
        col_idx = idx % 3
        with cols[col_idx]:
            st.markdown(f"""
            <div class="prize-summary-card">
                <div style="font-size: 2rem;">{tier["icon"]}</div>
                <div style="font-weight: 700; color: #333;">{tier["name"]}</div>
                <div style="font-size: 0.9rem; color: #666;">Peringkat {tier["start"]}-{tier["end"]}</div>
            </div>
            """, unsafe_allow_html=True)
