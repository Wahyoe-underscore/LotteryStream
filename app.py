import streamlit as st
import pandas as pd
import secrets

st.set_page_config(
    page_title="Sistem Undian Move & Groove",
    page_icon="🎉",
    layout="wide"
)

PRIZE_TIERS = [
    {"name": "Bensin Rp.100.000,-", "start": 1, "end": 100},
    {"name": "Top100 Rp.100.000,-", "start": 101, "end": 200},
    {"name": "SNL Rp.100.000,-", "start": 201, "end": 300},
    {"name": "Bensin Rp.150.000,-", "start": 301, "end": 400},
    {"name": "Top100 Rp.150.000,-", "start": 401, "end": 500},
    {"name": "SNL Rp.150.000,-", "start": 501, "end": 600},
    {"name": "Bensin Rp.200.000,-", "start": 601, "end": 700},
    {"name": "Top100 Rp.200.000,-", "start": 701, "end": 800},
    {"name": "SNL Rp.200.000,-", "start": 801, "end": 900},
]

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

st.title("🎉 Sistem Undian Move & Groove")
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
        df = pd.read_csv(uploaded_file)
        
        if "Nomor Undian" not in df.columns:
            st.error("❌ Error: File CSV harus memiliki kolom 'Nomor Undian'")
            st.info("Kolom yang ditemukan: " + ", ".join(df.columns.tolist()))
        else:
            participants = df["Nomor Undian"].dropna().astype(str).tolist()
            total_participants = len(participants)
            
            st.success(f"✅ File berhasil dimuat! Total peserta: **{total_participants:,}**")
            
            if total_participants < 900:
                st.warning(f"⚠️ Peringatan: Jumlah peserta ({total_participants}) kurang dari 900. Semua peserta akan menjadi pemenang.")
            
            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                start_lottery = st.button(
                    "🎲 Mulai Undian",
                    use_container_width=True,
                    type="primary"
                )
            
            if start_lottery:
                with st.spinner("🔄 Mengacak peserta..."):
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
            
            if st.session_state.get("lottery_done", False):
                results_df = st.session_state["results_df"]
                
                st.markdown("---")
                st.subheader("🏆 Daftar Pemenang")
                
                st.markdown("### 📊 Ringkasan Hadiah")
                prize_summary = results_df.groupby("Hadiah").size().reset_index(name="Jumlah Pemenang")
                prize_order = [tier["name"] for tier in PRIZE_TIERS]
                prize_summary["Urutan"] = prize_summary["Hadiah"].apply(
                    lambda x: prize_order.index(x) if x in prize_order else 999
                )
                prize_summary = prize_summary.sort_values("Urutan").drop("Urutan", axis=1)
                st.dataframe(prize_summary, use_container_width=True, hide_index=True)
                
                st.markdown("### 📋 Daftar Lengkap Pemenang")
                st.dataframe(results_df, use_container_width=True, hide_index=True, height=500)
                
                csv_data = results_df.to_csv(index=False).encode("utf-8")
                
                st.markdown("---")
                col1, col2, col3 = st.columns([1, 2, 1])
                with col2:
                    st.download_button(
                        label="📥 Download Hasil Undian",
                        data=csv_data,
                        file_name="hasil_undian_move_groove.csv",
                        mime="text/csv",
                        use_container_width=True,
                        type="primary"
                    )
                
    except Exception as e:
        st.error(f"❌ Error membaca file: {str(e)}")

else:
    st.info("👆 Silakan upload file CSV untuk memulai undian")
    
    st.markdown("### 📝 Format File CSV")
    st.markdown("""
    File CSV harus memiliki kolom dengan nama **'Nomor Undian'**
    
    Contoh format:
    """)
    
    example_df = pd.DataFrame({
        "Nomor Undian": ["001", "002", "003", "004", "005"]
    })
    st.dataframe(example_df, use_container_width=False, hide_index=True)
