"""
Generate Move & Groove Lottery Flow Presentation
Creates a comprehensive PPT documenting the entire lottery process
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from io import BytesIO

def create_flow_presentation():
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    
    # Color scheme
    PINK = RGBColor(233, 30, 99)
    PURPLE = RGBColor(156, 39, 176)
    GREEN = RGBColor(76, 175, 80)
    ORANGE = RGBColor(255, 152, 0)
    DARK = RGBColor(33, 33, 33)
    WHITE = RGBColor(255, 255, 255)
    
    def add_title_slide(title, subtitle=""):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # Background
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = PINK
        bg.line.fill.background()
        
        # Title
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.5), Inches(12.33), Inches(1.5))
        tf = title_box.text_frame
        tf.paragraphs[0].text = title
        tf.paragraphs[0].font.size = Pt(54)
        tf.paragraphs[0].font.bold = True
        tf.paragraphs[0].font.color.rgb = WHITE
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        
        if subtitle:
            p = tf.add_paragraph()
            p.text = subtitle
            p.font.size = Pt(28)
            p.font.color.rgb = WHITE
            p.alignment = PP_ALIGN.CENTER
        
        return slide
    
    def add_content_slide(title, content_items, highlight_color=PINK):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # Header bar
        header = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(1.2))
        header.fill.solid()
        header.fill.fore_color.rgb = highlight_color
        header.line.fill.background()
        
        # Title
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.33), Inches(0.8))
        tf = title_box.text_frame
        tf.paragraphs[0].text = title
        tf.paragraphs[0].font.size = Pt(36)
        tf.paragraphs[0].font.bold = True
        tf.paragraphs[0].font.color.rgb = WHITE
        
        # Content
        content_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.6), Inches(11.73), Inches(5.5))
        tf = content_box.text_frame
        tf.word_wrap = True
        
        for i, item in enumerate(content_items):
            if i == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()
            p.text = item
            p.font.size = Pt(24)
            p.font.color.rgb = DARK
            p.space_before = Pt(12)
            p.space_after = Pt(8)
        
        return slide
    
    def add_step_slide(step_num, title, description, details, color=PINK):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # Step number circle
        circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.5), Inches(0.4), Inches(1), Inches(1))
        circle.fill.solid()
        circle.fill.fore_color.rgb = color
        circle.line.fill.background()
        
        num_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.6), Inches(1), Inches(0.6))
        tf = num_box.text_frame
        tf.paragraphs[0].text = str(step_num)
        tf.paragraphs[0].font.size = Pt(36)
        tf.paragraphs[0].font.bold = True
        tf.paragraphs[0].font.color.rgb = WHITE
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        
        # Title
        title_box = slide.shapes.add_textbox(Inches(1.8), Inches(0.5), Inches(10), Inches(0.8))
        tf = title_box.text_frame
        tf.paragraphs[0].text = title
        tf.paragraphs[0].font.size = Pt(40)
        tf.paragraphs[0].font.bold = True
        tf.paragraphs[0].font.color.rgb = color
        
        # Description
        desc_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.6), Inches(11.73), Inches(1))
        tf = desc_box.text_frame
        tf.paragraphs[0].text = description
        tf.paragraphs[0].font.size = Pt(24)
        tf.paragraphs[0].font.color.rgb = DARK
        
        # Details list
        details_box = slide.shapes.add_textbox(Inches(1), Inches(2.8), Inches(11.33), Inches(4))
        tf = details_box.text_frame
        tf.word_wrap = True
        
        for i, detail in enumerate(details):
            if i == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()
            p.text = "• " + detail
            p.font.size = Pt(22)
            p.font.color.rgb = DARK
            p.space_before = Pt(10)
        
        return slide
    
    # ========== SLIDES ==========
    
    # 1. Cover slide
    add_title_slide(
        "SISTEM UNDIAN MOVE & GROOVE",
        "7 Desember 2024 | Panduan Alur Pengundian"
    )
    
    # 2. Overview
    add_content_slide("OVERVIEW SISTEM UNDIAN", [
        "Sistem undian digital dengan 3 tahap pengundian:",
        "",
        "1. E-VOUCHER: 700 pemenang (4 kategori x 175 hadiah)",
        "2. SHUFFLE: 90 pemenang (3 sesi x 30 hadiah)",
        "3. WHEEL: 10 Grand Prize (terurut dari termurah ke termahal)",
        "",
        "Total: 800 pemenang dari seluruh peserta eligible",
        "",
        "Fitur: Auto-save, Backup Google Drive, Validasi Duplikat"
    ], PURPLE)
    
    # 3. Data Setup
    add_step_slide(1, "UPLOAD DATA PESERTA", 
        "Langkah pertama: Upload data peserta yang akan diundi",
        [
            "Upload CSV file atau masukkan URL Google Sheets",
            "Format: Nomor Undian, Nama, No HP",
            "VIP/F otomatis terfilter (tidak ikut undian)",
            "Sistem menampilkan jumlah peserta eligible",
            "Data tersimpan otomatis untuk sesi berikutnya"
        ], PINK)
    
    # 4. Home Screen
    add_step_slide(2, "HALAMAN UTAMA",
        "Dashboard utama menampilkan 3 pilihan mode undian",
        [
            "Panel E-VOUCHER: Selalu aktif, dimulai pertama",
            "Panel SHUFFLE: Aktif setelah E-Voucher selesai",
            "Panel WHEEL: Aktif setelah E-Voucher selesai",
            "Menampilkan status progress setiap tahap",
            "Tombol hasil untuk melihat pemenang sebelumnya"
        ], PINK)
    
    # 5. E-Voucher Stage
    add_step_slide(3, "TAHAP 1: E-VOUCHER",
        "Undian 700 hadiah voucher dalam 4 kategori",
        [
            "Tokopedia Rp.100.000,- : 175 pemenang",
            "Indomaret Rp.100.000,- : 175 pemenang",
            "Bensin Rp.100.000,- : 175 pemenang",
            "SNL Rp.100.000,- : 175 pemenang",
            "",
            "Animasi cascade menampilkan semua pemenang",
            "Download Excel & PPT tersedia setelah selesai"
        ], GREEN)
    
    # 6. E-Voucher Animation
    add_content_slide("E-VOUCHER: ANIMASI PENGUNDIAN", [
        "Proses pengundian E-Voucher:",
        "",
        "1. Tekan tombol 'MULAI UNDIAN E-VOUCHER'",
        "2. Progress bar menunjukkan proses shuffle",
        "3. Animasi cascade menampilkan 700 pemenang",
        "4. Pemenang ditampilkan per kategori",
        "5. Download Excel/PPT untuk dokumentasi",
        "6. Tekan 'SISA NOMOR' untuk lanjut ke tahap berikut",
        "",
        "Peserta yang sudah menang TIDAK bisa menang lagi"
    ], GREEN)
    
    # 7. Shuffle Stage
    add_step_slide(4, "TAHAP 2: SHUFFLE",
        "Undian 90 hadiah dalam 3 sesi",
        [
            "Sesi 1: 30 pemenang",
            "Sesi 2: 30 pemenang", 
            "Sesi 3: 30 pemenang",
            "",
            "Setiap sesi memerlukan input nama hadiah",
            "Animasi slot-machine cascade untuk transparansi",
            "Download per sesi atau gabungan tersedia"
        ], ORANGE)
    
    # 8. Shuffle Animation
    add_content_slide("SHUFFLE: ANIMASI SLOT MACHINE", [
        "Proses pengundian setiap sesi Shuffle:",
        "",
        "1. Input nama hadiah untuk sesi tersebut",
        "2. Tekan tombol 'PUTAR UNDIAN'",
        "3. Animasi menampilkan SEMUA nomor peserta sisa",
        "4. 30 pemenang muncul satu per satu secara cascade",
        "5. Kartu pemenang menampilkan: Nomor, Nama, No HP",
        "6. Download hasil per sesi",
        "7. Lanjut ke sesi berikutnya sampai 3 sesi selesai"
    ], ORANGE)
    
    # 9. Wheel Stage
    add_step_slide(5, "TAHAP 3: GRAND PRIZE (WHEEL)",
        "Undian 10 hadiah utama - dari termurah ke termahal",
        [
            "Hadiah #1-10 dikonfigurasi sebelum mulai",
            "Urutan: Termurah dulu, LED TV terakhir",
            "Setiap spin menampilkan SELURUH nomor peserta sisa",
            "Counter menunjukkan progress: '1/2707', '2/2707', dst",
            "Animasi melambat dan berhenti di pemenang",
            "Kartu pemenang langsung ditampilkan dengan detail"
        ], PURPLE)
    
    # 10. Wheel Animation Detail
    add_content_slide("WHEEL: TRANSPARANSI PENUH", [
        "Fitur transparansi pada Wheel:",
        "",
        "1. Badge 'Mengundi dari X peserta' menunjukkan pool",
        "2. Semua nomor ditampilkan berurutan (bukan acak)",
        "3. Counter real-time: '1523/2707' dsb",
        "4. Progress bar mengikuti nomor yang sudah lewat",
        "5. Kecepatan: Cepat di awal, melambat di akhir",
        "6. Pemenang ditampilkan dengan warna emas",
        "7. Detail lengkap: Nomor, Nama, No HP, Hadiah"
    ], PURPLE)
    
    # 11. Validation & Download
    add_step_slide(6, "VALIDASI & DOWNLOAD",
        "Fitur keamanan dan dokumentasi hasil",
        [
            "VALIDASI: Cek duplikat di semua tahap",
            "Memastikan tidak ada nomor yang menang 2x",
            "",
            "DOWNLOAD EXCEL: Semua pemenang dengan detail lengkap",
            "DOWNLOAD PPT: Presentasi siap pakai untuk pengumuman",
            "",
            "Auto-save: Hasil tersimpan otomatis",
            "Google Drive backup: Sinkronisasi ke cloud"
        ], PINK)
    
    # 12. Summary
    add_content_slide("RINGKASAN ALUR", [
        "1. Upload data peserta (CSV/Google Sheets)",
        "2. E-Voucher: 700 pemenang dalam 4 kategori",
        "3. Shuffle: 90 pemenang dalam 3 sesi",
        "4. Wheel: 10 Grand Prize (termurah → termahal)",
        "",
        "Total: 800 pemenang",
        "",
        "Setiap tahap memiliki validasi dan backup otomatis",
        "Semua animasi menampilkan pool peserta untuk transparansi"
    ], PINK)
    
    # 13. Closing
    add_title_slide(
        "MOVE & GROOVE 2024",
        "Selamat Mengundi!"
    )
    
    return prs

if __name__ == "__main__":
    prs = create_flow_presentation()
    prs.save("MoveGroove_Flow_Presentation.pptx")
    print("Presentation saved: MoveGroove_Flow_Presentation.pptx")
