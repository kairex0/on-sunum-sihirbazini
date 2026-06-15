import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import io

# -------------------------------------------------------------
# GELİŞMİŞ POWERPOINT MOTORU
# -------------------------------------------------------------
def gelismis_sunum_uret(ilce, mahalle, m2, imar, arsa_fiyati, danisman, konsept, yuklenen_resim):
    prs = Presentation()
    
    # Kurumsal Renkler (ON Türkiye)
    LACIVERT = RGBColor(11, 29, 58)
    TURUNCU = RGBColor(241, 90, 36)
    BEYAZ = RGBColor(255, 255, 255)
    KOYU_GRI = RGBColor(60, 60, 60)
    
    # SLAYT 1: KAPAK
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid(); bg.fill.fore_color.rgb = LACIVERT; bg.line.fill.background()
    
    txBox = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(3))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = f"İZMİR {ilce.upper()} - {mahalle.upper()}"
    p.font.size = Pt(40); p.font.bold = True; p.font.color.rgb = TURUNCU
    
    p2 = tf.add_paragraph()
    p2.text = f"{konsept.upper()} PROJE GELİŞTİRME & YATIRIM ANALİZİ"
    p2.font.size = Pt(18); p2.font.color.rgb = BEYAZ; p2.space_before = Pt(10)
    
    p3 = tf.add_paragraph()
    p3.text = f"\nHazırlayan: {danisman} | ON Türkiye"
    p3.font.size = Pt(14); p3.font.color.rgb = TURUNCU

    # SLAYT 2: PORTFÖY ÖZELLİKLERİ VE FOTOĞRAF
    slide2 = prs.slides.add_slide(prs.slide_layouts[6])
    
    t_box = slide2.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(1))
    t_box.text_frame.paragraphs[0].text = "PORTFÖY ÖZELLİKLERİ VE DETAYLAR"
    t_box.text_frame.paragraphs[0].font.size = Pt(26); t_box.text_frame.paragraphs[0].font.bold = True; t_box.text_frame.paragraphs[0].font.color.rgb = LACIVERT
    
    # Sol tarafa bilgileri yazalım
    tf2 = slide2.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(5), Inches(5)).text_frame
    maddeler = [
        f"📍 Konum: İzmir / {ilce} / {mahalle}",
        f"📐 Alan: {m2} m²",
        f"📜 İmar: {imar}",
        "⚡ Altyapı: Elektrik & Su hatları hazır.",
        f"💰 Alım Bedeli: {arsa_fiyati:,.0f} TL".replace(",", ".")
    ]
    for m in maddeler:
        p_m = tf2.add_paragraph(); p_m.text = m; p_m.font.size = Pt(16); p_m.font.color.rgb = KOYU_GRI; p_m.space_after = Pt(12)
        
    # EĞER KULLANICI FOTOĞRAF YÜKLEDİYSE Sağ tarafa otomatik yerleştir
    if yuklenen_resim is not None:
        image_stream = io.BytesIO(yuklenen_resim.read())
        # Sağ tarafa resmi hizala (Sol: 5.8 inç, Üst: 1.5 inç, Genişlik: 3.8 inç)
        slide2.shapes.add_picture(image_stream, Inches(5.8), Inches(1.5), width=Inches(3.8))

    # SLAYT 3: SEÇİLEN KONSEPT VE MALIYET MOTORU
    slide3 = prs.slides.add_slide(prs.slide_layouts[6])
    t_box3 = slide3.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(1))
    t_box3.text_frame.paragraphs[0].text = f"ÖNERİLEN PROJE: {konsept.upper()}"
    t_box3.text_frame.paragraphs[0].font.size = Pt(26); t_box3.text_frame.paragraphs[0].font.bold = True; t_box3.text_frame.paragraphs[0].font.color.rgb = LACIVERT
    
    tf3 = slide3.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(9), Inches(5)).text_frame
    
    # Matematiksel Maliyet Motoru
    taban_alan = int(int(m2) * 0.15)
    toplam_insaat = taban_alan * 2
    insaat_maliyeti = toplam_insaat * 30000
    toplam_maliyet = arsa_fiyati + insaat_maliyeti
    tahmini_bitis_degeri = toplam_maliyet * 1.6
    
    # Konsepte göre özel pazarlama cümleleri dinamik değişiyor
    if konsept == "Premium Taş Ev":
        yorum = "💡 Bölgenin doğal dokusuna uygun taş mimari, eskidikçe değer kazanan en yüksek prim potansiyeline sahiptir."
    elif konsept == "Eko-Tiny House Yaşam Alanı":
        yorum = "💡 Altyapı hazır olduğu için inşaat stresi olmadan hemen yarın kurulabilir, lüks ve mobil bir kaçış noktasıdır."
    else: # Yüksek Getirili Villa
        yorum = "💡 Bölgedeki turizm ve villa talebini nakde çevirecek, kısa/uzun dönem yüksek kira getirili yatırım senaryosudur."

    finansallar = [
        yorum,
        f"🏗️ Planlanan İnşaat Alanı: {toplam_insaat} m² (Taban: {taban_alan} m²)",
        f"💵 Arsa Bedeli: {arsa_fiyati:,.0f} TL".replace(",", "."),
        f"🔨 Tahmini Yapım Maliyeti: {insaat_maliyeti:,.0f} TL".replace(",", "."),
        f"📊 Toplam Yatırım Bütçesi: {toplam_maliyet:,.0f} TL".replace(",", "."),
        f"📈 Tamamlandığında Tahmini Piyasa Değeri: {tahmini_bitis_degeri:,.0f} TL".replace(",", ".")
    ]
    for f in finansallar:
        p_f = tf3.add_paragraph(); p_f.text = f; p_f.font.size = Pt(16); p_f.font.color.rgb = KOYU_GRI; p_f.space_after = Pt(10)

    binary_output = io.BytesIO()
    prs.save(binary_output)
    binary_output.seek(0)
    return binary_output

# -------------------------------------------------------------
# WEB ARAYÜZÜ (STREAMLIT)
# -------------------------------------------------------------
st.set_page_config(page_title="ON Türkiye Sunum Sihirbazı v2", page_icon="🏢", layout="centered")

st.title("🏢 ON Türkiye Sunum Sihirbazı v2")
st.write("Gelişmiş özelliklerle portföyünü hazırla, fotoğrafını ekle ve sunumu indir!")
st.divider()

col1, col2 = st.columns(2)

with col1:
    ilce = st.text_input("İlçe Name", value="Urla")
    m2 = st.number_input("Metrekare (m²)", min_value=1, value=500)
    danisman = st.text_input("Danışman Adı Soyadı", value="Fatih Türkdönmez")

with col2:
    mahalle = st.text_input("Mahalle / Köy", value="Kuşçular")
    imar = st.text_input("İmar Durumu", value="%15/30 Konut İmarlı")
    arsa_fiyati = st.number_input("Arsa Fiyatı (TL)", min_value=0, value=6500000, step=50000)

st.divider()

# YENİ ÖZELLİKLER: Konsept Seçimi ve Fotoğraf Yükleme Girdileri
konsept = st.selectbox("Önerilecek Proje Konsepti", ["Premium Taş Ev", "Eko-Tiny House Yaşam Alanı", "Yüksek Getirili Villa"])
yuklenen_resim = st.file_uploader("Arsa / Drone Fotoğrafı Yükleyin (Opsiyonel)", type=["jpg", "jpeg", "png"])

st.divider()

if st.button("🚀 Profesyonel Sunum Dosyasını Hazırla", use_container_width=True):
    with st.spinner("Yyapay zeka verileri ve fotoğrafları işliyor..."):
        sunum_dosyasi = gelismis_sunum_uret(ilce, mahalle, str(m2), imar, arsa_fiyati, danisman, konsept, yuklenen_resim)
        st.success("🎉 Harika! Sunumunuz başarıyla fotoğrafınız ve seçtiğiniz konseptle harmanlandı!")
        
        st.download_button(
            label="📥 Gelişmiş PowerPoint Dosyasını İndir",
            data=sunum_dosyasi,
            file_name=f"ON_Turkiye_{ilce}_{mahalle}_Gelismiş_Rapor.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True
        )