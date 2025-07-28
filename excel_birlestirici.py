import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="Stil Diva - Depo Sihirbazı", layout="centered")

st.title("Stil Diva - Depo Sihirbazı")
st.write("Lütfen günlük sipariş excelinizi yükleyiniz. Raf referans dosyası sisteme dahildir.")

st.header("1. Günlük Sipariş Excelini Yükleyin")
gunluk_excel = st.file_uploader("Pixa Sipariş Excelini Yükleyin", type=["xlsx", "xls"])

if gunluk_excel:
    st.success("Sipariş dosyası başarıyla yüklendi. İşlem için butona tıklayın.")

    if st.button("İşlemi Başlat ve Yeni Excel'i Oluştur"):
        try:
            df_referans = pd.read_excel('raf_master.xlsx', engine="calamine")
            df_gunluk = pd.read_excel(gunluk_excel, engine="calamine")

            # Referans Excel'den gerekli sütunlar seçiliyor.
            df_referans_secili = df_referans[["Barkod", "Model", "Seçenek", "Raf Adresi"]].copy()
            df_referans_secili.drop_duplicates(subset="Barkod", inplace=True)
            
            # --- DEĞİŞİKLİK: Gerçek barkodları içeren sütunun adının 'Barkod' olduğunu varsayıyoruz. ---
            # Lütfen sipariş excelinizdeki DOĞRU barkodları içeren sütunun (AK sütunu) başlığının
            # tam olarak "Barkod" olduğundan emin olun. Farklıysa, tırnak içindeki ifadeyi değiştirin.
            gercek_barkod_sutun_adi = "Barkod"
            df_gunluk_secili = df_gunluk[["Sipariş No", "Platform", gercek_barkod_sutun_adi, "Miktar"]].copy()

            # --- DEĞİŞİKLİK: Birleştirme artık doğru barkod sütunları arasında yapılıyor. ---
            # İki tabloda da sütun adının "Barkod" olduğunu varsayarak birleştirme yapılıyor.
            df_yeni = pd.merge(df_gunluk_secili, df_referans_secili, on="Barkod", how="left")

            # Sonuçtaki sütun sırası ayarlanıyor.
            df_sonuc = df_yeni[[
                "Sipariş No",
                "Platform",
                "Barkod", # Doğru barkod sütunu
                "Miktar",
                "Model",      
                "Seçenek",    
                "Raf Adresi" 
            ]]

            # Çıktıdaki başlık isimleri ayarlanıyor.
            df_sonuc.columns = [
                "A) Sipariş Numarası",
                "B) Platform",
                "C) Barkod",
                "D) Miktar",
                "E) Model",
                "F) Seçenek",
                "G) Raf Adresi"
            ]

            st.success("Birleştirme işlemi tamamlandı! Aşağıdaki butondan indirebilirsiniz.")
            st.dataframe(df_sonuc)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_sonuc.to_excel(writer, index=False, sheet_name='Birlestirilmis_Liste')
            
            processed_data = output.getvalue()

            suanki_zaman = datetime.now()
            tarih_damgasi = suanki_zaman.strftime("%d-%m-%Y")
            dinamik_dosya_adi = f"Stil Diva Sipariş - {tarih_damgasi}.xlsx"

            st.download_button(
                label="✅ C) Yeni Excel Dosyasını İndir",
                data=processed_data,
                file_name=dinamik_dosya_adi,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except KeyError as e:
            st.error(f"Sütun Bulunamadı Hatası: {e}. Sipariş dosyanızda belirttiğiniz isimde bir sütun bulunamadı. Lütfen kod içerisindeki 'gercek_barkod_sutun_adi' değişkenini kontrol edin.")
        except Exception as e:
            st.error(f"Beklenmedik bir hata oluştu: {e}")
            st.warning("Lütfen Excel dosyalarınızdaki sütun adlarının veya formatlarının doğru olduğundan emin olun.")
