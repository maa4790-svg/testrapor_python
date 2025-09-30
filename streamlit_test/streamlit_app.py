#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Payzgate Veri Çekme Web App - Streamlit
"""

import streamlit as st
import pandas as pd
import time
import os
from datetime import datetime
from main import PayzgateScraper, save_multiple_orders_to_excel

# Sayfa konfigürasyonu
st.set_page_config(
    page_title="Payzgate Veri Çekme Botu",
    page_icon="🤖",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS stilleri
st.markdown("""
<style>
    .main-header {
        text-align: center;
        padding: 1rem 0;
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        color: white;
        border-radius: 10px;
        margin-bottom: 2rem;
    }
    .success-box {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 5px;
        padding: 1rem;
        margin: 1rem 0;
    }
    .error-box {
        background-color: #f8d7da;
        border: 1px solid #f5c6cb;
        border-radius: 5px;
        padding: 1rem;
        margin: 1rem 0;
    }
    .info-box {
        background-color: #d1ecf1;
        border: 1px solid #bee5eb;
        border-radius: 5px;
        padding: 1rem;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

def main():
    # Ana başlık
    st.markdown("""
    <div class="main-header">
        <h1>🤖 Payzgate Veri Çekme Botu</h1>
        <p>Çoklu sipariş verilerini çekin ve Excel olarak indirin</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar - Bilgi
    with st.sidebar:
        st.header("ℹ️ Bilgi")
        st.info("""
        **Kullanım:**
        1. Sipariş ID'lerini girin
        2. "Veri Çek" butonuna tıklayın
        3. İşlem tamamlandığında Excel'i indirin
        """)
        
        st.header("📝 ID Formatları")
        st.code("""
        Tek ID: 6172936
        Çoklu: 6172936,6172937,6172938
        Aralık: 6172936-6172940
        Karışık: 6172936,6172938-6172940,6172945
        """)
        
        st.header("🔐 Login Bilgileri")
        st.success("Login bilgileri otomatik olarak kullanılıyor")
    
    # Ana form
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.header("📋 Sipariş ID'leri")
        
        # ID girişi
        order_ids_input = st.text_area(
            "Sipariş ID'lerini girin:",
            placeholder="Örnek: 6172936,6172937,6172938\nveya\n6172936-6172940",
            height=100
        )
        
        # Test ID'leri butonu
        if st.button("🧪 Test ID'leri Kullan"):
            order_ids_input = "6172936,6172937,6172938"
            st.rerun()
    
    with col2:
        st.header("⚙️ Ayarlar")
        
        # İşlem ayarları
        show_progress = st.checkbox("Progress göster", value=True)
        auto_download = st.checkbox("Otomatik indirme", value=True)
        
        # Veri çekme butonu
        if st.button("🚀 Veri Çek", type="primary", use_container_width=True):
            if order_ids_input.strip():
                process_orders(order_ids_input.strip(), show_progress, auto_download)
            else:
                st.error("❌ Lütfen en az bir sipariş ID'si girin!")
    
    # Footer
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #666; padding: 1rem;">
        <p>🤖 Payzgate Veri Çekme Botu v1.0 | Made with ❤️ using Streamlit</p>
    </div>
    """, unsafe_allow_html=True)

def process_orders(order_ids_input, show_progress, auto_download):
    """Sipariş işleme fonksiyonu"""
    
    # ID'leri parse et
    try:
        order_ids = parse_order_ids(order_ids_input)
        if not order_ids:
            st.error("❌ Geçersiz ID formatı!")
            return
            
        st.success(f"✅ {len(order_ids)} adet sipariş ID'si bulundu")
        
    except Exception as e:
        st.error(f"❌ ID parse hatası: {str(e)}")
        return
    
    # Progress bar
    if show_progress:
        progress_bar = st.progress(0)
        status_text = st.empty()
    
    # Scraper oluştur
    try:
        email = None
        password = None

        # Önce Streamlit secrets, sonra ENV değişkenleri
        try:
            email = st.secrets.get("PAYZGATE_EMAIL", None)
            password = st.secrets.get("PAYZGATE_PASSWORD", None)
        except Exception:
            pass

        if not email:
            email = os.environ.get("PAYZGATE_EMAIL")
        if not password:
            password = os.environ.get("PAYZGATE_PASSWORD")

        if not email or not password:
            st.error("🔐 Giriş bilgileri eksik. Lütfen `PAYZGATE_EMAIL` ve `PAYZGATE_PASSWORD` değerlerini Secrets veya ortam değişkeni olarak tanımlayın.")
            return

        scraper = PayzgateScraper(email=email, password=password)
        
        # Veri çekme işlemi
        all_orders_data = []
        
        for i, order_id in enumerate(order_ids):
            if show_progress:
                progress = (i + 1) / len(order_ids)
                progress_bar.progress(progress)
                status_text.text(f"İşleniyor: {order_id} ({i+1}/{len(order_ids)})")
            
            # Debug bilgisi
            st.write(f"🔍 Debug: {order_id} için veri çekiliyor...")
            
            # Veri çek
            order_data = scraper.get_order_data(order_id)
            
            # Debug sonucu
            if order_data:
                st.write(f"✅ {order_id}: Veri başarıyla çekildi")
                all_orders_data.append(order_data)
            else:
                st.write(f"❌ {order_id}: Veri çekilemedi")
            
            # Kısa bekleme
            time.sleep(0.5)
        
        # Sonuçları göster
        if all_orders_data:
            # Excel oluştur
            filename = f"payzgate_orders_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            save_multiple_orders_to_excel(all_orders_data, filename)
            
            # Başarı mesajı
            st.markdown("""
            <div class="success-box">
                <h3>✅ İşlem Tamamlandı!</h3>
                <p><strong>Başarılı:</strong> {}/{} sipariş</p>
                <p><strong>Dosya:</strong> {}</p>
            </div>
            """.format(len(all_orders_data), len(order_ids), filename), unsafe_allow_html=True)
            
            # Veri önizleme
            st.header("📊 Veri Önizleme")
            preview_data = []
            for order_data in all_orders_data:
                if order_data and 'specific_data' in order_data:
                    specific_data = order_data['specific_data']
                    preview_data.append({
                        'Sipariş ID': order_data['order_id'],
                        'Status': specific_data['status'],
                        'Customer': specific_data['customer'],
                        'Order Subtotal': specific_data['order_subtotal'],
                        'Order Total': specific_data['order_total'],
                        'Reference Code': specific_data['reference_code'][:20] + '...' if len(specific_data['reference_code']) > 20 else specific_data['reference_code']
                    })
            
            if preview_data:
                df = pd.DataFrame(preview_data)
                st.dataframe(df, use_container_width=True)
                
                # İndirme butonu
                if os.path.exists(filename):
                    with open(filename, 'rb') as f:
                        st.download_button(
                            label="📁 Excel Dosyasını İndir",
                            data=f.read(),
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
            
        else:
            st.markdown("""
            <div class="error-box">
                <h3>❌ Hiç Veri Çekilemedi!</h3>
                <p>Lütfen ID'leri kontrol edin ve tekrar deneyin.</p>
            </div>
            """, unsafe_allow_html=True)
    
    except Exception as e:
        st.error(f"❌ Hata oluştu: {str(e)}")

def parse_order_ids(order_ids_input):
    """ID'leri parse eder"""
    order_ids = []
    
    # Virgülle ayrılmış kısımları işle
    parts = order_ids_input.split(',')
    
    for part in parts:
        part = part.strip()
        
        if '-' in part:
            # Aralık formatı: 6172936-6172940
            try:
                start, end = part.split('-', 1)
                start = int(start.strip())
                end = int(end.strip())
                
                for i in range(start, end + 1):
                    order_ids.append(str(i))
            except ValueError:
                st.error(f"❌ Geçersiz aralık formatı: {part}")
                continue
        else:
            # Tek ID
            if part.isdigit():
                order_ids.append(part)
            else:
                st.error(f"❌ Geçersiz ID formatı: {part}")
                continue
    
    return order_ids

if __name__ == "__main__":
    main()
