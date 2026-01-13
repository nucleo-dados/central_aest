import asyncio
import sys
import os
import subprocess
import json
import time
import requests
import streamlit as st
import zipfile
import re
from io import BytesIO
from playwright.sync_api import sync_playwright

# --- 1. INTEGRAÇÃO ---
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

try:
    from auth import page_protector 
    page_protector(page_name="Exportador Power BI")
except ImportError:
    pass

st.set_page_config(page_title="Exportador Power BI", page_icon="📊", layout="centered")

if sys.platform == 'win32':
    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

# --- 2. SETUP ---
def log(mensagem):
    print(f"[PowerBI Bot] {mensagem}")

@st.cache_resource
def install_playwright_browser():
    try:
        subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], check=True)
    except Exception as e:
        st.error(f"Erro setup navegador: {e}")

install_playwright_browser()

# --- 3. DADOS IBGE ---
@st.cache_data(ttl=86400)
def carregar_mapa_municipio_mesorregiao():
    url = "https://servicodados.ibge.gov.br/api/v1/localidades/estados/MG/municipios"
    try:
        response = requests.get(url, timeout=10)
        dados = response.json()
        mapa = {c['nome']: c['microrregiao']['mesorregiao']['nome'] for c in dados}
        return sorted(mapa.keys()), mapa
    except:
        return ["Belo Horizonte"], {"Belo Horizonte": "Metropolitana de Belo Horizonte"}

# --- 4. ROBÔ DE EXPORTAÇÃO ---
def executar_exportacao(url_relatorio, termo_filtro, output_folder, tipo_relatorio):
    if not os.path.exists(output_folder): os.makedirs(output_folder)
    
    debug_folder = "debug_files"
    if not os.path.exists(debug_folder): os.makedirs(debug_folder)

    nome_limpo = termo_filtro.replace(' ', '_')
    nome_arquivo = f"Relatorio_{tipo_relatorio}_{nome_limpo}.pdf"
    caminho_final = os.path.join(output_folder, nome_arquivo)

    # Auth
    caminho_auth = "auth.json" 
    if not os.path.exists(caminho_auth):
        caminho_auth = os.path.join("..", "auth.json")
        if not os.path.exists(caminho_auth):
            return None, "Arquivo 'auth.json' não encontrado.", None

    url_base_limpa = url_relatorio.split("?")[0] + "?experience=power-bi"

    try:
        with sync_playwright() as p:
            # Headless True para servidor
            browser = p.chromium.launch(headless=True, args=["--no-sandbox", "--disable-gpu"])
            context = browser.new_context(storage_state=caminho_auth, viewport={"width": 1920, "height": 1080})
            page = context.new_page()

            log(f"[{tipo_relatorio}] Acessando URL...")
            page.goto(url_base_limpa, timeout=60000, wait_until="domcontentloaded")
            
            time.sleep(3) 
            if "login.microsoftonline.com" in page.url:
                browser.close()
                return None, "Sessão expirada! Rode o script de login manual.", None
            
            try: page.wait_for_load_state("networkidle", timeout=15000)
            except: pass
            
            page.wait_for_timeout(5000) 

            # --- LÓGICA DE FILTRAGEM ---
            try:
                if tipo_relatorio == "Municipal":
                    log("Modo Municipal: Buscando container...")
                    container = page.locator("div.visual").filter(has_text="Selecione o município").last
                    if container.count() == 0: container = page.locator("div.visual").filter(has_text="Município").last
                    
                    campo_busca = container.locator("input.searchInput")
                    if not campo_busca.is_visible():
                         lupa = container.locator("i.search-icon")
                         if lupa.count() > 0: lupa.click()
                         page.wait_for_timeout(500)
                    if not campo_busca.is_visible(): campo_busca = page.locator("input.searchInput").first

                    log(f"Digitando '{termo_filtro}'...")
                    campo_busca.click()
                    campo_busca.fill("") 
                    campo_busca.press_sequentially(termo_filtro, delay=100)
                    
                    alvo_item = container.locator("span.slicerText").filter(has_text=termo_filtro).first
                    try:
                        alvo_item.wait_for(state="visible", timeout=15000)
                        page.wait_for_timeout(500) 
                        alvo_item.click()
                    except Exception as e_wait:
                        page.get_by_text(termo_filtro, exact=True).first.click(timeout=3000)

                elif tipo_relatorio == "Regional":
                    log("Modo Regional: Buscando dropdown...")
                    container = page.locator("div.visual").filter(has_text="Mesorregião, Municipio").last
                    if container.count() == 0: raise Exception("Container Regional não encontrado.")

                    seta = container.locator("i.powervisuals-glyph-chevron-down")
                    if seta.count() > 0 and seta.is_visible(): seta.click()
                    else: container.click()
                    
                    page.wait_for_timeout(1500)
                    campo_busca = page.locator("input.searchInput:visible").first
                    if not campo_busca.count() > 0: campo_busca = page.locator("div.slicer-dropdown-menu input.searchInput").first

                    log(f"Digitando '{termo_filtro}'...")
                    campo_busca.click()
                    campo_busca.fill("") 
                    campo_busca.press_sequentially(termo_filtro, delay=100)
                    
                    try:
                        item_regional = page.locator("span.slicerText").filter(has_text=termo_filtro).first
                        item_regional.wait_for(state="visible", timeout=10000)
                        item_regional.click()
                    except:
                        page.get_by_text(termo_filtro, exact=True).first.click()

                log(f"[{tipo_relatorio}] Filtro aplicado.")
                page.wait_for_timeout(5000)

            except Exception as e:
                ts = int(time.time())
                err_img = os.path.join(debug_folder, f"erro_filtro_{ts}.png")
                page.screenshot(path=err_img)
                browser.close()
                return None, f"Erro filtro {tipo_relatorio}: {e}", err_img

            # --- EXPORTAÇÃO ---
            try:
                log(f"[{tipo_relatorio}] Exportando...")
                page.get_by_role("button", name="Export").click() 
                page.get_by_text("PDF").click()
                page.wait_for_selector("mat-dialog-container", timeout=20000)
                
                # Regex para Valores atuais / Current values
                padrao_texto = re.compile(r"Valores atuais|Current values", re.IGNORECASE)
                page.get_by_text(padrao_texto).click(force=True, timeout=10000)

                with page.expect_download(timeout=180000) as download_info:
                    page.locator("mat-dialog-actions").get_by_role("button", name="Export").click()
                
                download = download_info.value
                download.save_as(caminho_final)
                
            except Exception as e:
                ts = int(time.time())
                err_img = os.path.join(debug_folder, f"erro_export_{tipo_relatorio}_{ts}.png")
                err_html = os.path.join(debug_folder, f"erro_export_{tipo_relatorio}_{ts}.html")
                
                try: page.screenshot(path=err_img, full_page=True)
                except: page.screenshot(path=err_img)
                
                with open(err_html, "w", encoding="utf-8") as f:
                    f.write(page.content())

                browser.close()
                return None, f"Erro download ({tipo_relatorio}): {e}", err_img

            browser.close()
            return caminho_final, "Sucesso", None

    except Exception as e:
        return None, f"Erro crítico: {e}", None

# --- 7. INTERFACE ---

st.title("📊 Exportador Power BI")

# --- NOVO TEXTO AQUI ---
st.markdown("""
Esta automação exporta o **Relatório Executivo** do município selecionado junto com o da sua **Mesorregião** correspondente. 
Basta escolher a cidade na lista abaixo e clicar em gerar: o sistema entregará um arquivo `.zip` contendo ambos os PDFs.
""")
# -----------------------

lista_cidades, mapa_mesorregioes = carregar_mapa_municipio_mesorregiao()

# URLs
URL_MUNICIPAL = "https://app.powerbi.com/groups/me/reports/848d470e-c20f-4948-8ab3-8223d80eed5a?experience=power-bi"
URL_REGIONAL = "https://app.powerbi.com/groups/me/reports/3af1e349-cb25-4352-b2ee-84875656161a/ReportSection48a3df4dcd707d559484?experience=power-bi" 

col_sel, col_btn = st.columns([3, 1])
with col_sel:
    municipio_selecionado = st.selectbox("Município:", options=lista_cidades, index=None)

with col_btn:
    st.write("")
    st.write("")
    iniciar = st.button("Exportar", type="primary", use_container_width=True)

if iniciar:
    if not municipio_selecionado:
        st.warning("Selecione um município.")
    elif "COLOQUE_AQUI" in URL_REGIONAL:
        st.error("Configure a URL do Regional no código!")
    else:
        mesorregiao = mapa_mesorregioes.get(municipio_selecionado)
        status = st.status(f"Processando: **{municipio_selecionado}** + **{mesorregiao}**", expanded=True)
        arquivos = []

        with status:
            st.write(f"🏙️ Municipal...")
            p_m, m_m, e_m = executar_exportacao(URL_MUNICIPAL, municipio_selecionado, "temp_pdfs", "Municipal")
            
            if p_m: 
                arquivos.append(p_m)
                st.write("✅ OK")
            else: 
                st.error(f"❌ Erro Mun: {m_m}")
                if e_m:
                    st.image(e_m, caption="Erro na tela (Municipal)")

            st.write(f"🌎 Regional...")
            p_r, m_r, e_r = executar_exportacao(URL_REGIONAL, mesorregiao, "temp_pdfs", "Regional")
            
            if p_r: 
                arquivos.append(p_r)
                st.write("✅ OK")
            else: 
                st.error(f"❌ Erro Reg: {m_r}")
                if e_r:
                    st.image(e_r, caption="Erro na tela (Regional)")

            if arquivos:
                status.update(label="Concluído!", state="complete", expanded=False)
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                    for f in arquivos: zf.write(f, os.path.basename(f))
                st.success("Sucesso!")
                st.download_button("Baixar ZIP", zip_buffer.getvalue(), "Relatorios.zip", "application/zip", type="primary")
            else:
                status.update(label="Falha", state="error")
                st.error("Falha total.")

st.divider()