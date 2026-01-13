import streamlit as st
import pandas as pd
import os
import requests
from bs4 import BeautifulSoup
import py7zr
from io import BytesIO
from datetime import datetime
import time
import glob
import dask.dataframe as dd
import zipfile
import shutil
from urllib.parse import quote

# --- CONFIGURAÇÃO INICIAL E PROTEÇÃO ---
st.set_page_config(page_title="Automação CAGED", layout="wide")
st.markdown("""<style>[data-testid="stSidebarNav"] {display: none;}</style>""", unsafe_allow_html=True)

try:
    from auth import draw_sidebar
    draw_sidebar()
except ImportError:
    pass

if not st.session_state.get('logged_in', False) or st.session_state.get('role') != 'admin':
    st.error("Acesso negado.")
    st.stop()

# --- CONSTANTES ---
# (Você pode enriquecer este dicionário com todos os dtypes originais)
DTYPES_BASE = {
    'município': 'str', 'seção': 'str', 'subclasse': 'str',
    'cbo2002ocupação': 'str', 'saldomovimentação': 'str',
    'competênciamov': 'str', 'uf': 'str'
}

DTYPES_MAP = {
    "Movimentações": DTYPES_BASE,
    "Fora de prazo": DTYPES_BASE,
    "Exclusões": DTYPES_BASE
}

# --- FUNÇÕES DE REDE (HTTP SCRAPING) ---

def listar_arquivos_http(url_diretorio):
    """
    Acessa uma URL de diretório (ex: .../202401/) e retorna a lista de nomes de arquivos (.7z) encontrados.
    Isso resolve o problema de case-sensitivity (arquivos com .7z ou .7Z).
    """
    try:
        response = requests.get(url_diretorio, timeout=10)
        if response.status_code != 200:
            return []
        
        soup = BeautifulSoup(response.text, 'html.parser')
        arquivos = []
        
        # Procura por links (tags <a>)
        for link in soup.find_all('a'):
            href = link.get('href')
            # Filtra apenas arquivos 7z
            if href and (href.lower().endswith('.7z')):
                arquivos.append(href)
        
        return arquivos
    except Exception as e:
        print(f"Erro ao listar diretório HTTP: {e}")
        return []

def baixar_arquivo_http(url_arquivo, destino_local):
    """Baixa o arquivo via HTTP com stream para não estourar a memória."""
    try:
        with requests.get(url_arquivo, stream=True, timeout=60) as r:
            r.raise_for_status()
            with open(destino_local, 'wb') as f:
                for chunk in r.iter_content(chunk_size=8192):
                    f.write(chunk)
        return True, "Sucesso"
    except Exception as e:
        return False, str(e)

# --- PROCESSAMENTO PRINCIPAL ---

def processar_caged(tipo_caged, ano, mes_inicial, mes_final, pasta_temp, tipos_selecionados):
    # Mapeamento: Transforma o caminho FTP em URL HTTP
    # O servidor ftp.mtps.gov.br responde na porta 80 (HTTP) espelhando a estrutura
    base_http = "http://ftp.mtps.gov.br/pdet/microdados"
    
    mapa_caminhos = {
        "NOVO CAGED": f"{base_http}/NOVO%20CAGED/{ano}/",
        "CAGED_AJUSTES": f"{base_http}/NOVO%20CAGED/AJUSTES/{ano}/",
        "CAGED (Antigo)": f"{base_http}/CAGED/{ano}/"
    }
    
    meses = range(mes_inicial, mes_final + 1) if tipo_caged != "CAGED (Antigo)" else [0]
    
    for mes in meses:
        status_mes = st.empty()
        
        # Monta a URL do diretório do mês
        url_diretorio = mapa_caminhos.get(tipo_caged)
        if tipo_caged in ["NOVO CAGED", "CAGED_AJUSTES"]:
            url_diretorio += f"{ano}{mes:02d}/"
            
        status_mes.info(f"🔎 Consultando diretório: {url_diretorio}")
        
        # Lista arquivos disponíveis na página HTML
        arquivos_no_servidor = listar_arquivos_http(url_diretorio)
        
        if not arquivos_no_servidor:
            st.warning(f"Diretório não encontrado ou vazio: {url_diretorio}")
            continue

        for prefixo, nome_amigavel in tipos_selecionados.items():
            # Define o padrão do nome esperado (ex: CAGEDMOV202401.7z)
            if tipo_caged == "CAGED (Antigo)":
                padrao_nome = f"{prefixo}{ano}"
            else:
                padrao_nome = f"{prefixo}{ano}{mes:02d}"
            
            # Encontra o nome exato no servidor (case insensitive)
            nome_real = next((f for f in arquivos_no_servidor if padrao_nome.lower() in f.lower()), None)
            
            if not nome_real:
                st.warning(f"Arquivo {padrao_nome} não encontrado na lista do servidor.")
                continue
                
            url_download = f"{url_diretorio}{nome_real}"
            path_local_7z = os.path.join(pasta_temp, nome_real)
            
            # 1. DOWNLOAD (HTTP)
            status_mes.info(f"⬇️ Baixando {nome_real} (HTTP)...")
            ok, msg = baixar_arquivo_http(url_download, path_local_7z)
            
            if not ok:
                st.error(f"Erro no download de {nome_real}: {msg}")
                continue
            
            # 2. EXTRAÇÃO
            status_mes.info(f"📦 Extraindo {nome_real}...")
            pasta_extracao = os.path.join(pasta_temp, f"temp_{padrao_nome}")
            os.makedirs(pasta_extracao, exist_ok=True)
            
            try:
                with py7zr.SevenZipFile(path_local_7z, mode='r') as z:
                    z.extractall(path=pasta_extracao)
            except Exception as e:
                st.error(f"Arquivo corrompido ou erro na extração: {e}")
                continue
            
            # 3. CONVERSÃO E LIMPEZA
            status_mes.info(f"🔄 Processando CSV...")
            try:
                # Localiza o TXT extraído
                txt_file = next((f for f in os.listdir(pasta_extracao) if f.lower().endswith('.txt')), None)
                
                if txt_file:
                    path_txt = os.path.join(pasta_extracao, txt_file)
                    
                    # Lê com Dtypes reduzidos para economizar memória e salva como CSV padrão
                    # (Aqui você usaria seu DTYPES_BASE completo)
                    df = pd.read_csv(path_txt, sep=';', encoding='latin-1', dtype=str)
                    
                    nome_csv_final = f"caged_{nome_amigavel}_{ano}_{mes:02d}.csv"
                    path_csv_final = os.path.join(pasta_temp, nome_csv_final)
                    
                    df.to_csv(path_csv_final, sep=';', index=False, encoding='utf-8-sig')
                else:
                    st.warning(f"Nenhum .txt encontrado dentro de {nome_real}")

            except Exception as e:
                st.error(f"Erro ao processar dados: {e}")
            finally:
                # Remove o 7z e a pasta temporária de extração para não lotar o servidor
                if os.path.exists(path_local_7z): os.remove(path_local_7z)
                if os.path.exists(pasta_extracao): shutil.rmtree(pasta_extracao)
        
        status_mes.empty()

# --- INTERFACE ---
st.title("🤖 Automação CAGED (Protocolo HTTP)")
st.info("Extração de dados via túnel HTTP (Porta 80) para contornar bloqueio de FTP.")

col_conf, col_mes = st.columns(2)

with col_conf:
    tipo_caged = st.selectbox("Base de Dados:", ["NOVO CAGED", "CAGED_AJUSTES", "CAGED (Antigo)"])
    anos = st.multiselect("Ano(s):", list(range(2020, 2026)), default=[2024])

with col_mes:
    mes_ini = st.number_input("Mês Inicial", 1, 12, 1)
    mes_fim = st.number_input("Mês Final", 1, 12, 1)

tipos_arquivos = {
    'CAGEDMOV': 'Movimentações',
    'CAGEDEXC': 'Exclusões',
    'CAGEDFOR': 'Fora de prazo'
}
selecao = st.multiselect("Tipos de Arquivo:", list(tipos_arquivos.values()), default=["Movimentações"])
tipos_finais = {k: v for k, v in tipos_arquivos.items() if v in selecao}

if st.button("🚀 Iniciar Extração na Nuvem", type="primary"):
    pasta_temp = f"dados_caged_{int(time.time())}"
    os.makedirs(pasta_temp, exist_ok=True)
    
    progresso = st.progress(0)
    
    try:
        for i, ano in enumerate(anos):
            processar_caged(tipo_caged, ano, mes_ini, mes_fim, pasta_temp, tipos_finais)
            progresso.progress((i + 1) / len(anos))
        
        st.success("Processamento concluído com sucesso!")
        
        # GERAÇÃO DO ZIP FINAL
        st.info("Compactando resultados...")
        zip_buffer = BytesIO()
        arquivos_csv = [f for f in os.listdir(pasta_temp) if f.endswith('.csv')]
        
        if arquivos_csv:
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for f in arquivos_csv:
                    zf.write(os.path.join(pasta_temp, f), f)
            
            st.download_button(
                label="📥 Baixar Dados Processados (.zip)",
                data=zip_buffer.getvalue(),
                file_name=f"caged_extracao_{datetime.now().strftime('%d%m%Y')}.zip",
                mime="application/zip",
                type="primary"
            )
        else:
            st.warning("Nenhum dado foi processado (verifique se os meses selecionados já estão disponíveis no site).")
            
    except Exception as e:
        st.error(f"Erro crítico: {e}")
    finally:
        # Limpeza final
        if os.path.exists(pasta_temp): shutil.rmtree(pasta_temp)