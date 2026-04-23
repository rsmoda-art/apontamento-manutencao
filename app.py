import streamlit as st
import pandas as pd
from datetime import datetime
import re
import os

# Configuração da Página
st.set_page_config(page_title="Apontamento Raízen", page_icon="⚙️", layout="centered")

# Caminho absoluto para o arquivo no servidor
diretorio_atual = os.path.dirname(os.path.abspath(__file__))
caminho_excel = os.path.join(diretorio_atual, "Manutencao_App.xlsx")

# --- FUNÇÕES DE APOIO ---

@st.cache_data
def carregar_dados():
    try:
        df_colab = pd.read_excel(caminho_excel, sheet_name="Colaboradores")
        df_ordens = pd.read_excel(caminho_excel, sheet_name="BDOrdens")
        df_colab.columns = df_colab.columns.str.strip()
        return df_colab, df_ordens
    except Exception as e:
        st.error(f"Erro ao carregar banco de dados: {e}")
        return pd.DataFrame(), pd.DataFrame()

def converter_para_horas(txt_hora):
    try:
        h, m = map(int, str(txt_hora).split(':'))
        return h + (m / 60.0)
    except:
        return 0

# --- DIALOG (POP-UP) DE RESUMO ---

@st.dialog("Resumo de Apontamentos")
def mostrar_resumo(colaborador, data_selecionada):
    st.write(f"**Colaborador:** {colaborador}")
    st.write(f"**Data:** {data_selecionada.strftime('%d/%m/%Y')}")
    st.divider()

    try:
        df_apont = pd.read_excel(caminho_excel, sheet_name="Apontamentos")
        df_apont['Data'] = pd.to_datetime(df_apont['Data'], dayfirst=True).dt.date
        
        filtro = df_apont[(df_apont['Colaborador'] == colaborador) & (df_apont['Data'] == data_selecionada)].copy()

        if not filtro.empty:
            filtro['Ordem'] = filtro['Ordem'].fillna(0).astype(int).astype(str)
            filtro['Operação'] = filtro['Operação'].fillna(0).astype(int).astype(str)
            
            st.table(filtro[['Ordem', 'Operação', 'Início', 'Fim', 'Progresso']])
            
            total_horas = 0
            for _, row in filtro.iterrows():
                ini = converter_para_horas(row['Início'])
                fim = converter_para_horas(row['Fim'])
                total_horas += (fim - ini)
            
            st.metric("Total de Horas no Dia", f"{total_horas:.2f} h")
        else:
            st.info("Nenhum registro encontrado para este dia.")
    except Exception as e:
        st.error(f"Erro ao ler resumo: {e}")

# --- INTERFACE PRINCIPAL ---

df_colab, df_ordens = carregar_dados()

st.header("⚙️ Apontamento de Manutenção")

if not df_colab.empty:
    opcoes_oficina = sorted(df_colab['Oficina'].unique())
    oficina = st.selectbox("Selecione a Oficina", opcoes_oficina)
    nomes_filtrados = df_colab[df_colab['Oficina'].str.upper() == oficina.upper()]['Nome'].tolist()
    colaborador = st.selectbox("Nome do Colaborador", sorted(nomes_filtrados) if nomes_filtrados else ["Nenhum"])

col1, col2 = st.columns(2)
ordem_input = col1.text_input("Número da Ordem", max_chars=8)
operacao_input = col2.text_input("Operação", max_chars=3)

nome_atividade = ""
if ordem_input and operacao_input:
    res = df_ordens[(df_ordens['Ordem'].astype(str).str.strip() == ordem_input.strip()) & 
                    (df_ordens['Operação'].astype(str).str.strip() == operacao_input.strip())]
    if not res.empty:
        nome_atividade = res['Txt.breve operação'].values[0]
        st.success(f"📌 Atividade: {nome_atividade}")

c_data, c_btn = st.columns([2, 1])
data_ativ = c_data.date_input("Data da Atividade", datetime.now(), format="DD/MM/YYYY")

if c_btn.button("Verificar Apontamento", use_container_width=True):
    mostrar_resumo(colaborador, data_ativ)

c_h1, c_h2 = st.columns(2)
h_inicio = c_h1.text_input("Início (HH:MM)", placeholder="08:00")
h_fim = c_h2.text_input("Fim (HH:MM)", placeholder="17:00")

andamento = st.slider("Porcentagem Executada", 0, 100, step=5)
descricao = st.text_area("Descrição da Atividade")

# --- BOTÃO GRAVAR ---

if st.button("Gravar Apontamento", use_container_width=True):
    regex_hora = r"^([0-1]?[0-9]|2[0-3]):([0-5][0-9])$"
    if not nome_atividade:
        st.error("Erro: Ordem inválida.")
    elif not re.match(regex_hora, h_inicio) or not re.match(regex_hora, h_fim):
        st.error("Formato de hora inválido.")
    else:
        novo_dado = {
            "Oficina": [oficina], "Colaborador": [colaborador], "Ordem": [ordem_input],
            "Operação": [operacao_input], "Data": [data_ativ.strftime('%d/%m/%Y')],
            "Início": [h_inicio], "Fim": [h_fim], "Progresso": [f"{andamento}%"], "Descrição": [descricao]
        }
        df_novo = pd.DataFrame(novo_dado)
        try:
            with pd.ExcelFile(caminho_excel) as xls:
                df_colab_orig = pd.read_excel(xls, "Colaboradores")
                df_ordens_orig = pd.read_excel(xls, "BDOrdens")
                try:
                    df_apont_atual = pd.read_excel(xls, "Apontamentos")
                except:
                    df_apont_atual = pd.DataFrame(columns=novo_dado.keys())
            
            df_apont_final = pd.concat([df_apont_atual, df_novo], ignore_index=True)
            if "__PowerAppsId__" in df_apont_final.columns:
                df_apont_final = df_apont_final.drop(columns=["__PowerAppsId__"])

            with pd.ExcelWriter(caminho_excel, engine="openpyxl") as writer:
                df_apont_final.to_excel(writer, sheet_name="Apontamentos", index=False)
                df_colab_orig.to_excel(writer, sheet_name="Colaboradores", index=False)
                df_ordens_orig.to_excel(writer, sheet_name="BDOrdens", index=False)
            
            st.balloons()
            st.success("Salvo com sucesso no servidor!")
            st.cache_data.clear()
        except Exception as e:
            st.error(f"Erro ao salvar: {e}")

# --- BOTÃO DE DOWNLOAD (NOVIDADE) ---
st.divider()
st.subheader("📦 Gestão de Dados")
try:
    with open(caminho_excel, "rb") as f:
        st.download_button(
            label="📥 Baixar Planilha de Apontamentos Atualizada",
            data=f,
            file_name=f"Apontamentos_Manutencao_{datetime.now().strftime('%d_%m_%Y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
except Exception as e:
    st.warning("Aguardando o primeiro apontamento para gerar o download.")