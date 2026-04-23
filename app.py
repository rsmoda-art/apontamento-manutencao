import streamlit as st
import pandas as pd
from datetime import datetime
import re
import os

# Configuração da Página
st.set_page_config(page_title="Apontamento Raízen", page_icon="⚙️", layout="centered")

# Descobrir o caminho real da pasta do script para não salvar em lugar errado
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
        # Lê o arquivo usando o caminho absoluto
        df_apont = pd.read_excel(caminho_excel, sheet_name="Apontamentos")
        df_apont['Data Atividade'] = pd.to_datetime(df_apont['Data Atividade'], dayfirst=True).dt.date
        
        filtro = df_apont[(df_apont['Nome Colaborador'] == colaborador) & (df_apont['Data Atividade'] == data_selecionada)].copy()

        if not filtro.empty:
            # Ajustando nomes para exibição baseada na sua foto
            filtro['Numero Ordem'] = filtro['Numero Ordem'].fillna(0).astype(int).astype(str)
            filtro['Operacao Ordem'] = filtro['Operacao Ordem'].fillna(0).astype(int).astype(str)
            
            st.table(filtro[['Numero Ordem', 'Operacao Ordem', 'Hora Inicio', 'Hora Fim', 'Porcentagem Executada']])
            
            total_horas = 0
            for _, row in filtro.iterrows():
                ini = converter_para_horas(row['Hora Inicio'])
                fim = converter_para_horas(row['Hora Fim'])
                total_horas += (fim - ini)
            
            st.metric("Total de Horas no Dia", f"{total_horas:.2f} h")
        else:
            st.info("Nenhum registro encontrado para hoje.")
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

# --- BOTÃO GRAVAR COM NOMES DE COLUNAS IGUAIS À SUA FOTO ---

if st.button("Gravar Apontamento", use_container_width=True):
    regex_hora = r"^([0-1]?[0-9]|2[0-3]):([0-5][0-9])$"

    if not nome_atividade:
        st.error("Erro: Ordem inválida.")
    elif not re.match(regex_hora, h_inicio) or not re.match(regex_hora, h_fim):
        st.error("Formato de hora inválido.")
    else:
        # IMPORTANTE: Usei os nomes EXATOS da sua foto do Excel
        novo_dado = {
            "Oficina": [oficina], 
            "Nome Colaborador": [colaborador], 
            "Numero Ordem": [ordem_input],
            "Operacao Ordem": [operacao_input], 
            "Data Atividade": [data_ativ.strftime('%d/%m/%Y')],
            "Hora Inicio": [h_inicio], 
            "Hora Fim": [h_fim], 
            "Porcentagem Executada": [f"{andamento}%"], 
            "Descricao Atividade": [descricao]
        }
        df_novo = pd.DataFrame(novo_dado)
        
        try:
            # Tenta ler a aba do arquivo correto
            try:
                df_atual = pd.read_excel(caminho_excel, sheet_name="Apontamentos")
            except:
                df_atual = pd.DataFrame(columns=novo_dado.keys())
            
            df_final = pd.concat([df_atual, df_novo], ignore_index=True)
            
            # Remove a coluna do PowerApps se ela aparecer para limpar a planilha
            if "__PowerAppsId__" in df_final.columns:
                df_final = df_final.drop(columns=["__PowerAppsId__"])

            # Grava no caminho absoluto
            with pd.ExcelWriter(caminho_excel, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                df_final.to_excel(writer, sheet_name="Apontamentos", index=False)
            
            st.balloons()
            st.success("Salvo no Excel com sucesso!")
            st.cache_data.clear()
            
        except Exception as e:
            st.error(f"FECHE O EXCEL! O Python não consegue salvar com o arquivo aberto. Erro: {e}")