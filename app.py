import streamlit as st
import pandas as pd
from datetime import datetime
import re

# Configuração da Página
st.set_page_config(page_title="Apontamento Raízen", page_icon="⚙️", layout="centered")

# --- FUNÇÕES DE APOIO ---

@st.cache_data
def carregar_dados():
    try:
        df_colab = pd.read_excel("Manutencao_App.xlsx", sheet_name="Colaboradores")
        df_ordens = pd.read_excel("Manutencao_App.xlsx", sheet_name="BDOrdens")
        
        df_colab.columns = df_colab.columns.str.strip()
        df_colab['Oficina'] = df_colab['Oficina'].astype(str).str.strip()
        df_colab['Nome'] = df_colab['Nome'].astype(str).str.strip()
        
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
    data_formatada_br = data_selecionada.strftime('%d/%m/%Y')
    
    st.write(f"**Colaborador:** {colaborador}")
    st.write(f"**Data:** {data_formatada_br}")
    st.divider()

    try:
        df_apont = pd.read_excel("Manutencao_App.xlsx", sheet_name="Apontamentos")
        # Converte Data garantindo que o formato seja comparável
        df_apont['Data'] = pd.to_datetime(df_apont['Data'], dayfirst=True).dt.date
        
        filtro = df_apont[(df_apont['Colaborador'] == colaborador) & (df_apont['Data'] == data_selecionada)].copy()

        if not filtro.empty:
            # Formata Ordem e Operação para Texto Inteiro (sem vírgulas ou .0)
            filtro['Ordem'] = filtro['Ordem'].fillna(0).astype(int).astype(str)
            filtro['Operação'] = filtro['Operação'].fillna(0).astype(int).astype(str)
            
            st.table(filtro[['Ordem', 'Operação', 'Início', 'Fim', 'Progresso']])
            
            total_horas = 0
            for _, row in filtro.iterrows():
                ini = converter_para_horas(row['Início'])
                fim = converter_para_horas(row['Fim'])
                total_horas += (fim - ini)
            
            st.metric("Total de Horas Apontadas no Dia", f"{total_horas:.2f} h")
        else:
            st.info(f"Nenhum registro encontrado para {data_formatada_br}.")
    except Exception:
        st.info("Aba de Apontamentos está vazia ou ainda não foi criada.")

# --- INTERFACE PRINCIPAL ---

df_colab, df_ordens = carregar_dados()

st.header("⚙️ Apontamento de Manutenção")

if not df_colab.empty:
    opcoes_oficina = sorted(df_colab['Oficina'].unique())
    oficina = st.selectbox("Selecione a Oficina", opcoes_oficina)
    nomes_filtrados = df_colab[df_colab['Oficina'].str.upper() == oficina.upper()]['Nome'].tolist()
    colaborador = st.selectbox("Nome do Colaborador", sorted(nomes_filtrados) if nomes_filtrados else ["Nenhum"])
else:
    st.error("Planilha 'Colaboradores' não encontrada.")

col1, col2 = st.columns(2)
ordem_input = col1.text_input("Número da Ordem (8 dígitos)", max_chars=8)
operacao_input = col2.text_input("Operação", max_chars=3)

nome_atividade = ""
if ordem_input and operacao_input:
    try:
        res = df_ordens[(df_ordens['Ordem'].astype(str).str.strip() == ordem_input.strip()) & 
                        (df_ordens['Operação'].astype(str).str.strip() == operacao_input.strip())]
        if not res.empty:
            nome_atividade = res['Txt.breve operação'].values[0]
            st.success(f"📌 Atividade: {nome_atividade}")
        else:
            st.warning("⚠️ Ordem/Operação não encontrada.")
    except:
        pass

c_data, c_btn = st.columns([2, 1])
data_ativ = c_data.date_input("Data da Atividade", datetime.now(), format="DD/MM/YYYY")

if c_btn.button("Verificar Apontamento", use_container_width=True):
    mostrar_resumo(colaborador, data_ativ)

c_h1, c_h2 = st.columns(2)
h_inicio = c_h1.text_input("Início (HH:MM)", placeholder="00:00")
h_fim = c_h2.text_input("Fim (HH:MM)", placeholder="00:00")

andamento = st.slider("Porcentagem Executada", 0, 100, step=5)
descricao = st.text_area("Descrição da Atividade")

# --- BOTÃO GRAVAR (VERSÃO CORRIGIDA) ---

if st.button("Gravar Apontamento", use_container_width=True):
    regex_hora = r"^([0-1]?[0-9]|2[0-3]):([0-5][0-9])$"

    if not nome_atividade:
        st.error("Erro: Ordem/Operação inválida.")
    elif not re.match(regex_hora, h_inicio) or not re.match(regex_hora, h_fim):
        st.error("Formato de hora inválido.")
    else:
        # Prepara a nova linha
        novo_dado = {
            "Oficina": [oficina], 
            "Colaborador": [colaborador], 
            "Ordem": [ordem_input],
            "Operação": [operacao_input], 
            "Data": [data_ativ.strftime('%d/%m/%Y')],
            "Início": [h_inicio], 
            "Fim": [h_fim], 
            "Progresso": [f"{andamento}%"], 
            "Descrição": [descricao]
        }
        df_novo = pd.DataFrame(novo_dado)
        
        try:
            # Tenta ler a planilha; se falhar ou estiver vazia, cria a estrutura do zero
            try:
                df_atual = pd.read_excel("Manutencao_App.xlsx", sheet_name="Apontamentos")
            except Exception:
                df_atual = pd.DataFrame(columns=novo_dado.keys())
            
            # Garante que mesmo se a aba existir mas estiver sem dados, o código funcione
            if df_atual.empty:
                df_final = df_novo
            else:
                df_final = pd.concat([df_atual, df_novo], ignore_index=True)
            
            # Grava no Excel
            with pd.ExcelWriter("Manutencao_App.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                df_final.to_excel(writer, sheet_name="Apontamentos", index=False)
            
            st.balloons()
            st.success("Registrado com sucesso!")
            
            # IMPORTANTE: Limpa o cache para o resumo atualizar na hora
            st.cache_data.clear()
            
        except Exception as e:
            st.error(f"Erro ao salvar: Verifique se o Excel está aberto! Erro: {e}")
