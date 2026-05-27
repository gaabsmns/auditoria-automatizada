"""
APP DE AUDITORIA AUTOMATIZADA

VERSAO RESISTENTE A AMBIENTES SEM:
- streamlit
- python-docx

Porque dependencia quebrada e praticamente patrimonio cultural da programacao.
"""

from io import BytesIO
from pathlib import Path
import argparse
import os
import pandas as pd

# ==================================================
# STREAMLIT OPCIONAL
# ==================================================
try:
    import streamlit as st
    STREAMLIT_AVAILABLE = True
except ModuleNotFoundError:
    STREAMLIT_AVAILABLE = False

# ==================================================
# PYTHON-DOCX OPCIONAL
# ==================================================
try:
    from docx import Document
    DOCX_AVAILABLE = True
except ModuleNotFoundError:
    DOCX_AVAILABLE = False

# ==================================================
# CONFIGURACAO STREAMLIT
# ==================================================
if STREAMLIT_AVAILABLE:
    st.set_page_config(
        page_title="Auditoria Automatizada",
        layout="wide"
    )

    st.title("Auditoria Automatizada de Faturamento")
    st.caption(
        "Sistema de conferencia automatica"
    )


def obter_senha_app():
    senha = os.getenv("AUDITORIA_SENHA")

    if senha:
        return senha

    if STREAMLIT_AVAILABLE:
        try:
            return st.secrets.get("AUDITORIA_SENHA")
        except Exception:
            return None

    return None


def verificar_acesso_streamlit():
    senha_configurada = obter_senha_app()

    if not senha_configurada:
        st.warning(
            "Senha nao configurada. Para publicar com seguranca, configure AUDITORIA_SENHA."
        )
        return True

    if st.session_state.get("acesso_liberado"):
        return True

    senha_digitada = st.text_input(
        "Senha de acesso",
        type="password"
    )

    if st.button("Entrar"):
        if senha_digitada == senha_configurada:
            st.session_state["acesso_liberado"] = True
            st.rerun()
        else:
            st.error("Senha incorreta")

    return False

# ==================================================
# LEITURA DO ARQUIVO
# ==================================================
def carregar_arquivo(file):
    nome = getattr(file, "name", str(file)).lower()

    if nome.endswith(".csv"):
        return pd.read_csv(file)

    return pd.read_excel(file)

# ==================================================
# NORMALIZACAO
# ==================================================
def normalizar_dados(df):
    df = df.copy()

    df.columns = [str(c).strip() for c in df.columns]

    obrigatorias = [
        "DATA",
        "HORA",
        "PACIENTE",
        "Procedimento",
        "Profissional",
        "Fornecedor"
    ]

    faltando = [
        c for c in obrigatorias
        if c not in df.columns
    ]

    if faltando:
        raise ValueError(
            f"Colunas obrigatorias ausentes: {', '.join(faltando)}"
        )

    df["DATA"] = pd.to_datetime(
        df["DATA"],
        errors="coerce"
    )

    if df["DATA"].isna().all():
        raise ValueError("Nenhuma data valida encontrada")

    df["HORA"] = (
        df["HORA"]
        .astype(str)
        .str.strip()
    )

    texto_cols = [
        "PACIENTE",
        "Procedimento",
        "Profissional",
        "Fornecedor"
    ]

    for col in texto_cols:
        df[col] = (
            df[col]
            .fillna("")
            .astype(str)
            .str.strip()
        )

    data_texto = df["DATA"].dt.strftime("%Y-%m-%d")

    df["DATAHORA"] = pd.to_datetime(
        data_texto + " " + df["HORA"],
        errors="coerce"
    )

    if "DuracaoHoras" in df.columns:
        df["DuracaoHoras"] = pd.to_numeric(
            df["DuracaoHoras"],
            errors="coerce"
        )

    if "Quantidade" in df.columns:
        df["Quantidade"] = pd.to_numeric(
            df["Quantidade"],
            errors="coerce"
        ).fillna(0)

    return df.dropna(subset=["DATAHORA"])


def dataframe_vazio(df_base):
    return pd.DataFrame(columns=df_base.columns)

# ==================================================
# REGRA 1
# ==================================================
def regra_1(df):
    resultado = (
        df.groupby([
            "Fornecedor",
            "Profissional",
            "DATA",
            "HORA"
        ])
        .filter(lambda x: x["PACIENTE"].nunique() > 1)
        .sort_values([
            "Profissional",
            "DATA",
            "HORA"
        ])
    )

    return resultado.reset_index(drop=True)

# ==================================================
# REGRA 2
# ==================================================
def regra_2(df):
    sessao = df[
        df["Procedimento"]
        .str.upper()
        .str.startswith("SESSAO", na=False)
    ].copy()

    if sessao.empty:
        return dataframe_vazio(df)

    conflitos = []

    for _, grupo in sessao.groupby([
        "Profissional",
        "PACIENTE"
    ]):

        grupo = grupo.sort_values("DATAHORA")

        for i in range(len(grupo) - 1):
            atual = grupo.iloc[i]
            proximo = grupo.iloc[i + 1]

            intervalo = (
                proximo["DATAHORA"] - atual["DATAHORA"]
            ).total_seconds() / 60

            if 0 <= intervalo < 30:
                conflitos.extend([atual, proximo])

    if conflitos:
        return pd.DataFrame(conflitos).drop_duplicates().reset_index(drop=True)

    return dataframe_vazio(df)

# ==================================================
# REGRA 3
# ==================================================
def regra_3(df):
    if "DuracaoHoras" not in df.columns:
        return dataframe_vazio(df)

    plantoes = df[
        df["Procedimento"]
        .str.upper()
        .str.contains("PLANTAO", na=False)
    ].copy()

    if plantoes.empty:
        return dataframe_vazio(df)

    plantoes = plantoes.dropna(subset=["DuracaoHoras"])

    if plantoes.empty:
        return dataframe_vazio(df)

    plantoes["INICIO"] = plantoes["DATAHORA"]

    plantoes["FIM"] = (
        plantoes["INICIO"] +
        pd.to_timedelta(plantoes["DuracaoHoras"], unit="h")
    )

    conflitos = []

    for _, grupo in plantoes.groupby("Profissional"):
        grupo = grupo.sort_values("INICIO")

        for i in range(len(grupo) - 1):
            atual = grupo.iloc[i]
            proximo = grupo.iloc[i + 1]

            if proximo["INICIO"] < atual["FIM"]:
                conflitos.extend([atual, proximo])

    if conflitos:
        return pd.DataFrame(conflitos).drop_duplicates().reset_index(drop=True)

    return dataframe_vazio(df)

# ==================================================
# REGRA 4
# ==================================================
def regra_4(df):
    sessao = df[
        df["Procedimento"]
        .str.upper()
        .str.startswith("SESSAO", na=False)
    ].copy()

    if sessao.empty:
        return dataframe_vazio(df)

    conflitos = []

    for _, grupo in sessao.groupby("PACIENTE"):
        grupo = grupo.sort_values("DATAHORA")

        for i in range(len(grupo) - 1):
            atual = grupo.iloc[i]
            proximo = grupo.iloc[i + 1]

            intervalo = (
                proximo["DATAHORA"] - atual["DATAHORA"]
            ).total_seconds() / 60

            if (
                atual["Profissional"] != proximo["Profissional"] and
                0 <= intervalo < 30
            ):
                conflitos.extend([atual, proximo])

    if conflitos:
        return pd.DataFrame(conflitos).drop_duplicates().reset_index(drop=True)

    return dataframe_vazio(df)

# ==================================================
# REGRA 5
# ==================================================
def regra_5(df):
    resultado = (
        df.groupby([
            "Fornecedor",
            "Profissional",
            "PACIENTE",
            "Procedimento",
            "DATA"
        ])
        .filter(lambda x: len(x) > 1)
        .sort_values([
            "Profissional",
            "PACIENTE",
            "DATA"
        ])
    )

    return resultado.reset_index(drop=True)

# ==================================================
# REGRA 6
# ==================================================
def regra_6(df):
    if "DuracaoHoras" not in df.columns:
        return dataframe_vazio(df)

    plantoes = df[
        df["Procedimento"]
        .str.upper()
        .str.contains("PLANTAO", na=False)
    ].copy()

    outros = df[
        ~df["Procedimento"]
        .str.upper()
        .str.contains("PLANTAO", na=False)
    ].copy()

    if plantoes.empty:
        return dataframe_vazio(df)

    plantoes = plantoes.dropna(subset=["DuracaoHoras"])

    if plantoes.empty:
        return dataframe_vazio(df)

    plantoes["INICIO"] = plantoes["DATAHORA"]

    plantoes["FIM"] = (
        plantoes["INICIO"] +
        pd.to_timedelta(plantoes["DuracaoHoras"], unit="h")
    )

    conflitos = []

    for _, plantao in plantoes.iterrows():
        conflito = outros[
            (outros["Profissional"] == plantao["Profissional"]) &
            (outros["DATAHORA"] >= plantao["INICIO"]) &
            (outros["DATAHORA"] <= plantao["FIM"])
        ]

        if not conflito.empty:
            conflitos.append(conflito)

    if conflitos:
        return pd.concat(conflitos).drop_duplicates().reset_index(drop=True)

    return dataframe_vazio(df)

# ==================================================
# REGRA 7
# ==================================================
def regra_7(df):
    if "DuracaoHoras" not in df.columns:
        return dataframe_vazio(df)

    plantoes = df[
        df["Procedimento"]
        .str.upper()
        .str.contains("PLANTAO", na=False)
    ].copy()

    if plantoes.empty:
        return dataframe_vazio(df)

    conflitos = []

    for _, row in plantoes.iterrows():
        horas = row["DuracaoHoras"]
        quantidade = row.get("Quantidade", 0)

        if pd.isna(horas):
            conflitos.append(row)
            continue

        minimo = 4 if horas >= 4 else int(horas)

        if quantidade < minimo:
            conflitos.append(row)

    if conflitos:
        return pd.DataFrame(conflitos).reset_index(drop=True)

    return dataframe_vazio(df)

# ==================================================
# GERACAO RELATORIO
# ==================================================
def gerar_relatorio_texto(resultados):
    linhas = []

    linhas.append("APONTAMENTOS DE AUDITORIA")
    linhas.append("=" * 50)
    linhas.append("")

    for nome_regra, df_regra in resultados.items():
        linhas.append(nome_regra)
        linhas.append("-" * 50)

        if df_regra.empty:
            linhas.append("Nenhuma ocorrencia encontrada")
            linhas.append("")
            continue

        for _, row in df_regra.iterrows():
            texto = (
                f"{row.get('DATA', '')} | "
                f"{row.get('HORA', '')} | "
                f"Profissional: {row.get('Profissional', '')} | "
                f"Paciente: {row.get('PACIENTE', '')} | "
                f"Procedimento: {row.get('Procedimento', '')}"
            )

            linhas.append(texto)

        linhas.append("")

    return "\n".join(linhas)

# ==================================================
# GERACAO WORD
# ==================================================
def gerar_word(resultados):

    if not DOCX_AVAILABLE:
        return gerar_relatorio_texto(resultados).encode("utf-8")

    doc = Document()

    doc.add_heading(
        "APONTAMENTOS DE AUDITORIA",
        level=1
    )

    for nome_regra, df_regra in resultados.items():
        doc.add_heading(nome_regra, level=2)

        if df_regra.empty:
            doc.add_paragraph(
                "Nenhuma ocorrencia encontrada"
            )
            continue

        for _, row in df_regra.iterrows():
            texto = (
                f"{row.get('DATA', '')} | "
                f"{row.get('HORA', '')} | "
                f"Profissional: {row.get('Profissional', '')} | "
                f"Paciente: {row.get('PACIENTE', '')} | "
                f"Procedimento: {row.get('Procedimento', '')}"
            )

            doc.add_paragraph(texto)

    arquivo = BytesIO()
    doc.save(arquivo)
    arquivo.seek(0)

    return arquivo


# ==================================================
# FORMATACAO DE SAIDA
# ==================================================
def preparar_saida_apontamentos(df):
    df_saida = df.copy()

    colunas_internas = [
        "DATAHORA"
    ]

    df_saida = df_saida.drop(
        columns=[c for c in colunas_internas if c in df_saida.columns],
        errors="ignore"
    )

    for coluna in df_saida.columns:
        if pd.api.types.is_datetime64_any_dtype(df_saida[coluna]):
            if coluna == "DATA":
                df_saida[coluna] = df_saida[coluna].dt.strftime("%d/%m/%Y")
            else:
                df_saida[coluna] = df_saida[coluna].dt.strftime(
                    "%d/%m/%Y %H:%M:%S"
                )

    return df_saida


# ==================================================
# GERACAO EXCEL
# ==================================================
def _nome_aba_excel(nome):
    proibidos = ["\\", "/", "*", "[", "]", ":", "?"]

    for caractere in proibidos:
        nome = nome.replace(caractere, "-")

    return nome[:31]


def gerar_excel(resultados):
    arquivo = BytesIO()

    resumo = pd.DataFrame([
        {
            "Tipo de apontamento": nome_regra,
            "Quantidade": len(df_regra)
        }
        for nome_regra, df_regra in resultados.items()
    ])

    with pd.ExcelWriter(arquivo, engine="openpyxl") as writer:
        resumo.to_excel(
            writer,
            sheet_name="Resumo",
            index=False
        )

        for nome_regra, df_regra in resultados.items():
            nome_aba = _nome_aba_excel(nome_regra)

            if df_regra.empty:
                pd.DataFrame({
                    "Resultado": ["Nenhuma ocorrencia encontrada"]
                }).to_excel(
                    writer,
                    sheet_name=nome_aba,
                    index=False
                )
            else:
                df_saida = preparar_saida_apontamentos(df_regra)

                df_saida.to_excel(
                    writer,
                    sheet_name=nome_aba,
                    index=False
                )

        for sheet in writer.book.worksheets:
            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = sheet.dimensions

            for column_cells in sheet.columns:
                maior = 0
                letra = column_cells[0].column_letter

                for cell in column_cells:
                    valor = "" if cell.value is None else str(cell.value)
                    maior = max(maior, len(valor))

                sheet.column_dimensions[letra].width = min(max(maior + 2, 12), 45)

    arquivo.seek(0)

    return arquivo

# ==================================================
# EXECUCAO AUDITORIA
# ==================================================
def executar_auditoria(df):
    return {
        "Regra 1 - Mesmo horario": regra_1(df),
        "Regra 2 - Sessao < 30 min": regra_2(df),
        "Regra 3 - Sobreposicao plantao": regra_3(df),
        "Regra 4 - Sessao profissionais diferentes": regra_4(df),
        "Regra 5 - Duplicidade": regra_5(df),
        "Regra 6 - Plantonista ocupado": regra_6(df),
        "Regra 7 - Validade plantao": regra_7(df)
    }

# ==================================================
# INTERFACE STREAMLIT
# ==================================================
if STREAMLIT_AVAILABLE:

    if not verificar_acesso_streamlit():
        st.stop()

    uploaded_file = st.file_uploader(
        "Envie o Excel",
        type=["xlsx", "xls", "csv"]
    )

    if uploaded_file:

        try:
            df = carregar_arquivo(uploaded_file)
            df = normalizar_dados(df)

            resultados = executar_auditoria(df)

            st.success("Arquivo processado")

            cols = st.columns(7)

            for i, (nome, resultado) in enumerate(resultados.items()):
                cols[i].metric(f"R{i+1}", len(resultado))

            for nome, resultado in resultados.items():
                with st.expander(nome):
                    if resultado.empty:
                        st.info("Nenhuma ocorrencia")
                    else:
                        st.dataframe(preparar_saida_apontamentos(resultado))

            relatorio = gerar_word(resultados)
            relatorio_excel = gerar_excel(resultados)

            st.download_button(
                label="Baixar Excel com apontamentos",
                data=relatorio_excel,
                file_name="auditoria_apontamentos.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            if DOCX_AVAILABLE:
                st.download_button(
                    label="Baixar Word",
                    data=relatorio,
                    file_name="auditoria.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            else:
                st.download_button(
                    label="Baixar TXT",
                    data=relatorio,
                    file_name="auditoria.txt",
                    mime="text/plain"
                )

        except Exception as e:
            st.error(str(e))

def mostrar_dependencias_ausentes():
    print("=" * 60)
    print("AUDITORIA AUTOMATIZADA")
    print("=" * 60)
    print()

    print("Dependencias ausentes:")

    if not STREAMLIT_AVAILABLE:
        print("- streamlit")

    if not DOCX_AVAILABLE:
        print("- python-docx")

    print()
    print("Instale com:")
    print("pip install streamlit python-docx")
    print()


def imprimir_resumo(resultados):
    print("Resumo dos apontamentos:")
    print("-" * 60)

    total = 0

    for nome, resultado in resultados.items():
        quantidade = len(resultado)
        total += quantidade
        print(f"{nome}: {quantidade}")

    print("-" * 60)
    print(f"Total de linhas apontadas: {total}")


def executar_terminal(caminho_entrada, caminho_saida=None):
    arquivo_entrada = Path(caminho_entrada)

    if not arquivo_entrada.exists():
        raise FileNotFoundError(f"Arquivo nao encontrado: {arquivo_entrada}")

    df = carregar_arquivo(arquivo_entrada)
    df = normalizar_dados(df)
    resultados = executar_auditoria(df)

    imprimir_resumo(resultados)

    if caminho_saida is None:
        caminho_saida = arquivo_entrada.with_name(
            f"{arquivo_entrada.stem}_apontamentos.txt"
        )

    caminho_saida = Path(caminho_saida)
    caminho_saida.write_text(
        gerar_relatorio_texto(resultados),
        encoding="utf-8"
    )

    caminho_excel = caminho_saida.with_suffix(".xlsx")
    relatorio_excel = gerar_excel(resultados)
    caminho_excel.write_bytes(relatorio_excel.getvalue())

    print()
    print(f"Relatorio TXT salvo em: {caminho_saida}")
    print(f"Relatorio Excel salvo em: {caminho_excel}")


def main():
    parser = argparse.ArgumentParser(
        description="Executa auditoria automatizada em arquivo CSV ou Excel."
    )
    parser.add_argument(
        "arquivo",
        nargs="?",
        help="Caminho do arquivo .csv, .xlsx ou .xls"
    )
    parser.add_argument(
        "--saida",
        help="Caminho do relatorio .txt de saida"
    )
    parser.add_argument(
        "--testes",
        action="store_true",
        help="Executa os testes internos"
    )

    args = parser.parse_args()

    if args.testes:
        _teste_regra_1()
        _teste_regra_2()
        _teste_regra_5()
        _teste_sem_conflito()
        print("Testes internos executados com sucesso.")
        return

    if args.arquivo:
        executar_terminal(args.arquivo, args.saida)
        return

    if not STREAMLIT_AVAILABLE:
        mostrar_dependencias_ausentes()
        print("Para usar pelo terminal:")
        print("python app_auditoria.py caminho\\da\\planilha.xlsx")
        print()

# ==================================================
# TESTES
# ==================================================
def _teste_regra_1():
    dados = pd.DataFrame({
        "Fornecedor": ["X", "X"],
        "Profissional": ["ANA", "ANA"],
        "DATA": ["2025-01-01", "2025-01-01"],
        "HORA": ["08:00", "08:00"],
        "PACIENTE": ["JOAO", "MARIA"],
        "Procedimento": ["RX", "RX"]
    })

    dados = normalizar_dados(dados)

    resultado = regra_1(dados)

    assert len(resultado) == 2


def _teste_regra_2():
    dados = pd.DataFrame({
        "Fornecedor": ["X", "X"],
        "Profissional": ["ANA", "ANA"],
        "DATA": ["2025-01-01", "2025-01-01"],
        "HORA": ["08:00", "08:20"],
        "PACIENTE": ["JOAO", "JOAO"],
        "Procedimento": ["SESSAO FISIO", "SESSAO FISIO"]
    })

    dados = normalizar_dados(dados)

    resultado = regra_2(dados)

    assert len(resultado) == 2


def _teste_regra_5():
    dados = pd.DataFrame({
        "Fornecedor": ["X", "X"],
        "Profissional": ["ANA", "ANA"],
        "DATA": ["2025-01-01", "2025-01-01"],
        "HORA": ["08:00", "09:00"],
        "PACIENTE": ["JOAO", "JOAO"],
        "Procedimento": ["RX", "RX"]
    })

    dados = normalizar_dados(dados)

    resultado = regra_5(dados)

    assert len(resultado) == 2


def _teste_sem_conflito():
    dados = pd.DataFrame({
        "Fornecedor": ["X"],
        "Profissional": ["ANA"],
        "DATA": ["2025-01-01"],
        "HORA": ["08:00"],
        "PACIENTE": ["JOAO"],
        "Procedimento": ["RX"]
    })

    dados = normalizar_dados(dados)

    resultado = regra_1(dados)

    assert resultado.empty


if __name__ == "__main__":
    main()
