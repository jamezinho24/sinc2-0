import pandas as pd
import re
import unicodedata


def normalizar_texto(texto):
    if pd.isna(texto):
        return ""

    texto = str(texto).strip().upper()

    if not texto:
        return ""

    texto = unicodedata.normalize("NFKD", texto)

    texto = "".join(
        caractere
        for caractere in texto
        if not unicodedata.combining(caractere)
    )

    texto = re.sub(r"\s+", " ", texto)

    return texto.strip()


def converter_preco(valor):
    if pd.isna(valor):
        return 0.0

    if isinstance(valor, (int, float)):
        return float(valor)

    texto = str(valor).strip()

    if texto == "":
        return 0.0

    texto = texto.replace("R$", "").replace(" ", "").strip()
    texto = re.sub(r"[^0-9,.\-]", "", texto)

    if texto == "":
        return 0.0

    if "," in texto and "." in texto:
        texto = texto.replace(".", "").replace(",", ".")
    elif "," in texto:
        texto = texto.replace(",", ".")

    try:
        return float(texto)
    except (ValueError, TypeError):
        return 0.0


def localizar_coluna(colunas, palavras):
    for coluna in colunas:
        nome = normalizar_texto(coluna)

        for palavra in palavras:
            palavra = normalizar_texto(palavra)

            if palavra in nome:
                return coluna

    return None


def encontrar_cabecalho(caminho_arquivo):
    amostra = pd.read_excel(
        caminho_arquivo,
        header=None,
        nrows=15
    )

    melhor_linha = 0
    melhor_pontuacao = -1

    for indice, linha in amostra.iterrows():
        valores = [
            normalizar_texto(valor)
            for valor in linha.tolist()
        ]

        pontuacao = 0

        for valor in valores:
            if not valor:
                continue

            if (
                "DESCRICAO" in valor
                or "PRODUTO" == valor
                or "ITEM" == valor
            ):
                pontuacao += 4

            if (
                "VALOR" in valor
                or "PRECO" in valor
                or "VLR" in valor
            ):
                pontuacao += 3

            if (
                "MARCA" in valor
                or "FABRICANTE" in valor
            ):
                pontuacao += 2

        if pontuacao > melhor_pontuacao:
            melhor_pontuacao = pontuacao
            melhor_linha = indice

    return melhor_linha


def ler_fornecedor(caminho_arquivo):
    try:
        linha_cabecalho = encontrar_cabecalho(caminho_arquivo)

        df = pd.read_excel(
            caminho_arquivo,
            header=linha_cabecalho
        )

        df = df.dropna(
            axis=1,
            how="all"
        )

        if df.empty:
            raise Exception(
                "A cotação do fornecedor está vazia."
            )

        descricao = localizar_coluna(
            df.columns,
            [
                "descricao",
                "descrição",
                "descricao do produto",
                "descrição do produto",
                "produto",
                "item"
            ]
        )

        preco = localizar_coluna(
            df.columns,
            [
                "valor unitario",
                "valor unitário",
                "preco unitario",
                "preço unitário",
                "valor",
                "preco",
                "preço",
                "vlr",
                "unit"
            ]
        )

        marca = None

        for coluna in df.columns:
            nome_coluna = normalizar_texto(coluna)

            if nome_coluna == "MARCA":
                marca = coluna
                break

        if marca is None:
            for coluna in df.columns:
                nome_coluna = normalizar_texto(coluna)

                if nome_coluna == "FABRICANTE":
                    marca = coluna
                    break

        if descricao is None:
            raise Exception(
                "Não foi encontrada a coluna de descrição "
                "na cotação do fornecedor. "
                "Colunas encontradas: "
                + ", ".join(
                    str(coluna)
                    for coluna in df.columns
                )
            )

        if preco is None:
            raise Exception(
                "Não foi encontrada a coluna de preço "
                "na cotação do fornecedor. "
                "Colunas encontradas: "
                + ", ".join(
                    str(coluna)
                    for coluna in df.columns
                )
            )

        resultado = pd.DataFrame()

        resultado["Descrição"] = (
            df[descricao]
            .fillna("")
            .astype(str)
            .str.strip()
        )

        resultado["Preço"] = (
            df[preco]
            .apply(converter_preco)
        )

        if marca is None:
            resultado["Marca Cotada"] = ""
        else:
            resultado["Marca Cotada"] = (
                df[marca]
                .fillna("")
                .astype(str)
                .str.strip()
            )

        resultado = resultado[
            resultado["Descrição"].str.strip() != ""
        ].copy()

        resultado = resultado[
            ~resultado["Descrição"]
            .apply(normalizar_texto)
            .isin(
                [
                    "DESCRICAO",
                    "DESCRICAO DO PRODUTO",
                    "PRODUTO",
                    "ITEM"
                ]
            )
        ]

        resultado.reset_index(
            drop=True,
            inplace=True
        )

        return resultado

    except Exception as erro:
        raise Exception(
            f"ERRO AO LER FORNECEDOR: {erro}"
        )
