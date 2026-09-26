import pandas as pd
import re
import unicodedata


# =========================================================
# NORMALIZA TEXTO
# =========================================================

def normalizar_texto(texto):

    if pd.isna(texto):
        return ""

    texto = str(texto).strip().upper()

    if not texto:
        return ""

    texto = unicodedata.normalize(
        "NFKD",
        texto
    )

    texto = "".join(
        caractere
        for caractere in texto
        if not unicodedata.combining(caractere)
    )

    texto = re.sub(
        r"\s+",
        " ",
        texto
    )

    return texto.strip()


# =========================================================
# LIMPA REFERÊNCIA
# =========================================================

def limpar_referencia(referencia):

    referencia = normalizar_texto(
        referencia
    )

    if not referencia:
        return ""

    # Remove espaços desnecessários ao redor
    # de separadores de múltiplas referências.

    referencia = re.sub(
        r"\s*/\s*",
        " / ",
        referencia
    )

    referencia = re.sub(
        r"\s*;\s*",
        " / ",
        referencia
    )

    return referencia.strip()


# =========================================================
# EXTRAI REFERÊNCIA
# =========================================================

def extrair_referencia(texto):
    """
    Extrai a referência diretamente da descrição.

    Exemplos:

        [01002395] OLEO MOTOR
            -> 01002395

        [1080302003 / 0100207] OLEO
            -> 1080302003 / 0100207

        01002395 - OLEO MOTOR
            -> 01002395

        OLEO MOTOR 15W40
            -> OLEO MOTOR 15W40

    Mantém zeros à esquerda e múltiplas referências.
    """

    texto = normalizar_texto(
        texto
    )

    if not texto:
        return ""


    # -------------------------------------------------
    # 1. PROCURA REFERÊNCIA ENTRE [ ]
    # -------------------------------------------------

    encontrados = re.findall(
        r"\[([^\]]+)\]",
        texto
    )

    referencias = []


    for encontrado in encontrados:

        partes = re.split(
            r"[/;,]+",
            encontrado
        )

        for parte in partes:

            parte = limpar_referencia(
                parte
            )

            if (
                parte
                and parte not in referencias
            ):

                referencias.append(
                    parte
                )


    if referencias:

        return " / ".join(
            referencias
        )


    # -------------------------------------------------
    # 2. PROCURA REFERÊNCIA ANTES DE " - "
    # -------------------------------------------------

    if " - " in texto:

        primeira_parte = texto.split(
            " - ",
            1
        )[0].strip()


        partes = re.split(
            r"[/;,]+",
            primeira_parte
        )


        for parte in partes:

            parte = limpar_referencia(
                parte
            )

            if (
                parte
                and parte not in referencias
            ):

                referencias.append(
                    parte
                )


        if referencias:

            return " / ".join(
                referencias
            )


    # -------------------------------------------------
    # 3. NÃO EXISTE REFERÊNCIA SEPARADA
    # -------------------------------------------------

    # Nesse caso mantém a descrição completa.
    # O comparador poderá tentar localizar
    # a descrição quando não houver código explícito.

    return texto


# =========================================================
# LÊ PEDIDO
# =========================================================

def ler_pedido(caminho_arquivo):
    """
    Lê o arquivo Excel do pedido.

    Identifica automaticamente:

        Código
        Descrição
        Marca
        Quantidade

    A referência é extraída da coluna Descrição.
    """

    try:

        df = pd.read_excel(
            caminho_arquivo
        )

    except Exception as erro:

        raise Exception(
            f"Não foi possível ler o arquivo do pedido: {erro}"
        )


    # -------------------------------------------------
    # VERIFICA SE ESTÁ VAZIO
    # -------------------------------------------------

    if df.empty:

        raise Exception(
            "O arquivo do pedido está vazio."
        )


    # -------------------------------------------------
    # IDENTIFICA AS COLUNAS
    # -------------------------------------------------

    colunas = {}


    for coluna in df.columns:

        nome = normalizar_texto(
            coluna
        )


        # Código

        if (
            "CODIGO" in nome
            or "COD" == nome
            or "CÓDIGO" in str(coluna).upper()
        ):

            colunas["codigo"] = coluna


        # Descrição

        elif (
            "DESCRICAO" in nome
            or "PRODUTO" in nome
            or "ITEM" in nome
        ):

            colunas["descricao"] = coluna


        # Marca

        elif "MARCA" in nome:

            colunas["marca"] = coluna


        # Quantidade

        elif (
            "QTD" in nome
            or "QUANT" in nome
            or "QUANTIDADE" in nome
        ):

            colunas["quantidade"] = coluna


    # -------------------------------------------------
    # COLUNAS OBRIGATÓRIAS
    # -------------------------------------------------

    obrigatorias = [
        "codigo",
        "descricao",
        "marca",
        "quantidade"
    ]


    campos_faltantes = [
        campo
        for campo in obrigatorias
        if campo not in colunas
    ]


    if campos_faltantes:

        raise Exception(
            "Não encontrei no pedido as seguintes "
            "colunas: "
            + ", ".join(
                campos_faltantes
            )
        )


    # -------------------------------------------------
    # SELECIONA AS COLUNAS
    # -------------------------------------------------

    pedido = df[
        [
            colunas["codigo"],
            colunas["descricao"],
            colunas["marca"],
            colunas["quantidade"]
        ]
    ].copy()


    # -------------------------------------------------
    # PADRONIZA NOMES
    # -------------------------------------------------

    pedido.columns = [
        "Código",
        "Descrição",
        "Marca",
        "Qtd"
    ]


    # -------------------------------------------------
    # LIMPA CAMPOS
    # -------------------------------------------------

    pedido["Descrição"] = (
        pedido["Descrição"]
        .fillna("")
        .astype(str)
        .str.strip()
    )


    pedido["Marca"] = (
        pedido["Marca"]
        .fillna("")
        .astype(str)
        .str.strip()
    )


    # -------------------------------------------------
    # EXTRAI A REFERÊNCIA
    # -------------------------------------------------

    pedido["Referência"] = (
        pedido["Descrição"]
        .apply(
            extrair_referencia
        )
    )


    # -------------------------------------------------
    # ORGANIZA AS COLUNAS
    # -------------------------------------------------

    pedido = pedido[
        [
            "Código",
            "Referência",
            "Descrição",
            "Marca",
            "Qtd"
        ]
    ]


    # -------------------------------------------------
    # SUBSTITUI NULOS
    # -------------------------------------------------

    pedido = pedido.fillna("")


    # -------------------------------------------------
    # RETORNA
    # -------------------------------------------------

    return pedido
