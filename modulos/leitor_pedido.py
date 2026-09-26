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

    if pd.isna(referencia):
        return ""

    referencia = str(referencia).strip()

    referencia = re.sub(
        r"\s+",
        " ",
        referencia
    )

    return referencia


# =========================================================
# EXTRAI REFERÊNCIA DA DESCRIÇÃO
# =========================================================

def extrair_referencia(descricao):

    if pd.isna(descricao):
        return ""

    descricao = str(
        descricao
    ).strip()

    if not descricao:
        return ""

    # -----------------------------------------------------
    # PROCURA REFERÊNCIAS DENTRO DE [ ]
    # -----------------------------------------------------

    encontrados = re.findall(
        r"\[([^\]]+)\]",
        descricao
    )

    if encontrados:

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

                if parte:
                    referencias.append(
                        parte
                    )

        if referencias:
            return " / ".join(
                referencias
            )

    # -----------------------------------------------------
    # REFERÊNCIA ANTES DE " - "
    # -----------------------------------------------------

    if " - " in descricao:

        referencia = descricao.split(
            " - ",
            1
        )[0].strip()

        referencia = limpar_referencia(
            referencia
        )

        if referencia:
            return referencia

    # -----------------------------------------------------
    # SE NÃO ENCONTROU, USA A DESCRIÇÃO NORMALIZADA
    # -----------------------------------------------------

    return normalizar_texto(
        descricao
    )


# =========================================================
# ENCONTRA COLUNA
# =========================================================

def localizar_coluna(
    df,
    nomes
):

    nomes_normalizados = [
        normalizar_texto(nome)
        for nome in nomes
    ]

    for coluna in df.columns:

        coluna_normalizada = normalizar_texto(
            coluna
        )

        if coluna_normalizada in nomes_normalizados:
            return coluna

    return None


# =========================================================
# LÊ O PEDIDO
# =========================================================

def ler_pedido(
    caminho
):

    try:

        df = pd.read_excel(
            caminho
        )

    except Exception as erro:

        raise Exception(
            f"ERRO AO LER O PEDIDO: {erro}"
        )

    if df.empty:

        raise Exception(
            "A planilha do pedido está vazia."
        )

    # =====================================================
    # LOCALIZA AS COLUNAS
    # =====================================================

    coluna_codigo = localizar_coluna(
        df,
        [
            "Código",
            "Codigo",
            "Cód.",
            "Cod.",
            "Código do Produto",
            "Codigo do Produto"
        ]
    )

    coluna_descricao = localizar_coluna(
        df,
        [
            "Descrição",
            "Descricao",
            "Descrição do Produto",
            "Descricao do Produto",
            "Produto",
            "Item"
        ]
    )

    # -----------------------------------------------------
    # NOSSA MARCA
    #
    # A planilha do pedido utiliza exatamente:
    #
    # Marca (Se for Marca Específica)
    #
    # Essa informação será levada para a coluna
    # "Marca" do resultado do SINC.
    # -----------------------------------------------------

    coluna_marca = localizar_coluna(
        df,
        [
            "Marca (Se for Marca Específica)",
            "Marca (Se for Marca Especifica)"
        ]
    )

    coluna_quantidade = localizar_coluna(
        df,
        [
            "Quantidade",
            "Qtd",
            "Qtde",
            "Quant."
        ]
    )

    # =====================================================
    # VALIDA AS COLUNAS
    # =====================================================

    colunas_faltando = []

    if coluna_codigo is None:
        colunas_faltando.append(
            "Código"
        )

    if coluna_descricao is None:
        colunas_faltando.append(
            "Descrição"
        )

    if coluna_marca is None:
        colunas_faltando.append(
            "Marca (Se for Marca Específica)"
        )

    if coluna_quantidade is None:
        colunas_faltando.append(
            "Quantidade"
        )

    if colunas_faltando:

        raise Exception(
            "Não foi possível localizar as seguintes "
            "colunas no pedido: "
            + ", ".join(
                colunas_faltando
            )
        )

    # =====================================================
    # CRIA O DATAFRAME PADRONIZADO
    # =====================================================

    resultado = pd.DataFrame()

    resultado["Código"] = df[
        coluna_codigo
    ]

    resultado["Descrição"] = df[
        coluna_descricao
    ]

    # -----------------------------------------------------
    # AQUI ESTÁ A CORREÇÃO
    #
    # Pega a nossa marca diretamente da coluna:
    #
    # Marca (Se for Marca Específica)
    #
    # e coloca na coluna "Marca".
    # -----------------------------------------------------

    resultado["Marca"] = df[
        coluna_marca
    ]

    resultado["Qtd"] = df[
        coluna_quantidade
    ]

    # =====================================================
    # EXTRAI REFERÊNCIA
    # =====================================================

    resultado["Referência"] = resultado[
        "Descrição"
    ].apply(
        extrair_referencia
    )

    # =====================================================
    # ORGANIZA AS COLUNAS
    # =====================================================

    resultado = resultado[
        [
            "Código",
            "Referência",
            "Descrição",
            "Marca",
            "Qtd"
        ]
    ]

    # =====================================================
    # LIMPA VALORES
    # =====================================================

    resultado["Código"] = resultado[
        "Código"
    ].apply(
        lambda valor:
        ""
        if pd.isna(valor)
        else str(valor).strip()
    )

    resultado["Descrição"] = resultado[
        "Descrição"
    ].apply(
        lambda valor:
        ""
        if pd.isna(valor)
        else str(valor).strip()
    )

    resultado["Marca"] = resultado[
        "Marca"
    ].apply(
        lambda valor:
        ""
        if pd.isna(valor)
        else str(valor).strip()
    )

    # =====================================================
    # REMOVE LINHAS COMPLETAMENTE VAZIAS
    # =====================================================

    resultado = resultado[
        (
            resultado["Código"].astype(str).str.strip() != ""
        )
        |
        (
            resultado["Descrição"].astype(str).str.strip() != ""
        )
        |
        (
            resultado["Marca"].astype(str).str.strip() != ""
        )
    ]

    resultado = resultado.reset_index(
        drop=True
    )

    return resultado