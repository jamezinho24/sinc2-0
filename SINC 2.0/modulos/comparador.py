import pandas as pd
import re
import os
import unicodedata



# =========================================================
# LIMPA NOME DO FORNECEDOR
# =========================================================

def limpar_nome_fornecedor(nome):

    nome = os.path.basename(
        str(nome)
    )

    nome = re.sub(
        r"\.(xlsx|xls)$",
        "",
        nome,
        flags=re.IGNORECASE
    )

    nome = re.sub(
        r"\(\d+\)",
        "",
        nome
    )

    return nome.strip().upper()


# =========================================================
# NORMALIZA TEXTO
# =========================================================

def normalizar_texto(texto):

    if pd.isna(texto):
        return ""

    texto = str(texto).strip().upper()

    texto = unicodedata.normalize("NFKD", texto)
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
# EXTRAI REFERÊNCIAS
# =========================================================

def extrair_referencias(texto):

    if pd.isna(texto):
        return []

    texto = normalizar_texto(
        texto
    )

    if not texto:
        return []

    referencias = []

    # -----------------------------------------------------
    # REFERÊNCIAS DENTRO DE [ ]
    # -----------------------------------------------------

    encontrados = re.findall(
        r"\[([^\]]*)\]",
        texto
    )

    for encontrado in encontrados:

        partes = re.split(
            r"[/;,]+",
            encontrado
        )

        for parte in partes:

            parte = normalizar_texto(
                parte
            )

            if (
                parte
                and parte not in referencias
            ):

                referencias.append(
                    parte
                )

    # -----------------------------------------------------
    # REFERÊNCIA ANTES DE " - "
    # -----------------------------------------------------

    if not referencias and " - " in texto:

        primeira_parte = texto.split(
            " - ",
            1
        )[0].strip()

        partes = re.split(
            r"[/;,]+",
            primeira_parte
        )

        for parte in partes:

            parte = normalizar_texto(
                parte
            )

            if (
                parte
                and parte not in referencias
            ):

                referencias.append(
                    parte
                )

    return referencias


# =========================================================
# VALIDA REFERÊNCIA
# =========================================================

def referencia_pedido_valida(
    referencia
):

    if pd.isna(referencia):
        return False

    referencia = normalizar_texto(
        referencia
    )

    if referencia == "":
        return False

    if referencia in [
        "[ ]",
        "[]",
        "-",
        "NAN"
    ]:
        return False

    return True


# =========================================================
# COMPARA REFERÊNCIAS
# =========================================================

def referencias_sao_iguais(
    referencia_pedido,
    descricao_fornecedor
):

    if not referencia_pedido_valida(
        referencia_pedido
    ):
        return False

    referencia_pedido = normalizar_texto(
        referencia_pedido
    )

    # -----------------------------------------------------
    # REFERÊNCIAS DO PEDIDO
    # -----------------------------------------------------

    referencias_pedido = []

    partes_pedido = re.split(
        r"[/;,]+",
        referencia_pedido
    )

    for parte in partes_pedido:

        parte = normalizar_texto(
            parte
        )

        if parte:

            referencias_pedido.append(
                parte
            )

    # -----------------------------------------------------
    # REFERÊNCIAS DO FORNECEDOR
    # -----------------------------------------------------

    referencias_fornecedor = (
        extrair_referencias(
            descricao_fornecedor
        )
    )

    # -----------------------------------------------------
    # COMPARAÇÃO EXATA
    # -----------------------------------------------------

    for ref_pedido in referencias_pedido:

        for ref_fornecedor in referencias_fornecedor:

            if ref_pedido == ref_fornecedor:

                return True

    # -----------------------------------------------------
    # PROCURA NO TEXTO COMPLETO
    # -----------------------------------------------------

    descricao = normalizar_texto(
        descricao_fornecedor
    )

    for ref_pedido in referencias_pedido:

        if not ref_pedido:
            continue

        # Evita códigos muito pequenos
        # gerando falsas correspondências.

        if len(ref_pedido) < 4:
            continue

        padrao = (
            r"(?<!\d)"
            + re.escape(ref_pedido)
            + r"(?!\d)"
        )

        if re.search(
            padrao,
            descricao
        ):

            return True

    return False


# =========================================================
# CONVERTE PREÇO
# =========================================================

def converter_preco(valor):

    if pd.isna(valor):
        return None

    if isinstance(
        valor,
        (int, float)
    ):

        try:

            valor = float(
                valor
            )

        except (
            ValueError,
            TypeError
        ):

            return None

        # ZERO NÃO É PREÇO VÁLIDO

        if valor <= 0:
            return None

        return valor

    valor = str(
        valor
    ).strip()

    if valor == "":
        return None

    valor = valor.upper().replace("R$", "")
    valor = valor.replace(" ", "")
    valor = re.sub(r"[^0-9,.-]", "", valor)

    if valor.count("-") > 0 and not valor.startswith("-"):
        return None


    # -----------------------------------------------------
    # FORMATO BRASILEIRO
    #
    # 1.234,56 -> 1234.56
    # 1234,56  -> 1234.56
    # -----------------------------------------------------

    if (
        "," in valor
        and "." in valor
    ):

        valor = valor.replace(
            ".",
            ""
        )

        valor = valor.replace(
            ",",
            "."
        )

    elif "," in valor:

        valor = valor.replace(
            ",",
            "."
        )

    valor = re.sub(
        r"[^0-9.\-]",
        "",
        valor
    )

    if valor == "":
        return None

    try:

        valor = float(
            valor
        )

    except (
        ValueError,
        TypeError
    ):

        return None

    # -----------------------------------------------------
    # PREÇOS ZERO OU NEGATIVOS NÃO PARTICIPAM
    # -----------------------------------------------------

    if valor <= 0:
        return None

    return valor


# =========================================================
# COMPARADOR
# =========================================================

def comparar(
    df_pedido,
    fornecedores
):

    resultado = []
    totais = {}
    economia_total = 0.0

    # =====================================================
    # PREPARA TODOS OS FORNECEDORES
    # =====================================================

    fornecedores_normalizados = []

    for arquivo_fornecedor in fornecedores.keys():

        nome_fornecedor = (
            limpar_nome_fornecedor(
                arquivo_fornecedor
            )
        )

        fornecedores_normalizados.append(
            (
                arquivo_fornecedor,
                nome_fornecedor
            )
        )

        totais[
            nome_fornecedor
        ] = 0.0

    # =====================================================
    # PERCORRE TODOS OS ITENS DO PEDIDO
    # =====================================================

    for _, item in df_pedido.iterrows():

        referencia = normalizar_texto(
            item.get(
                "Referência",
                ""
            )
        )

        linha = {

            "Código": item.get(
                "Código",
                ""
            ),

            "Referência": item.get(
                "Referência",
                ""
            ),

            "Descrição": item.get(
                "Descrição",
                ""
            ),

            "Marca": item.get(
                "Marca",
                ""
            ),

            "Qtd": item.get(
                "Qtd",
                ""
            )

        }

        # =================================================
        # QUANTIDADE
        # =================================================

        try:

            quantidade = item.get(
                "Qtd",
                0
            )

            if isinstance(
                quantidade,
                str
            ):

                quantidade = (
                    quantidade
                    .strip()
                    .replace(
                        ".",
                        ""
                    )
                    .replace(
                        ",",
                        "."
                    )
                )

            quantidade = float(
                quantidade
            )

            if pd.isna(quantidade) or quantidade < 0:
                quantidade = 0.0

        except (
            ValueError,
            TypeError
        ):

            quantidade = 0.0


        # =================================================
        # GUARDA OS PREÇOS VÁLIDOS
        # =================================================

        precos_validos = {}

        # =================================================
        # ANALISA TODOS OS FORNECEDORES
        # =================================================

        for (
            arquivo_fornecedor,
            nome_fornecedor
        ) in fornecedores_normalizados:

            df_fornecedor = fornecedores[
                arquivo_fornecedor
            ]

            precos_encontrados = []

            produtos_encontrados = []

            # ---------------------------------------------
            # PROCURA TODAS AS LINHAS DO FORNECEDOR
            # ---------------------------------------------

            for indice, produto in df_fornecedor.iterrows():

                descricao_fornecedor = produto.get(
                    "Descrição",
                    ""
                )

                if not referencias_sao_iguais(
                    referencia,
                    descricao_fornecedor
                ):

                    continue

                preco = converter_preco(
                    produto.get(
                        "Preço",
                        ""
                    )
                )

                # PREÇO ZERO / VAZIO / NEGATIVO
                # NÃO PARTICIPA DA COMPARAÇÃO

                if preco is None:
                    continue

                precos_encontrados.append(
                    preco
                )

                produtos_encontrados.append(
                    (
                        indice,
                        produto,
                        preco
                    )
                )

            # ---------------------------------------------
            # SE ENCONTROU PREÇOS
            # ---------------------------------------------

            if precos_encontrados:

                menor_preco_fornecedor = min(
                    precos_encontrados
                )

                precos_validos[
                    nome_fornecedor
                ] = menor_preco_fornecedor

                linha[
                    nome_fornecedor
                ] = round(
                    menor_preco_fornecedor,
                    2
                )

                # -----------------------------------------
                # MARCA DO MENOR PREÇO
                # -----------------------------------------

                marca_cotada = ""

                for (
                    indice,
                    produto,
                    preco
                ) in produtos_encontrados:

                    if preco != menor_preco_fornecedor:
                        continue

                    marca = produto.get(
                        "Marca Cotada",
                        ""
                    )

                    if pd.isna(marca):
                        marca = ""

                    else:

                        marca = str(
                            marca
                        ).strip()

                    if (
                        marca
                        and marca.upper() != "NAN"
                    ):

                        marca_cotada = marca

                    break

                linha[
                    f"{nome_fornecedor} Marca Cotada"
                ] = marca_cotada

            else:

                linha[
                    nome_fornecedor
                ] = ""

                linha[
                    f"{nome_fornecedor} Marca Cotada"
                ] = ""

        # =================================================
        # NENHUM FORNECEDOR ENCONTROU PREÇO
        # =================================================

        if not precos_validos:

            linha[
                "Economia"
            ] = 0.0

            resultado.append(
                linha
            )

            continue

        # =================================================
        # ENCONTRA O MENOR PREÇO
        # ENTRE TODOS OS FORNECEDORES
        # =================================================

        menor_preco = min(
            precos_validos.values()
        )

        # =================================================
        # ENCONTRA O FORNECEDOR VENCEDOR
        # =================================================

        fornecedores_menor_preco = [

            fornecedor

            for fornecedor, preco
            in precos_validos.items()

            if preco == menor_preco

        ]

        fornecedor_vencedor = (
            fornecedores_menor_preco[0]
        )

        # =================================================
        # MAIOR PREÇO VÁLIDO
        # =================================================

        maior_preco = max(
            precos_validos.values()
        )

        # =================================================
        # ECONOMIA
        # =================================================

        economia_item = (
            maior_preco
            - menor_preco
        ) * quantidade

        if economia_item < 0:
            economia_item = 0.0

        economia_item = round(
            economia_item,
            2
        )

        economia_total += (
            economia_item
        )

        linha[
            "Economia"
        ] = economia_item

        # =================================================
        # TOTAL DO FORNECEDOR VENCEDOR
        # =================================================

        totais[
            fornecedor_vencedor
        ] += (
            menor_preco
            * quantidade
        )

        # =================================================
        # ADICIONA RESULTADO
        # =================================================

        resultado.append(
            linha
        )

    # =====================================================
    # CRIA DATAFRAME
    # =====================================================

    resultado = pd.DataFrame(
        resultado
    )

    # =====================================================
    # COLUNAS FIXAS
    # =====================================================

    colunas_fixas = [

        "Código",
        "Referência",
        "Descrição",
        "Marca",
        "Qtd"

    ]

    # =====================================================
    # COLUNAS DE TODOS OS FORNECEDORES
    # =====================================================

    colunas_fornecedores = []

    for (
        arquivo_fornecedor,
        nome_fornecedor
    ) in fornecedores_normalizados:

        colunas_fornecedores.append(
            nome_fornecedor
        )

        colunas_fornecedores.append(
            f"{nome_fornecedor} Marca Cotada"
        )

    # =====================================================
    # COLUNAS FINAIS
    #
    # MENOR PREÇO E FORNECEDOR FORAM RETIRADOS.
    # =====================================================

    colunas_finais = (

        colunas_fixas

        + colunas_fornecedores

        + [
            "Economia"
        ]

    )

    # =====================================================
    # GARANTE TODAS AS COLUNAS
    # =====================================================

    for coluna in colunas_finais:

        if coluna not in resultado.columns:

            resultado[
                coluna
            ] = ""

    resultado = resultado.reindex(
        columns=colunas_finais
    )

    # =====================================================
    # RETORNO
    # =====================================================

    return {

        "tabela": resultado,

        "totais": totais,

        "economia": round(
            economia_total,
            2
        )

    }
