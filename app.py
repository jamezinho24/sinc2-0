from flask import Flask, render_template, request, send_file
import os
import tempfile
from werkzeug.utils import secure_filename

from modulos.leitor_pedido import ler_pedido
from modulos.leitor_fornecedor import ler_fornecedor
from modulos.comparador import comparar
from modulos.exportador import exportar_excel


app = Flask(__name__)


# =========================================================
# CONFIGURAÇÕES
# =========================================================

UPLOAD_FOLDER = "uploads"

os.makedirs(
    UPLOAD_FOLDER,
    exist_ok=True
)

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024


# Guarda o último resultado gerado
ultimo_resultado = None


# =========================================================
# EXTENSÕES PERMITIDAS
# =========================================================

EXTENSOES_PERMITIDAS = {
    ".xlsx",
    ".xls"
}


def arquivo_excel_valido(nome):
    """
    Verifica se o arquivo possui extensão Excel permitida.
    """

    if not nome:
        return False

    extensao = os.path.splitext(nome)[1].lower()

    return extensao in EXTENSOES_PERMITIDAS


# =========================================================
# PÁGINA INICIAL
# =========================================================

@app.route("/")
def inicio():

    return render_template(
        "index.html"
    )


# =========================================================
# PÁGINA DE NOVA ANÁLISE
# =========================================================

@app.route("/pedido")
def pedido():

    return render_template(
        "pedido.html"
    )


# =========================================================
# RECEBE OS ARQUIVOS
# =========================================================

@app.route(
    "/upload",
    methods=["POST"]
)
def upload():

    global ultimo_resultado

    try:

        # -------------------------------------------------
        # RECEBE O PEDIDO
        # -------------------------------------------------

        pedido = request.files.get(
            "pedido"
        )

        # -------------------------------------------------
        # RECEBE AS COTAÇÕES
        # -------------------------------------------------

        cotacoes = request.files.getlist(
            "cotacoes"
        )

        # -------------------------------------------------
        # VALIDA O PEDIDO
        # -------------------------------------------------

        if pedido is None:

            return (
                "Nenhum arquivo de pedido foi enviado.",
                400
            )

        if pedido.filename == "":

            return (
                "Nenhum arquivo de pedido foi selecionado.",
                400
            )

        if not arquivo_excel_valido(
            pedido.filename
        ):

            return (
                "O pedido deve ser um arquivo Excel "
                "(.xlsx ou .xls).",
                400
            )

        # -------------------------------------------------
        # VALIDA AS COTAÇÕES
        # -------------------------------------------------

        cotacoes_validas = []

        for arquivo in cotacoes:

            if arquivo is None:
                continue

            if arquivo.filename == "":
                continue

            if not arquivo_excel_valido(
                arquivo.filename
            ):

                return (
                    f"O arquivo '{arquivo.filename}' "
                    "não é um Excel válido. "
                    "Use arquivos .xlsx ou .xls.",
                    400
                )

            cotacoes_validas.append(
                arquivo
            )

        if len(cotacoes_validas) == 0:

            return (
                "Nenhuma cotação de fornecedor foi enviada.",
                400
            )

        # -------------------------------------------------
        # SALVA O PEDIDO
        # -------------------------------------------------

        nome_pedido = secure_filename(
            pedido.filename
        )

        if not nome_pedido:

            return (
                "Nome do arquivo do pedido inválido.",
                400
            )

        caminho_pedido = os.path.join(
            app.config["UPLOAD_FOLDER"],
            nome_pedido
        )

        pedido.save(
            caminho_pedido
        )

        # -------------------------------------------------
        # LÊ AS COTAÇÕES
        # -------------------------------------------------

        fornecedores = {}

        for arquivo in cotacoes_validas:

            nome_original = arquivo.filename

            nome_seguro = secure_filename(
                nome_original
            )

            if not nome_seguro:
                continue

            caminho_fornecedor = os.path.join(
                app.config["UPLOAD_FOLDER"],
                nome_seguro
            )

            arquivo.save(
                caminho_fornecedor
            )

            fornecedores[nome_seguro] = (
                ler_fornecedor(
                    caminho_fornecedor
                )
            )

        if len(fornecedores) == 0:

            return (
                "Não foi possível processar "
                "as cotações dos fornecedores.",
                400
            )

        # -------------------------------------------------
        # LÊ O PEDIDO
        # -------------------------------------------------

        df_pedido = ler_pedido(
            caminho_pedido
        )

        # -------------------------------------------------
        # COMPARA PEDIDO E FORNECEDORES
        # -------------------------------------------------

        resultado = comparar(
            df_pedido,
            fornecedores
        )

        tabela = resultado[
            "tabela"
        ]

        totais = resultado[
            "totais"
        ]

        economia = resultado[
            "economia"
        ]

        # -------------------------------------------------
        # GUARDA O ÚLTIMO RESULTADO
        # -------------------------------------------------

        ultimo_resultado = {

            "tabela": tabela,

            "totais": totais,

            "economia": economia
        }

        # -------------------------------------------------
        # EXIBE O RESULTADO
        # -------------------------------------------------

        return render_template(

            "resultado.html",

            tabelas=tabela.to_dict(
                orient="records"
            ),

            colunas=tabela.columns.tolist(),

            totais=totais,

            economia=economia
        )

    except Exception as erro:

        print(
            f"ERRO DURANTE A COMPARAÇÃO: {erro}"
        )

        return (
            f"""
            <h2>Erro ao processar a comparação</h2>

            <p>{erro}</p>

            <br>

            <a href="/pedido">
                Voltar para Nova Análise
            </a>
            """,
            500
        )


# =========================================================
# EXPORTAÇÃO PARA EXCEL
# =========================================================

@app.route("/exportar")
def exportar():

    global ultimo_resultado

    if ultimo_resultado is None:

        return (
            "Nenhuma comparação foi realizada.",
            400
        )

    arquivo_temporario = None

    try:

        # -------------------------------------------------
        # CRIA ARQUIVO TEMPORÁRIO
        # -------------------------------------------------

        arquivo_temporario = tempfile.NamedTemporaryFile(
            delete=False,
            suffix=".xlsx"
        )

        arquivo_temporario.close()

        # -------------------------------------------------
        # GERA O EXCEL
        # -------------------------------------------------

        exportar_excel(

            ultimo_resultado[
                "tabela"
            ],

            ultimo_resultado[
                "totais"
            ],

            ultimo_resultado[
                "economia"
            ],

            arquivo_temporario.name
        )

        # -------------------------------------------------
        # ENVIA O ARQUIVO PARA DOWNLOAD
        # -------------------------------------------------

        return send_file(

            arquivo_temporario.name,

            as_attachment=True,

            download_name="Comparacao_SINC_2.0.xlsx",

            mimetype=(
                "application/vnd.openxmlformats-"
                "officedocument.spreadsheetml.sheet"
            )
        )

    except Exception as erro:

        return (
            f"""
            <h2>Erro ao exportar o Excel</h2>

            <p>{erro}</p>

            <br>

            <a href="/pedido">
                Voltar
            </a>
            """,
            500
        )


# =========================================================
# TRATAMENTO DE ARQUIVO GRANDE
# =========================================================

@app.errorhandler(413)
def arquivo_muito_grande(erro):

    return (
        """
        <h2>Arquivo muito grande</h2>

        <p>
            O tamanho máximo permitido para envio
            é de 50 MB.
        </p>

        <br>

        <a href="/pedido">
            Voltar para Nova Análise
        </a>
        """,
        413
    )


# =========================================================
# EXECUÇÃO
# =========================================================

if __name__ == "__main__":

    app.run(

        debug=True,

        host="0.0.0.0",

        port=5000
    )
