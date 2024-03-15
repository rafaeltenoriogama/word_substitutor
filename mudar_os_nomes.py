import mysql.connector
from docx import Document
import copy

def substituir_valores(documento, marcacao, novo_valor):
    for paragrafo in documento.paragraphs:
        if marcacao in paragrafo.text:
            # Substituir marcacao pelo novo_valor
            paragrafo.text = paragrafo.text.replace(marcacao, str(novo_valor))

# Conectar ao banco de dados
conexao = mysql.connector.connect(
    host="localhost",
    user="yuta",
    password="abc123",
    database="projetos"
)

cursor = conexao.cursor()

# Recuperar dados do banco de dados
cursor.execute("SELECT Pessoa, Processo, Numero, Data FROM mala")
dados = cursor.fetchall()

# Carregar o documento existente
documento_existente = Document("documento.docx")

# Loop através dos dados e substituir no documento
for pessoa, processo, numero, data in dados:
    # Clonar o documento existente para manter as alterações feitas nas iterações anteriores
    documento_modificado = copy.deepcopy(documento_existente)

    # Substituir valores no documento clonado
    substituir_valores(documento_modificado, "{PROCESSO}", processo)
    substituir_valores(documento_modificado, "{PESSOA}", pessoa)
    substituir_valores(documento_modificado, "{NUMERO_DO_PROCESSO}", numero)
    substituir_valores(documento_modificado, "{DATA}", data)

    # Salvar o documento modificado para cada iteração
    nome_arquivo = f"{pessoa} da {processo}° vara do trabalho.docx"
    documento_modificado.save(nome_arquivo)

import os

# Criar lista com todos os nomes de arquivo esperados
nomes_arquivos_esperados = []
for pessoa, processo, numero, data in dados:
    nome_arquivo = f"{pessoa} da {processo}° vara do trabalho.docx"
    nomes_arquivos_esperados.append(nome_arquivo)

# Verificar se cada arquivo existe
for nome_arquivo in nomes_arquivos_esperados:
    if not os.path.exists(nome_arquivo):
        print(f"O documento {nome_arquivo} está faltando.")

# Fechar a conexão com o banco de dados
conexao.close()