import logging
from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError
import pandas as pd
from datetime import datetime

# Define o token de acesso do Slack
token_acesso_slack = 'xoxb-6961128526948-6944155784727-C19oWQ8wweuANPIdO0tnfqdE'

# Inicializa o WebClient com o token de acesso do Slack
cliente_slack = WebClient(token=token_acesso_slack)

# Configura o logger
logger = logging.getLogger(__name__)

# Nome do canal que estamos procurando
nome_canal = "sdv-plano"

# Variável para armazenar o ID da conversa (canal)
id_conversa = None

try:
    # Obtém a lista de conversas (canais) usando o WebClient
    for resultado in cliente_slack.conversations_list():
        # Verifica se já encontramos o ID da conversa
        if id_conversa is not None:
            break
        # Percorre todos os canais retornados pela API
        for canal in resultado["channels"]:
            # Verifica se o nome do canal corresponde ao canal que estamos procurando
            if canal["name"] == nome_canal:
                # Se encontrarmos o canal, armazenamos seu ID e saímos do loop
                id_conversa = canal["id"]
                # Imprime o ID do canal encontrado
                print(f"ID da conversa encontrado: {id_conversa}")
                break

except SlackApiError as e:
    # Se ocorrer um erro ao chamar a API do Slack, imprime o erro
    print(f"Erro: {e}")

# Verifica se o ID da conversa foi encontrado
if id_conversa:
    try:
        # Obtém a data atual
        data_atual = datetime.now()
        
        # Converte a data atual para um timestamp Unix
        timestamp_atual = int(data_atual.timestamp())

        # Chama o método conversations.history para obter as últimas 5 mensagens no canal
        resultado = cliente_slack.conversations_history(
            channel=id_conversa,  # ID do canal
            oldest=timestamp_atual - (24 * 60 * 60),  # Definindo o timestamp do dia atual menos 1 dia
            inclusive=True,  # Inclui mensagens do timestamp especificado
            limit=10  # Limita o número de mensagens retornadas
        )
        
        # Inicializa uma lista para armazenar os dados de todas as mensagens
        todos_dados_mensagens = []

        # Itera sobre as mensagens retornadas
        for mensagem in resultado["messages"]:
            # Extrai os valores relevantes da mensagem
            id_mensagem = mensagem["ts"]
            texto_mensagem = mensagem["text"]
            indice_work = texto_mensagem.find("Work:")
            indice_descricao = texto_mensagem.find("Descrição da Atividade:")
            indice_inicio = texto_mensagem.find("Início:")
            indice_termino = texto_mensagem.find("Término:")
            indice_impedimento = texto_mensagem.find("Impedimento:")
            indice_responsavel = texto_mensagem.find("Responsável:")

            work = texto_mensagem[indice_work:indice_descricao].strip().replace("Work:", "").strip()
            descricao = texto_mensagem[indice_descricao:indice_inicio].strip().replace("Descrição da Atividade:", "").strip()
            inicio = texto_mensagem[indice_inicio:indice_termino].strip().replace("Início:", "").strip()
            termino = texto_mensagem[indice_termino:indice_impedimento].strip().replace("Término:", "").strip()
            impedimento = texto_mensagem[indice_impedimento:indice_responsavel].strip().replace("Impedimento:", "").strip()
            responsavel = texto_mensagem[indice_responsavel:].strip().replace("Responsável:", "").strip()

            # Adiciona os valores extraídos à lista
            todos_dados_mensagens.append({
                "Work": work,
                "Descrição da Atividade": descricao,
                "Início": inicio,
                "Término": termino,
                "Impedimento": impedimento,
                "Responsável": responsavel
            })

            # print(todos_dados_mensagens)

        # Cria um DataFrame a partir da lista de dados de mensagens
        df = pd.DataFrame(todos_dados_mensagens)

        # Salva o DataFrame como um arquivo Excel
        nome_arquivo_excel = 'dados.xlsx'
        df.to_excel(nome_arquivo_excel, index=False)

        print("Arquivo Excel '{}' gerado com sucesso.".format(nome_arquivo_excel))

    except SlackApiError as e:
        print("Ocorreu um erro ao obter as mensagens do Slack: {}".format(e.response["error"]))
else:
    # Se o ID da conversa não foi encontrado, imprime uma mensagem indicando isso
    print("Canal não encontrado.")