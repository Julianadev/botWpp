import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import os
import time

def botWpp(navegador):
    """
    :param navegador: Emulador do navegador
    """
    # Caminho do arquivo
    caminho_arquivo = "SEU_ARQUIVO"

    if os.path.exists(caminho_arquivo):
        try:
            # Leitura do arquivo, e o nome da planilha
            df_leitura = pd.read_excel(caminho_arquivo, sheet_name='Planilha1')

            # URL do WhatsApp Web
            navegador.get("https://web.whatsapp.com/")

            # Aguarda que o usuário faça o login manualmente no WhatsApp Web
            input("Pressione Enter após fazer o login no WhatsApp Web e garantir que ele está carregado.")

            for linha, row in df_leitura.iterrows():
                nome_contato = row['Nome']
                mensagem = row['Mensagem']

                # Aguarda para garantir que os elementos sejam encontrados
                time.sleep(3)

                try:
                    # Campo que pesquisa o nome do contato que está na planilha
                    campo_pesquisa = navegador.find_element(By.XPATH,
                                                            '//*[@id="side"]/div[1]/div/div[2]/div[2]/div/div[1]/p')
                    campo_pesquisa.clear()
                    campo_pesquisa.send_keys(nome_contato)
                    campo_pesquisa.send_keys(Keys.ENTER)
                    time.sleep(2)

                    # Campo que envia a mensagem ao contato pesquisado
                    campo_mensagem = navegador.find_element(By.XPATH,
                                                            '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p')
                    campo_mensagem.send_keys(mensagem)
                    campo_mensagem.send_keys(Keys.ENTER)

                except Exception as e:
                    print(f"Erro ao enviar mensagem para {nome_contato}: {e}")

        except Exception as e:
            print('Ocorreu um erro ao ler o arquivo ou processar as mensagens: ', e)
    else:
        print('Caminho inválido.')


if __name__ == "__main__":
    navegador = webdriver.Edge()
    botWpp(navegador)
