# Extrarir keywords de multiplos PDFs, criar um DF e exportar para um arquivo .xlsx.

import os              # Disponibiliza funções para criar e remover um diretório, buscando o conteúdo, modifucando e identifcando o diretório. 
import pandas as pd    # Ferramenta flexível de análise/manipulação de dados de código aberto
import glob            # Gera listas de arquivos que correspondem a determinados padrões
import pdfplumber      # Extrai informações de documentos .pdf 

"""
Obtenha palavras-chave de documentos repetitivos e extraia como um dataframe para um arquivo .xlsx 
"""

# definindo as funções usadas em main()
def get_keyword(start, end, text):
    """
    start: deve ser a palavra antes da palavra-chave.
    end: deve ser a palavra que vem depois da palavra-chave.
    text: representa o texto da(s) página(s) que você acabou de extrair.
    """
    for i in range(len(start)):
        try:
            field = ((text.split(start[i]))[1].split(end[i])[0])
            return field
        except:
            continue

def main():

    # crie um dataframe vazio, a partir do qual as palavras-chave de vários arquivos .pdf serão posteriormente anexadas por linhas.
    my_dataframe = pd.DataFrame()

    for files in glob.glob("C:\\Users\\ccece\\Documents\\Projetos AD\\"
                    "Recuperacao_reconstrucao_dados\\arquivos\\*.pdf"):
        with pdfplumber.open(files) as pdf:
            page = pdf.pages[0]
            text = page.extract_text()
            text = " ".join(text.split())

            # use a função get_keyword para obter todas as palavras-chave desejadas de um documento pdf.

            # obtendo NOME_DESTINATARIO #1
            start = ['Caro']
            end = [':']
            NOME_DESTINATARIO = get_keyword(start, end, text)

            # obtendo ENDEREÇO #2
            start = ['endereço']
            end = [', de']
            ENDEREÇO = get_keyword(start, end, text)

            # obtendo CEP #3
            start = ['CEP']
            end = ['Deve']
            CEP = get_keyword(start, end, text)

            # obtendo DATA_ENTREGA #4
            start = ['data']
            end = [', tentamos']
            DATA_ENTREGA = get_keyword(start, end, text)

            # obtendo NUMERO_OC #5
            start = ['referência']
            end = ['no endereço']
            NUMERO_OC = get_keyword(start, end, text)

            # obtendo PRODUTO #6
            start = ['produto']
            end = [', com']
            PRODUTO = get_keyword(start, end, text)

            # crie uma lista com as palavras-chave extraídas do documento atual.
            my_list = [NOME_DESTINATARIO, ENDEREÇO, CEP, DATA_ENTREGA, NUMERO_OC, PRODUTO]

            # append my list as a row in the dataframe.
            my_list = pd.Series(my_list)

            # anexe minha lista como uma linha no dataframe.
            my_dataframe = my_dataframe.append(my_list, ignore_index=True)

            print("Document's keywords have been extracted successfully!")

    # renomeie colunas de dataframe usando dicionários.
    my_dataframe = my_dataframe.rename(columns={0:'NOME_DESTINATARIO',
                                                1:'ENDEREÇO',
                                                2:'CEP',
                                                3:'DATA_ENTREGA',
                                                4:'NUMERO_OC',
                                                5:'PRODUTO'})

    # alterar meu diretório de trabalho atual
    save_path = ("C:\\Users\\ccece\\Documents\\Projetos AD\\"
        "Recuperacao_reconstrucao_dados\\lista-contatos")
    os.chdir(save_path)

    # extraia meu dataframe para um arquivo .xlsx!
    my_dataframe.to_excel('contacts-list.xlsx', sheet_name = 'my dataframe')
    print("")
    print(my_dataframe)

if __name__ == '__main__':
    main()