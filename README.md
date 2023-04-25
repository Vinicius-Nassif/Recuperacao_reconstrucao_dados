# Projeto Recuperação e Reconstrução de Dados

## 1. Introdução 

A automação é o uso de tecnologia para automatizar ações e processos, reduzindo os trabalhos manuais para aumentar a eficiência na comunicação com seus contatos desejados. 

Nesse projeto, o *Python* foi adotado como linguagem de programação para instrumentalizar o serviço de um freelancer responsável que deveria, a partir um repositório de arquivos .pdf, recuperar e reconstruir os dados em uma planilha Excel. A planilha que continha as informações desejadas foi perdida por um infortúnio.

As informações que devem ser identificadas e extraídas de cada .pdf e alocadas em uma planilha do excel são:

- Nome do destinatário;

- Endereço;

- CEP;

- Data da entrega;

- Número da OC; e

- Produto.

  

O trabalho teve como paradigma uma empresa de tecnologia e venda de equipamentos eletrônicos, que perdeu seu banco de dados de sua clientela. 

O intuito dessa atividade irá melhor expor a automação de processos, a recuperação e reestruturação de um banco de dados por meio da divisão de módulos, variáveis, laço de repetição, manipulação de arquivos e MySQL, conforme demonstrado a seguir.




## 2. Módulos

Módulo é um arquivo contendo funções em *Python* para serem usados em outros programas da mesma linguagem. 

Para instrumentalizar essa atividade de maneira organizada, foi realizada em dois módulos de código *Python*:

- pdf_to_excel.py; e
- data_to_mysql.py.

Cada um será responsável por etapas do processo e serão abordados no decorrer desse trabalho. 



## 3. pdf_to_excel.py

### 3.1 Bibliotecas

​		O projeto teve início no primeiro módulo com a importação das bibliotecas e funções utilizadas na elaboração do código em Python:

![1](https://user-images.githubusercontent.com/111388699/234327257-e8103b24-d491-4299-bb4a-65508f91de39.png)

​		Primeiramente, ocorreu a importação da os, que disponibiliza funções para criar e remover um diretório, buscando o conteúdo.

​		Na sequência, a biblioteca *pandas*, tem como como a sigla "pd", com o objetivo de manipulação e análise de dados.

​		A biblioteca *glob* para gerar uma lista de arquivos que correspondem a determinados padrões. Por último, foi importada a biblioteca pdfpumbler, que extrai informações de documentos .pdf.



### 3.2 Código

#### def get_keyword(start, end, text):

O código teve início com a descrição do que será feito em suas funções, ou seja, obter palavras-chave de documentos repetitivos e extrair como um *dataframe* para um arquivo .xlsx.

![2](https://user-images.githubusercontent.com/111388699/234327334-75aa3468-7f02-4d94-92db-bd9ddb1bf701.png)

A primeira função *get_keyword* tem como objetivo encontrar palavras-chave em um texto e retornar o conteúdo entre essas palavras-chave. Ela recebe três parâmetros: *start*, *end* e *text*.

O parâmetro *start* é uma lista de strings que representam a palavra que antecede a palavra-chave. O parâmetro *end* é uma lista de strings que representam a palavra que segue a palavra-chave. Já o parâmetro *text* é o texto onde serão procuradas as palavras-chave.

A função itera sobre o comprimento de *start* usando o loop "for i in range(len(start))". Para cada palavra na lista *start*, a função tenta encontrar a palavra-chave correspondente no texto usando a função *split* do Python.

Se a palavra-chave é encontrada, a função usa a função *split* novamente para dividir o texto com base na palavra *end* correspondente. O conteúdo entre as palavras-chave é armazenado na variável *field* e é retornado pela função usando o comando *return* *field*.

Se a palavra-chave não é encontrada, a função continua a iterar sobre as próximas palavras-chave na lista *start* até que uma palavra-chave seja encontrada e o conteúdo entre as palavras-chave seja retornado.

No caso de algum erro ocorrer durante a execução da função, como por exemplo, um índice fora do intervalo, a exceção é capturada pelo bloco *try*-*except* e a função continua a procurar pelas próximas palavras-chave na lista *start*.



#### def main():

A função "main" é a função principal do script e coordena a extração de palavras-chave de arquivos PDF e o armazenamento dessas palavras-chave em um dataframe do pandas.

![3](https://user-images.githubusercontent.com/111388699/234327433-9b821ea9-791c-4135-ba9b-1a7cd3b4c904.png)
![4](https://user-images.githubusercontent.com/111388699/234327453-68feae65-f2e3-4d4a-bd73-da2835a906ff.png)

Primeiro, a função cria um dataframe vazio usando a função DataFrame() do pandas e armazena esse dataframe na variável my_dataframe.

Em seguida, a função itera sobre cada arquivo PDF no diretório especificado usando o loop for files in glob.glob(). A função pdfplumber.open() do pdfplumber é usada para abrir cada arquivo PDF e o texto na primeira página é extraído usando o método extract_text(). O texto é limpo usando o método split() e join() para remover espaços em branco extras e garantir que as palavras-chave sejam extraídas corretamente.

A função get_keyword() é então usada para extrair as palavras-chave desejadas do texto extraído. Para cada palavra-chave, a função get_keyword() é chamada com as palavras-chave correspondentes para obter o conteúdo entre elas. As palavras-chave e o conteúdo extraído são armazenados em variáveis correspondentes, como "NOME_DESTINATARIO", "ENDEREÇO", "CEP", "DATA_ENTREGA", "NUMERO_OC" e "PRODUTO".

![5](https://user-images.githubusercontent.com/111388699/234327546-a0d2779b-6993-4511-b1ad-0e21312d60bd.png)

Uma lista é criada com as palavras-chave extraídas do documento atual e é armazenada na variável my_list como um objeto do pandas Series(). A lista é adicionada ao dataframe my_dataframe usando o método append() e o parâmetro ignore_index é definido como "True" para garantir que cada lista seja adicionada como uma nova linha no dataframe.

As colunas do dataframe my_dataframe são renomeadas usando o método rename() do pandas e um dicionário é passado como parâmetro para definir os novos nomes das colunas. Em seguida, o diretório de trabalho é alterado usando a função os.chdir() para salvar o arquivo .xlsx em um local específico.

Por fim, o dataframe my_dataframe é exportado para um arquivo .xlsx usando o método to_excel() do pandas. O arquivo é salvo com o nome "contacts-list.xlsx" e a planilha é nomeada como my dataframe. A função imprime o dataframe my_dataframe para a saída padrão e uma mensagem de conclusão é exibida.



#### if __name__ == '__main__':

A linha "if **name** == '**main**':" é uma expressão condicional que verifica se o script está sendo executado como um programa principal ou se está sendo importado como um módulo por outro script.

![6](https://user-images.githubusercontent.com/111388699/234327663-c6d1427b-e312-4ddb-99b0-273d28c58f45.png)

Se o script estiver sendo executado como um programa principal, a função main() será chamada, que é responsável por extrair as palavras-chave de vários arquivos .pdf e armazená-las em um dataframe.

Se o script estiver sendo importado como um módulo, a função main() não será executada, mas ainda assim o módulo e todas as suas funções estarão disponíveis para o script que o importou.



## 4. connection.py

### 4.1 Bibliotecas

​		O segundo módulo teve início com a importação das seguintes bilbiotecas

![7](https://user-images.githubusercontent.com/111388699/234327769-e9237465-9fb3-438f-934d-9f58d18fd483.png)

​		O MySQL é um sistema de gerenciamento de dados estruturados relacionais e como o *Python* não tem acesso nativo à banco de dados MySQL, foi necessário importar a biblioteca *mysql.connector*, crianco um script de conexão.

​		Na sequência, a biblioteca *xlrd*, tem como funcionalidade de recuperar ou fazer leitura de informações constantes em uma planilha. Já a biblioteca *os* é uma ferramenta empregada para obter funcionalidades dependentes do sistema operacional.



### 4.2 Código

#### def __init__(self, host, schema):

​		O código teve início com a denominação de uma classe chamada *Excel_to_Mysql():* e sua primeira função foi a __init__():, como podemos ver a seguir:

![8](https://user-images.githubusercontent.com/111388699/234327849-a9187420-8b9d-4668-a09c-f1e68899b648.png)

​		O método *__init__* foi definido para ter três parâmetros: o *self*, *host e schema*.

​		Logo abaixo, iniciam-se os objetos *self.conn* e *self.cur*  como vazios com o objetivo de receberem seus valores no desenvolvimento do código. As variáveis *self.my_db_user* e *self.my_db_password* receberam, por meio da biblioteca *os* e do método *get()*, variáveis de ambiente para não ocorrer exposição de credenciais particulares no script. 



#### def criar_servidor(self):

​		Essa função é responsável pela criação da tabela no servidor MySQL.

![9](https://user-images.githubusercontent.com/111388699/234327921-4681572d-0272-43da-9271-39e8e907bf4f.png)

​		A variável *self.conn* recebeu o método  da biblioteca mysql.connector para estabelecer a conexão, indicando como parâmetros a hospedagem do *database* com o *self.host*, o usuário com *self.my_db_user*, a senha como *self.my_db_password* e o nome do banco de dados no MySQL como *self.schema*.

​		A variável *self.cur* recebeu a variável *self.conn* com o método *cursor()*, que inicia a interação com o servidor MySQL. 



#### def conexao_mysql(self):

​		Essa função é responsável por estabelecer a conexão com servidor MySQL e por inteirar os dados na tabela já criada.

![10](https://user-images.githubusercontent.com/111388699/234327953-2cfe22c3-c9af-45ac-8a47-361b6d83d4c4.png)

​		A variável *self.conn* recebeu o método  da biblioteca mysql.connector para estabelecer a conexão, indicando como parâmetros a hospedagem do *database* com o *self.host*, o usuário com *self.my_db_user*, a senha como *self.my_db_password* e o nome do banco de dados como *self.schema*.

​		A variável *self.cur* recebeu a variável *self.conn* com o método *cursor()*, que inicia a interação com o servidor MySQL. 

​		A variável *lista* recebeu o método *list()* para criar uma lista de objetos. A variável diretório especifica o diretório da planilha contacts-list.xls. 

​		A variável arquivo recebeu o método *xlrd.open_workbook()* para acessar os dados da planilha no diretório.  A variável *sheet* recebeu a variável e método *arquivo.sheet_by_index()* para contagem do index. Na sequência foi estabelecido os valores das células da planilha com o método *cell_value()*.

​		Para que os dados sejam inseridos forma desejada,  foi introduzido um laço de repetição *for*. Para cada i em alcance de 1 a 51, a lista será preenchida pelo método *append* de todos os valores da planilha. Encerrado aqui o laço de repetição. 

​		Criada a variável *command_to_sql*, recebeu a instrução de que havendo algum dado repetido na chave primária (PK), a primeira inserção permanecerá e, em caso de eventual repetição, sejam as demais desprezadas. Com isso evita-se a duplicidade de dados no *database*. 

​		A variável *self.cur* recebeu o método executemany(command_to_sql, lista) para preparar a operação do *database* (query ou command) e executar os parâmetros sequenciais ou mapeados estabelecidos.

​		Então, a variável *self.conn* recebeu o método *commit* para efetivar as alterações feitas na tabela no servidor MySQL. 

​		Por último, a variável *self.conn* recebeu o método *close()* para desconectar do servidor MySQL. 



#### Execução - if __name__ == '__main__':

![11](https://user-images.githubusercontent.com/111388699/234328016-874a58bf-eee9-4eac-98d3-86458dbfb5a3.png)

​		Novamente, tem o objetivo de orquestrar a execução sequencial de todas as fases do envio das informações da planilha do excel para o MySQL, como um índice representativo e intuitivo, com o chamado de todos os métodos já explicados e trazendo também mensagens de êxito nas conclusões de cada etapa pelo método *print()*. 

​		Foi criada a variável *build* recebendo a classe Excel_to_Mysql, dando como parâmetros o host="localhost" e schema=’vert’. Lembrando que esses parâmetros estavam sendo preparados desde a *def __init__():*.

​		É estabelecida a variável *servidor_vazio* como *True* (verdadeira) e com a condição *if* abaixo indentada tendo essa variável ainda como verdadeira, é iniciada com a orientação *try*, ou seja, que o *Python* realize a tentativa recebendo a variável *build* com o método *criar_servidor()* e exibir uma mensagem de êxito, caso ainda não exista a tabela pretendida. Ao ser finalizada a orientação *try:*, é indicado o *except Exception as e:* para que seja colocado em um método *print()* a exceção que ocorreu no código.

​		Passada essa etapa, é criada a variável tabela_existe como True (verdadeira) e com a condição if abaixo indentada tendo essa variável ainda como verdadeira, é iniciada com a orientação Try, recebendo a variável *build* com o método *conexao_mysql()*, exibindo mensagem de êxito e conclusão de todo o procedimento. Ao ser finalizada a orientação *try:*, é indicado o *except Exception as e:* para que seja colocado em um método *print()* a exceção que ocorreu no código.



## 6. Conclusão

​		Esse projeto foi criado para melhor elucidar a automação de processos na recuperação e reestruturação de um banco de dados, pois no seu desenvolvimento foi necessária a interação de várias bibliotecas do Python, bem como entender conceitos e aplicações da lógica de programação, bem como conceitos de MySQL.

​		Percebeu-se também que para solucionar o problema principal foi necessário dividi-lo em pequenas tarefas para o desenvolvimento. 

​		Portanto, essa é uma forma de usar a coleta e manipulação de dados a seu favor, seja para uso pessoal ou profissional.

Projeto completo no repositório: https://github.com/Vinicius-Nassif/Recuperacao_reconstrucao_dados
