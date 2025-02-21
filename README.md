# Automação Vale

Este projeto automatiza o processo manual de integração de Notas Fiscais de Serviço (NFS) emitidas contra a Vale no portal que ela disponibiliza. A automação conta com uma interface para interação do usuário, e faz a comunicação com o portal por meio da técnica de raspagem de tela.

Em detalhes, essa automação tem como base a técnica de raspagem de tela. Por meio da biblioteca Selenium, que possibilita a execução da técnica em sites e interfaces web pois se comunica diretamente com o JavaScript da aplicação (linguagem utilizada na programação dessas interfaces), consegue-se fazer o mapeamento e a manipulação de todos os botões, inputs e campos de um formulário web, tornando possível assim o desenvolvimento de um algoritmo que emula as ações humanas para executar a tarefa.

O operador antes precisava separar a NFS, o XML e o RF - documentos próprios da operação que precisam ser enviados anexados no formulário -, coletar dados presentes nesses documentos para efetuar o preenchimento do formulário, e anexar os arquivos. A Automação segue o mesmo roteiro de ação. Ela extrai a NFS e o XML direto do portal da prefeitura de Serra - ES, local onde fica a sede da Bratec - empresa que realiza os serviços e efetua o faturamento das nossas operações com a Vale -. Essa extração também se dá por intermédio da técnica de raspagem de tela com o Selenium; já as RFs vêm de uma pasta inserida pelo operador na interface de interação da automação, além também de uma planilha em Excel que funciona como um repertório de E-mails dos gestores de contratos, que também é um dado que precisa ser passado para o formulário web da Vale. A automação, em termos vulgares, pega esses documentos, relaciona-os, junta-os em uma pasta individual por processo - além de também manter uma pasta por tipo de documento: XML e NFS - e os integra ao portal da vale pela ação emulatória.
<br/>
<br/>
## Tecnologias Utilizadas

- Python;
- Selenium – Raspagem de tela para automação de interações no portal da Vale;
- Webdriver Manager – Gerenciamento automático do ChromeDriver para o Selenium;
- PyAutoGUI – Simulação de teclas e atalhos do teclado;
- Tkinter – Interface gráfica para interação do usuário;
- Pyperclip – Manipulação da área de transferência (copiar e colar);
- pypdf – Leitura e manipulação de arquivos PDF;
- Pandas – Manipulação de planilhas e estruturação de dados;
- xmltodict – Conversão de arquivos XML em dicionários Python;
- Threading – Execução de processos em paralelo para melhorar o desempenho;
- Queue – Comunicação segura entre threads;
- OS & Shutil – Manipulação de arquivos e diretórios;
- Datetime – Trabalhar com datas e horários no código;

<br/>

## Instalação

1. Clone o repositório ou baixe o arquivo ZIP do programa:
```bash
   https://github.com/git-financeiro-eqs/Automacao_Vale.git
```
2. Instale as dependências:
```bash
    pip install -r requirements.txt
```
3. Execute o programa:
```bash
    python gui.py
```
<br/>

## Como Usar<br/>

1. Abra o programa.
2. No primeiro input insira o número da NFS inicial e no input seguinte o número da NFS final.
3. Clique no botão "Pasta RF" e insira o arquivo que contém todas as correspondentes aos processos que serão enviados.
4. Clique no botão "Planilha Gestores" e insira o arquivo Excel correspondente.
5. Aperte o botão "Play" posicionado à direita e isso ativará a execução da automação. Aguarde finalizar todos os envios e encerre a aplicação.
