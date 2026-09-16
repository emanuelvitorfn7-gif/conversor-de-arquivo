# Conversor de Arquivos

Aplicação desktop desenvolvida em Python para converter arquivos entre diferentes formatos de dados de maneira simples e rápida.


## Funcionalidades

* Conversão entre Excel, CSV, JSON e TSV
* Seleção do arquivo por uma interface gráfica
* Escolha do formato de saída
* Definição do local onde o arquivo convertido será salvo
* Interface escura, limpa e fácil de utilizar
* Mensagens de confirmação e tratamento de erros

## Formatos suportados

| Formato | Extensão |
| ------- | -------- |
| Excel   | `.xlsx`  |
| CSV     | `.csv`   |
| JSON    | `.json`  |
| TSV     | `.tsv`   |

## Tecnologias utilizadas

* Python
* pandas
* openpyxl
* tkinter

> O `tkinter` normalmente já acompanha a instalação do Python no Windows.

## Requisitos

Antes de começar, instale:

* Python 3.10 ou superior
* Visual Studio Code
* Extensão Python para o VS Code

## Como executar no Windows

Clone o repositório:

```bash
git clone https://github.com/SEU-USUARIO/conversor-de-arquivo.git
```

Entre na pasta:

```bash
cd conversor-de-arquivo
```

Crie o ambiente virtual:

```bash
python -m venv .venv
```

Ative o ambiente no PowerShell:

```powershell
.venv\Scripts\Activate.ps1
```

Instale as dependências:

```bash
python -m pip install -r requirements.txt
```

Execute a aplicação:

```bash
python app.py
```

## Como executar no Linux ou macOS

Crie e ative o ambiente virtual:

```bash
python3 -m venv .venv
source .venv/bin/activate
```

Instale as dependências:

```bash
python3 -m pip install -r requirements.txt
```

Execute a aplicação:

```bash
python3 app.py
```

## Como usar

1. Clique em **Selecionar arquivo**.
2. Escolha o arquivo que deseja converter.
3. Selecione o formato de saída.
4. Clique em **Converter arquivo**.
5. Escolha a pasta e o nome do novo arquivo.
6. Aguarde a mensagem de confirmação.

## Estrutura do projeto

```text
conversor-de-arquivo/
├── app.py
├── requirements.txt
├── README.md
└── .gitignore
```

## Dependências

As principais dependências estão disponíveis no arquivo `requirements.txt`:

```text
pandas
openpyxl
```

## Objetivo do projeto

Este projeto foi criado para praticar desenvolvimento em Python, manipulação de dados com pandas, conversão de arquivos e criação de interfaces gráficas com tkinter.

## Autor

Desenvolvido por **Emanuel Vítor Fernandes Nascimento**.

[GitHub](https://github.com/emanuelvitorfn7-gif)

