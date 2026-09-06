# Conversor de Arquivos

Aplicação desktop minimalista feita em Python para converter arquivos entre Excel, CSV, JSON e TSV. A interface escura utiliza uma estética limpa e sofisticada, inspirada em layouts editoriais modernos.

## Formatos suportados

- Excel (`.xlsx`)
- CSV (`.csv`)
- JSON (`.json`)
- TSV (`.tsv`)

## Como executar no VS Code

Abra a pasta do projeto no VS Code e execute no terminal:

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

No Linux ou macOS, ative o ambiente com:

```bash
source .venv/bin/activate
pip install -r requirements.txt
python3 app.py
```

## Como usar

1. Clique em **Selecionar arquivo**.
2. Escolha o formato de saída.
3. Clique em **Converter arquivo**.
4. Escolha onde salvar o resultado.

## Tecnologias

- Python
- pandas
- openpyxl
- tkinter

## Autor

Emanuel Vítor Fernandes Nascimento
