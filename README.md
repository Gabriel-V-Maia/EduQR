<p align="center">
  <img src="assets/icon.png" width="120" />
</p>

<h1 align="center">EduQR</h1>

<p align="center">
  Gera bilhetes com QR Code para turmas escolares 📚
</p>

<p align="center">
  <img src="https://img.shields.io/badge/python-3.11+-blue">
  <img src="https://img.shields.io/badge/status-active-success">
  <img src="https://img.shields.io/github/v/release/Gabriel-V-Maia/EduQR">
  <img src="https://github.com/Gabriel-V-Maia/EduQR/actions/workflows/python-app.yml/badge.svg")
</p>


Os bilhetes são salvos como um `.docx` paginado, pronto para imprimir e distribuir aos pais ou alunos.

## Requisitos

- Python 3.11+

## Instalação

Você pode baixar o aplicativo direto na páginas de [Releases](https://github.com/Gabriel-V-Maia/EduQR/releases)

Caso prefira rodar o programa com python:

```bash
pip install -r requirements.txt
python main.py
```

## Como usar

1. Cole os links das turmas no campo da esquerda — código da turma na linha de cima, link do WhatsApp na linha de baixo
2. Ajuste a quantidade de bilhetes por turma
3. Escolha o layout de página
4. Clique em **Gerar DOCX**

Opcionalmente: adicione um logo ao QR Code, edite o texto do bilhete em **Editar texto**, salve a sessão para reutilizar depois.

## Build

```bash
pip install pyinstaller
pyinstaller --name EduQR --windowed --onefile main.py
# saída: dist/EduQR.exe
```

## Arquitetura

Veja aqui: [Architecture](ARCHITECTURE.md)
