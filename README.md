# EduQR

Gera bilhetes com QR Code para grupos de WhatsApp de turmas escolares. A saída é um `.docx` paginado, pronto para imprimir e distribuir aos pais.

## Requisitos

- Python 3.11+

## Instalação

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

Veja [ARCHITECTURE.md](ARCHITECTURE.md).
