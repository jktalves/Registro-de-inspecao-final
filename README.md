# Registro de Inspeção

Sistema desktop desenvolvido em Python para registro e controle de inspeções de qualidade. Preenche automaticamente um documento Word com os dados digitados, organizando as informações em três tabelas conforme o padrão interno da Qualimol.

---

## Funcionalidades

- Cadastro de até 14 itens de inspeção
- Preenchimento automático do documento Word (`Planilha.doc`)
- **Tabela 1** — dados gerais do item (lote, nota fiscal, datas, valores medidos)
- **Tabela 2** — características específicas por item selecionado
- **Tabela 3** — grade F3 com 8 posições de medição, com validação de tolerância automática
- Valores fora da tolerância F3 (0,45 – 0,55) ficam destacados em **vermelho** na tela e no documento
- Executável `.exe` para distribuição sem necessidade de instalar Python

---

## Tecnologias utilizadas

- Python 3.14
- Tkinter — interface gráfica
- win32com.client (pywin32) — automação do Microsoft Word
- PyInstaller — geração do executável

---

## Requisitos

| Requisito | Necessário? |
|---|---|
| Microsoft Word instalado | **Sim — obrigatório** |
| Python instalado | Não (já embutido no `.exe`) |
| Dependências extras | Nenhuma |

> O aplicativo controla o Word via automação COM. Sem o Word instalado, o sistema abre mas não consegue salvar os dados.

---

## Como executar (desenvolvimento)

**1. Clone o repositório:**
```bash
git clone https://github.com/jktalves/Registro-de-inspecao-final.git
cd Registro-de-inspecao-final
```

**2. Instale a dependência:**
```bash
pip install pywin32
```

**3. Execute:**
```bash
python app_planilha.py
```

> O arquivo `Planilha.doc` deve estar na mesma pasta que o `app_planilha.py`.

---

## Como gerar o executável

```bash
pip install pyinstaller

pyinstaller --onefile --windowed --name "Registro de Inspeção" --icon ArquivosApp/QMimage.ico --hidden-import win32com --hidden-import win32com.client --hidden-import win32api --hidden-import win32con --hidden-import pywintypes app_planilha.py
```

O `.exe` será gerado na pasta `dist/`.

---

## Estrutura do projeto

```
📁 Registro-de-inspecao-final/
├── app_planilha.py                          ← código-fonte principal
├── Planilha.doc                             ← documento Word utilizado pelo sistema
├── Manual do Usuario - Registro de Inspecao.pdf  ← manual do usuário
├── ArquivosApp/
│   ├── QMimage.ico                          ← ícone do aplicativo
│   └── Registro de Inspeção.spec            ← configuração do PyInstaller
└── README.md
```

---

## Entrega ao cliente

Para distribuir ao cliente, forneça uma pasta contendo:

```
📁 Registro de Inspeção/
├── Registro de Inspeção.exe
├── Planilha.doc
└── Manual do Usuario - Registro de Inspecao.pdf
```

---

## Tolerâncias F3 (Tabela 3)

| Resultado | Faixa |
|---|---|
| Normal | Entre **0,45** e **0,55** |
| **Vermelho (fora da tolerância)** | Menor que 0,45 ou maior que 0,55 |

---

Desenvolvido para **Qualimol Indústria e Comércio de Molas**.
