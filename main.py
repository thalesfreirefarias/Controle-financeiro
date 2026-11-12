from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

# Nome do arquivo
arquivo = Path("controle_financeiro.xlsx")

# Cores do projeto
AZUL = "17365D"
BRANCO = "FFFFFF"

# Verificar se o arquivo já existe
if arquivo.exists():
    wb = load_workbook(arquivo)
    print("Arquivo existente carregado!")

else:
    # Criar arquivo Excel
    wb = Workbook()

    # Criar e nomear as abas
    dashboard = wb.active
    dashboard.title = "Dashboard"

    lancamentos = wb.create_sheet("Lançamentos")
    categorias = wb.create_sheet("Categorias")
    instrucoes = wb.create_sheet("Como Usar")

    # Títulos das abas
    titulos = {
        "Dashboard": "CONTROLE FINANCEIRO ESSENCIAL",
        "Lançamentos": "REGISTRO DE RECEITAS E DESPESAS",
        "Categorias": "CATEGORIAS FINANCEIRAS",
        "Como Usar": "MANUAL DE UTILIZAÇÃO"
    }

    # Formatar títulos
    for nome, titulo in titulos.items():
        ws = wb[nome]

        ws["A1"] = titulo
        ws["A1"].font = Font(
            name="Aptos",
            size=16,
            bold=True,
            color=BRANCO
        )

        ws["A1"].fill = PatternFill(
            start_color=AZUL,
            end_color=AZUL,
            fill_type="solid"
        )

        ws["A1"].alignment = Alignment(
            horizontal="center"
        )

        ws.merge_cells("A1:F2")
        ws.column_dimensions["A"].width = 25

    # Cabeçalhos dos lançamentos
    cabecalhos = ["Data", "Descrição", "Categoria", "Tipo", "Valor"]

    for coluna, titulo in enumerate(cabecalhos, start=1):
        lancamentos.cell(row=4, column=coluna, value=titulo)

    # Dados de exemplo
    dados = [
        ["12/11/2026", "Salário", "Trabalho", "Receita", 3500],
        ["12/11/2026", "Mercado", "Alimentação", "Despesa", 250],
        ["12/11/2026", "Academia", "Saúde", "Despesa", 120]
    ]

    for linha in dados:
        lancamentos.append(linha)

    print("Estrutura inicial criada!")

# Salvar arquivo
wb.save(arquivo)

print("Controle financeiro salvo com sucesso!")