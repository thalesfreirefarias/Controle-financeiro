from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

# Criar arquivo Excel
wb = Workbook()

# Criar e nomear as abas
dashboard = wb.active
dashboard.title = "Dashboard"

lancamentos = wb.create_sheet("Lançamentos")
categorias = wb.create_sheet("Categorias")
instrucoes = wb.create_sheet("Como Usar")

# Definir cores do projeto
AZUL = "17365D"
BRANCO = "FFFFFF"

# Criar títulos nas abas
titulos = {
    "Dashboard": "CONTROLE FINANCEIRO ESSENCIAL",
    "Lançamentos": "REGISTRO DE RECEITAS E DESPESAS",
    "Categorias": "CATEGORIAS FINANCEIRAS",
    "Como Usar": "MANUAL DE UTILIZAÇÃO"
}

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

# Salvar
wb.save("controle_financeiro.xlsx")

print("Estrutura criada com sucesso!")