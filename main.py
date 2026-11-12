import os
import sys
import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime, timedelta

import customtkinter as ctk
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.worksheet.page import PageMargins


# ============================================================
# CONFIGURAÇÕES
# ============================================================

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

NOME_SISTEMA = "SAIM - Sistema de Automação de Instruções Militares"
NOME_BANCO = "saim.db"


def obter_pasta_programa():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


CAMINHO_BANCO = os.path.join(obter_pasta_programa(), NOME_BANCO)


# ============================================================
# BANCO DE DADOS
# ============================================================

def conectar_banco():
    conexao = sqlite3.connect(CAMINHO_BANCO)
    conexao.row_factory = sqlite3.Row
    return conexao


def inicializar_banco():
    conexao = conectar_banco()
    cursor = conexao.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS instrucoes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data TEXT NOT NULL,
            tempos INTEGER NOT NULL,
            tema TEXT NOT NULL,
            instrutor TEXT NOT NULL,
            destinatario TEXT DEFAULT 'ESI'
        )
    """)

    cursor.execute("PRAGMA table_info(instrucoes)")
    colunas = [coluna["name"] for coluna in cursor.fetchall()]

    if "destinatario" not in colunas:
        cursor.execute("""
            ALTER TABLE instrucoes
            ADD COLUMN destinatario TEXT DEFAULT 'ESI'
        """)

    conexao.commit()
    conexao.close()


# ============================================================
# FUNÇÕES DE DATA
# ============================================================

def validar_data(texto):
    try:
        return datetime.strptime(texto.strip(), "%d/%m/%Y").date()
    except ValueError:
        return None


def formatar_data(data_obj):
    return data_obj.strftime("%d/%m/%Y")


def obter_inicio_semana(data_obj):
    return data_obj - timedelta(days=data_obj.weekday())


def obter_dias_semana(data_obj):
    segunda = obter_inicio_semana(data_obj)
    return [segunda + timedelta(days=i) for i in range(5)]


def data_para_banco(data_obj):
    return data_obj.strftime("%Y-%m-%d")


def data_do_banco_para_objeto(texto):
    return datetime.strptime(texto, "%Y-%m-%d").date()


# ============================================================
# MENSAGENS E DIÁLOGOS
# ============================================================

def mensagem_info(titulo, mensagem):
    messagebox.showinfo(titulo, mensagem, parent=app)


def mensagem_erro(titulo, mensagem):
    messagebox.showerror(titulo, mensagem, parent=app)


def mensagem_confirmacao(titulo, mensagem):
    return messagebox.askyesno(titulo, mensagem, parent=app)


def abrir_dialogo_salvar(nome_sugerido):
    return filedialog.asksaveasfilename(
        parent=app,
        title="Salvar planilha semanal",
        defaultextension=".xlsx",
        filetypes=[("Planilha Excel", "*.xlsx")],
        initialfile=nome_sugerido
    )


# ============================================================
# CADASTRO DE INSTRUÇÕES
# ============================================================

def salvar_instrucao():
    data_texto = campo_data.get().strip()
    tempos_texto = campo_tempos.get().strip()
    tema = campo_tema.get().strip()
    instrutor = campo_instrutor.get().strip()
    destinatario = campo_destinatario.get().strip()

    data_obj = validar_data(data_texto)

    if not data_obj:
        mensagem_erro("Data inválida", "Informe a data no formato DD/MM/AAAA.")
        campo_data.focus()
        return

    if not tempos_texto.isdigit():
        mensagem_erro("Quantidade inválida", "Selecione a quantidade de tempos.")
        campo_tempos.focus()
        return

    tempos = int(tempos_texto)

    if tempos < 1 or tempos > 4:
        mensagem_erro("Quantidade inválida", "A quantidade de tempos deve ser de 1 a 4.")
        campo_tempos.focus()
        return

    if not tema:
        mensagem_erro("Campo obrigatório", "Informe o tema da instrução.")
        campo_tema.focus()
        return

    if not instrutor:
        mensagem_erro("Campo obrigatório", "Informe o instrutor.")
        campo_instrutor.focus()
        return

    if destinatario not in ["ESI", "EPA", "ESI e EPA"]:
        mensagem_erro("Destinatário inválido", "Selecione ESI, EPA ou ESI e EPA.")
        return

    conexao = conectar_banco()
    cursor = conexao.cursor()

    cursor.execute("""
        INSERT INTO instrucoes (data, tempos, tema, instrutor, destinatario)
        VALUES (?, ?, ?, ?, ?)
    """, (
        data_para_banco(data_obj),
        tempos,
        tema,
        instrutor,
        destinatario
    ))

    conexao.commit()
    conexao.close()

    limpar_campos()
    atualizar_consulta()

    mensagem_info("Sucesso", "Instrução cadastrada com sucesso.")


def limpar_campos():
    campo_data.delete(0, tk.END)
    campo_data.insert(0, datetime.now().strftime("%d/%m/%Y"))

    campo_tempos.set("1")
    campo_tema.delete(0, tk.END)
    campo_instrutor.delete(0, tk.END)
    campo_destinatario.set("ESI")

    campo_data.focus()


# ============================================================
# CONSULTA DE INSTRUÇÕES
# ============================================================

def obter_instrucoes(filtro_data=None):
    conexao = conectar_banco()
    cursor = conexao.cursor()

    if filtro_data:
        data_obj = validar_data(filtro_data)

        if not data_obj:
            conexao.close()
            return None

        cursor.execute("""
            SELECT id, data, tempos, tema, instrutor, destinatario
            FROM instrucoes
            WHERE data = ?
            ORDER BY data, id
        """, (data_para_banco(data_obj),))
    else:
        cursor.execute("""
            SELECT id, data, tempos, tema, instrutor, destinatario
            FROM instrucoes
            ORDER BY data DESC, id DESC
        """)

    registros = cursor.fetchall()
    conexao.close()
    return registros


def atualizar_consulta():
    for item in tabela_consulta.get_children():
        tabela_consulta.delete(item)

    filtro = campo_filtro_data.get().strip()

    if filtro:
        registros = obter_instrucoes(filtro)
        if registros is None:
            mensagem_erro("Data inválida", "Informe a data no formato DD/MM/AAAA.")
            return
    else:
        registros = obter_instrucoes()

    for registro in registros:
        data_obj = data_do_banco_para_objeto(registro["data"])

        tabela_consulta.insert(
            "",
            tk.END,
            values=(
                registro["id"],
                formatar_data(data_obj),
                registro["tempos"],
                registro["tema"],
                registro["instrutor"],
                registro["destinatario"] or "ESI"
            )
        )


def limpar_filtro():
    campo_filtro_data.delete(0, tk.END)
    atualizar_consulta()


def obter_id_selecionado():
    selecionados = tabela_consulta.selection()

    if not selecionados:
        mensagem_erro("Nenhuma instrução selecionada", "Selecione uma instrução na tabela.")
        return None

    valores = tabela_consulta.item(selecionados[0], "values")
    return int(valores[0])


def editar_instrucao():
    instrucao_id = obter_id_selecionado()

    if instrucao_id is None:
        return

    conexao = conectar_banco()
    cursor = conexao.cursor()

    cursor.execute("""
        SELECT id, data, tempos, tema, instrutor, destinatario
        FROM instrucoes
        WHERE id = ?
    """, (instrucao_id,))

    registro = cursor.fetchone()
    conexao.close()

    if not registro:
        mensagem_erro("Erro", "Não foi possível localizar a instrução.")
        return

    janela = ctk.CTkToplevel(app)
    janela.title("Editar instrução")
    janela.geometry("520x500")
    janela.resizable(False, False)
    janela.transient(app)
    janela.grab_set()

    ctk.CTkLabel(
        janela,
        text="EDITAR INSTRUÇÃO",
        font=ctk.CTkFont(size=20, weight="bold")
    ).pack(pady=15)

    frame = ctk.CTkFrame(janela)
    frame.pack(fill="both", expand=True, padx=20, pady=10)

    ctk.CTkLabel(frame, text="Data (DD/MM/AAAA)").pack(anchor="w", padx=15, pady=(10, 0))
    entrada_data = ctk.CTkEntry(frame)
    entrada_data.pack(fill="x", padx=15, pady=5)
    entrada_data.insert(0, formatar_data(data_do_banco_para_objeto(registro["data"])))

    ctk.CTkLabel(frame, text="Quantidade de tempos").pack(anchor="w", padx=15, pady=(5, 0))
    entrada_tempos = ctk.CTkComboBox(
        frame,
        values=["1", "2", "3", "4"],
        state="readonly"
    )
    entrada_tempos.pack(fill="x", padx=15, pady=5)
    entrada_tempos.set(str(registro["tempos"]))

    ctk.CTkLabel(frame, text="Tema da instrução").pack(anchor="w", padx=15, pady=(5, 0))
    entrada_tema = ctk.CTkEntry(frame)
    entrada_tema.pack(fill="x", padx=15, pady=5)
    entrada_tema.insert(0, registro["tema"])

    ctk.CTkLabel(frame, text="Instrutor").pack(anchor="w", padx=15, pady=(5, 0))
    entrada_instrutor = ctk.CTkEntry(frame)
    entrada_instrutor.pack(fill="x", padx=15, pady=5)
    entrada_instrutor.insert(0, registro["instrutor"])

    ctk.CTkLabel(frame, text="Destinatário").pack(anchor="w", padx=15, pady=(5, 0))
    entrada_destinatario = ctk.CTkComboBox(
        frame,
        values=["ESI", "EPA", "ESI e EPA"],
        state="readonly"
    )
    entrada_destinatario.pack(fill="x", padx=15, pady=5)
    entrada_destinatario.set(registro["destinatario"] or "ESI")

    def salvar_edicao():
        data_obj = validar_data(entrada_data.get())

        if not data_obj:
            messagebox.showerror("Data inválida", "Informe a data no formato DD/MM/AAAA.", parent=janela)
            return

        tempos_texto = entrada_tempos.get().strip()

        if tempos_texto not in ["1", "2", "3", "4"]:
            messagebox.showerror("Quantidade inválida", "Selecione de 1 a 4 tempos.", parent=janela)
            return

        tema = entrada_tema.get().strip()
        instrutor = entrada_instrutor.get().strip()
        destinatario = entrada_destinatario.get().strip()

        if not tema or not instrutor:
            messagebox.showerror("Campos obrigatórios", "Preencha tema e instrutor.", parent=janela)
            return

        if destinatario not in ["ESI", "EPA", "ESI e EPA"]:
            messagebox.showerror("Destinatário inválido", "Selecione ESI, EPA ou ESI e EPA.", parent=janela)
            return

        conexao_edicao = conectar_banco()
        cursor_edicao = conexao_edicao.cursor()

        cursor_edicao.execute("""
            UPDATE instrucoes
            SET data = ?, tempos = ?, tema = ?, instrutor = ?, destinatario = ?
            WHERE id = ?
        """, (
            data_para_banco(data_obj),
            int(tempos_texto),
            tema,
            instrutor,
            destinatario,
            instrucao_id
        ))

        conexao_edicao.commit()
        conexao_edicao.close()

        janela.destroy()
        atualizar_consulta()
        mensagem_info("Sucesso", "Instrução atualizada.")

    ctk.CTkButton(
        janela,
        text="Salvar alterações",
        command=salvar_edicao,
        height=40
    ).pack(pady=15)


def excluir_instrucao():
    instrucao_id = obter_id_selecionado()

    if instrucao_id is None:
        return

    if not mensagem_confirmacao(
        "Confirmar exclusão",
        "Tem certeza de que deseja excluir a instrução selecionada?"
    ):
        return

    conexao = conectar_banco()
    cursor = conexao.cursor()

    cursor.execute("DELETE FROM instrucoes WHERE id = ?", (instrucao_id,))

    conexao.commit()
    conexao.close()

    atualizar_consulta()
    mensagem_info("Sucesso", "Instrução excluída.")


# ============================================================
# EXPORTAÇÃO DA PLANILHA SEMANAL
# ============================================================

def gerar_planilha():
    data_texto = campo_data_planilha.get().strip()
    data_obj = validar_data(data_texto)

    if not data_obj:
        mensagem_erro(
            "Data inválida",
            "Informe uma data válida no formato DD/MM/AAAA para selecionar a semana."
        )
        campo_data_planilha.focus()
        return

    dias = obter_dias_semana(data_obj)
    segunda = dias[0]
    sexta = dias[-1]

    conexao = conectar_banco()
    cursor = conexao.cursor()

    cursor.execute("""
        SELECT id, data, tempos, tema, instrutor, destinatario
        FROM instrucoes
        WHERE data >= ? AND data <= ?
        ORDER BY data ASC, id ASC
    """, (
        data_para_banco(segunda),
        data_para_banco(sexta)
    ))

    registros = cursor.fetchall()
    conexao.close()

    instrucoes_por_dia = {data_para_banco(dia): [] for dia in dias}

    for registro in registros:
        chave = registro["data"]
        if chave in instrucoes_por_dia:
            instrucoes_por_dia[chave].append(registro)

    nome_arquivo = (
        f"{segunda.strftime('%d-%m-%Y')}"
        f"_a_{sexta.strftime('%d-%m-%Y')}.xlsx"
    )

    caminho = abrir_dialogo_salvar(nome_arquivo)

    if not caminho:
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Instruções Semanais"

    # Cores
    azul_escuro = "17365D"
    azul_medio = "D9EAF7"
    azul_claro = "EAF3F8"
    cinza_claro = "F2F2F2"
    verde_claro = "E2F0D9"
    amarelo_claro = "FFF2CC"
    branco = "FFFFFF"
    preto = "000000"

    borda_fina = Side(style="thin", color="A6A6A6")
    borda_media = Side(style="medium", color=azul_escuro)

    borda_celula = Border(
        left=borda_fina,
        right=borda_fina,
        top=borda_fina,
        bottom=borda_fina
    )

    alinhamento_centro = Alignment(
        horizontal="center",
        vertical="center",
        wrap_text=True
    )

    alinhamento_esquerda = Alignment(
        horizontal="left",
        vertical="center",
        wrap_text=True
    )

    # Larguras ajustadas:
    # Aumentamos a coluna A para "TEMPOS" caber inteira.
    # Aumentamos a coluna B para os nomes dos dias do resumo caberem.
    larguras = {
        "A": 14,
        "B": 42,
        "C": 28,
        "D": 22
    }

    for coluna, largura in larguras.items():
        ws.column_dimensions[coluna].width = largura

    # Título principal
    ws.merge_cells("A1:D1")
    ws["A1"] = "SISTEMA DE AUTOMAÇÃO DE INSTRUÇÕES MILITARES"
    ws["A1"].font = Font(name="Arial", size=15, bold=True, color=branco)
    ws["A1"].fill = PatternFill("solid", fgColor=azul_escuro)
    ws["A1"].alignment = alinhamento_centro
    ws["A1"].border = Border(bottom=borda_media)
    ws.row_dimensions[1].height = 30

    # Subtítulo
    ws.merge_cells("A2:D2")
    ws["A2"] = (
        f"PROGRAMAÇÃO SEMANAL — "
        f"{formatar_data(segunda)} a {formatar_data(sexta)}"
    )
    ws["A2"].font = Font(name="Arial", size=12, bold=True, color=azul_escuro)
    ws["A2"].fill = PatternFill("solid", fgColor=azul_medio)
    ws["A2"].alignment = alinhamento_centro
    ws.row_dimensions[2].height = 24

    linha = 4
    total_semanal_tempos = 0
    total_semanal_instrucoes = 0

    resumo_dias = []

    nomes_dias = [
        "SEGUNDA-FEIRA",
        "TERÇA-FEIRA",
        "QUARTA-FEIRA",
        "QUINTA-FEIRA",
        "SEXTA-FEIRA"
    ]

    # Seções de cada dia
    for indice, dia in enumerate(dias):
        chave = data_para_banco(dia)
        instrucoes = instrucoes_por_dia[chave]
        total_dia = sum(int(item["tempos"]) for item in instrucoes)
        quantidade_dia = len(instrucoes)

        total_semanal_tempos += total_dia
        total_semanal_instrucoes += quantidade_dia

        resumo_dias.append({
            "dia": nomes_dias[indice],
            "data": dia,
            "instrucoes": quantidade_dia,
            "tempos": total_dia
        })

        # Cabeçalho do dia
        ws.merge_cells(start_row=linha, start_column=1, end_row=linha, end_column=4)
        celula_dia = ws.cell(row=linha, column=1)
        celula_dia.value = f"{nomes_dias[indice]} — {formatar_data(dia)}"
        celula_dia.font = Font(name="Arial", size=11, bold=True, color=branco)
        celula_dia.fill = PatternFill("solid", fgColor=azul_escuro)
        celula_dia.alignment = alinhamento_centro
        celula_dia.border = borda_celula
        ws.row_dimensions[linha].height = 23
        linha += 1

        # Cabeçalho das instruções
        cabecalhos = [
            "TEMPOS",
            "TEMA DA INSTRUÇÃO",
            "INSTRUTOR",
            "DESTINATÁRIO"
        ]

        for coluna, texto in enumerate(cabecalhos, start=1):
            celula = ws.cell(row=linha, column=coluna)
            celula.value = texto
            celula.font = Font(name="Arial", size=10, bold=True, color=azul_escuro)
            celula.fill = PatternFill("solid", fgColor=azul_medio)
            celula.alignment = alinhamento_centro
            celula.border = borda_celula

        ws.row_dimensions[linha].height = 22
        linha += 1

        # Instruções cadastradas
        if instrucoes:
            for item in instrucoes:
                valores = [
                    item["tempos"],
                    item["tema"],
                    item["instrutor"],
                    item["destinatario"] or "ESI"
                ]

                for coluna, valor in enumerate(valores, start=1):
                    celula = ws.cell(row=linha, column=coluna)
                    celula.value = valor
                    celula.font = Font(name="Arial", size=10, color=preto)
                    celula.fill = PatternFill(
                        "solid",
                        fgColor=branco if linha % 2 == 0 else azul_claro
                    )
                    celula.alignment = (
                        alinhamento_centro if coluna in [1, 4]
                        else alinhamento_esquerda
                    )
                    celula.border = borda_celula

                ws.row_dimensions[linha].height = 28
                linha += 1
        else:
            ws.merge_cells(start_row=linha, start_column=1, end_row=linha, end_column=4)
            celula_vazia = ws.cell(row=linha, column=1)
            celula_vazia.value = "Nenhuma instrução cadastrada para este dia."
            celula_vazia.font = Font(name="Arial", size=10, italic=True, color="666666")
            celula_vazia.fill = PatternFill("solid", fgColor=cinza_claro)
            celula_vazia.alignment = alinhamento_centro
            celula_vazia.border = borda_celula
            ws.row_dimensions[linha].height = 24
            linha += 1

        # Total do dia
        ws.merge_cells(start_row=linha, start_column=1, end_row=linha, end_column=3)

        celula_total_texto = ws.cell(row=linha, column=1)
        celula_total_texto.value = f"TOTAL DE TEMPOS — {nomes_dias[indice]}"
        celula_total_texto.font = Font(name="Arial", size=10, bold=True, color=azul_escuro)
        celula_total_texto.fill = PatternFill("solid", fgColor=verde_claro)
        celula_total_texto.alignment = Alignment(horizontal="right", vertical="center")
        celula_total_texto.border = borda_celula

        celula_total_valor = ws.cell(row=linha, column=4)
        celula_total_valor.value = total_dia
        celula_total_valor.font = Font(name="Arial", size=11, bold=True, color=azul_escuro)
        celula_total_valor.fill = PatternFill("solid", fgColor=verde_claro)
        celula_total_valor.alignment = alinhamento_centro
        celula_total_valor.border = borda_celula

        ws.row_dimensions[linha].height = 23
        linha += 2

    # Resumo semanal
    ws.merge_cells(start_row=linha, start_column=1, end_row=linha, end_column=4)
    celula_resumo = ws.cell(row=linha, column=1)
    celula_resumo.value = "RESUMO SEMANAL"
    celula_resumo.font = Font(name="Arial", size=12, bold=True, color=branco)
    celula_resumo.fill = PatternFill("solid", fgColor=azul_escuro)
    celula_resumo.alignment = alinhamento_centro
    celula_resumo.border = borda_celula
    ws.row_dimensions[linha].height = 25
    linha += 1

    cabecalhos_resumo = [
        "DIA",
        "DATA",
        "Nº DE INSTRUÇÕES",
        "TOTAL DE TEMPOS"
    ]

    for coluna, texto in enumerate(cabecalhos_resumo, start=1):
        celula = ws.cell(row=linha, column=coluna)
        celula.value = texto
        celula.font = Font(name="Arial", size=10, bold=True, color=azul_escuro)
        celula.fill = PatternFill("solid", fgColor=azul_medio)
        celula.alignment = alinhamento_centro
        celula.border = borda_celula

    ws.row_dimensions[linha].height = 30
    linha += 1

    # Dias do resumo em negrito e sem quebra de palavras
    for resumo in resumo_dias:
        valores = [
            resumo["dia"],
            formatar_data(resumo["data"]),
            resumo["instrucoes"],
            resumo["tempos"]
        ]

        for coluna, valor in enumerate(valores, start=1):
            celula = ws.cell(row=linha, column=coluna)
            celula.value = valor
            celula.font = Font(
                name="Arial",
                size=10,
                bold=(coluna == 1),
                color=preto
            )
            celula.fill = PatternFill(
                "solid",
                fgColor=branco if linha % 2 == 0 else azul_claro
            )

            if coluna == 1:
                # Evita que "SEGUNDA-FEIRA" e os demais dias sejam divididos.
                celula.alignment = Alignment(
                    horizontal="left",
                    vertical="center",
                    wrap_text=False,
                    shrink_to_fit=True
                )
            else:
                celula.alignment = alinhamento_centro

            celula.border = borda_celula

        ws.row_dimensions[linha].height = 22
        linha += 1

    # Totais gerais
    linha += 1

    ws.merge_cells(start_row=linha, start_column=1, end_row=linha, end_column=2)

    celula_total_geral = ws.cell(row=linha, column=1)
    celula_total_geral.value = "TOTAL GERAL DA SEMANA"
    celula_total_geral.font = Font(name="Arial", size=11, bold=True, color=azul_escuro)
    celula_total_geral.fill = PatternFill("solid", fgColor=amarelo_claro)
    celula_total_geral.alignment = Alignment(horizontal="right", vertical="center")
    celula_total_geral.border = borda_celula

    celula_qtd = ws.cell(row=linha, column=3)
    celula_qtd.value = total_semanal_instrucoes
    celula_qtd.font = Font(name="Arial", size=11, bold=True, color=azul_escuro)
    celula_qtd.fill = PatternFill("solid", fgColor=amarelo_claro)
    celula_qtd.alignment = alinhamento_centro
    celula_qtd.border = borda_celula

    celula_tempos = ws.cell(row=linha, column=4)
    celula_tempos.value = total_semanal_tempos
    celula_tempos.font = Font(name="Arial", size=11, bold=True, color=azul_escuro)
    celula_tempos.fill = PatternFill("solid", fgColor=amarelo_claro)
    celula_tempos.alignment = alinhamento_centro
    celula_tempos.border = borda_celula

    ws.row_dimensions[linha].height = 26

    # Configuração de impressão
    ws.freeze_panes = "A4"
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    ws.page_margins = PageMargins(
        left=0.25,
        right=0.25,
        top=0.5,
        bottom=0.5,
        header=0.2,
        footer=0.2
    )

    ws.print_title_rows = "1:2"
    ws.sheet_properties.outlinePr.summaryBelow = True

    try:
        wb.save(caminho)
    except PermissionError:
        mensagem_erro(
            "Arquivo em uso",
            "Não foi possível salvar. Feche a planilha no Excel e tente novamente."
        )
        return
    except Exception as erro:
        mensagem_erro("Erro ao salvar", f"Não foi possível salvar a planilha.\n\n{erro}")
        return

    mensagem_info(
        "Planilha gerada",
        f"Planilha semanal salva com sucesso em:\n\n{caminho}"
    )


# ============================================================
# INTERFACE PRINCIPAL
# ============================================================

app = ctk.CTk()
app.title(NOME_SISTEMA)
app.geometry("1200x850")
app.minsize(950, 650)


def alternar_tela_cheia(event=None):
    app.attributes("-fullscreen", not app.attributes("-fullscreen"))


def sair_tela_cheia(event=None):
    app.attributes("-fullscreen", False)


app.bind("<F11>", alternar_tela_cheia)
app.bind("<Escape>", sair_tela_cheia)


# Container com rolagem
container_principal = ctk.CTkScrollableFrame(app)
container_principal.pack(fill="both", expand=True, padx=10, pady=10)


# Cabeçalho
titulo = ctk.CTkLabel(
    container_principal,
    text="SAIM",
    font=ctk.CTkFont(size=30, weight="bold")
)
titulo.pack(pady=(10, 0))

subtitulo = ctk.CTkLabel(
    container_principal,
    text="Sistema de Automação de Instruções Militares",
    font=ctk.CTkFont(size=15)
)
subtitulo.pack(pady=(0, 15))


# ============================================================
# CADASTRO
# ============================================================

frame_cadastro = ctk.CTkFrame(container_principal)
frame_cadastro.pack(fill="x", padx=10, pady=10)

ctk.CTkLabel(
    frame_cadastro,
    text="CADASTRAR INSTRUÇÃO",
    font=ctk.CTkFont(size=18, weight="bold")
).pack(pady=12)

formulario = ctk.CTkFrame(frame_cadastro)
formulario.pack(fill="x", padx=15, pady=10)

formulario.grid_columnconfigure(1, weight=1)
formulario.grid_columnconfigure(3, weight=1)

ctk.CTkLabel(formulario, text="Data (DD/MM/AAAA):").grid(
    row=0, column=0, padx=10, pady=8, sticky="w"
)
campo_data = ctk.CTkEntry(formulario, placeholder_text="DD/MM/AAAA")
campo_data.grid(row=0, column=1, padx=10, pady=8, sticky="ew")
campo_data.insert(0, datetime.now().strftime("%d/%m/%Y"))

ctk.CTkLabel(formulario, text="Quantidade de tempos:").grid(
    row=0, column=2, padx=10, pady=8, sticky="w"
)

# Lista de seleção: não precisa digitar a quantidade
campo_tempos = ctk.CTkComboBox(
    formulario,
    values=["1", "2", "3", "4"],
    state="readonly"
)
campo_tempos.grid(row=0, column=3, padx=10, pady=8, sticky="ew")
campo_tempos.set("1")

ctk.CTkLabel(formulario, text="Tema da instrução:").grid(
    row=1, column=0, padx=10, pady=8, sticky="w"
)
campo_tema = ctk.CTkEntry(formulario, placeholder_text="Tema da instrução")
campo_tema.grid(row=1, column=1, padx=10, pady=8, sticky="ew")

ctk.CTkLabel(formulario, text="Instrutor:").grid(
    row=1, column=2, padx=10, pady=8, sticky="w"
)
campo_instrutor = ctk.CTkEntry(formulario, placeholder_text="Nome do instrutor")
campo_instrutor.grid(row=1, column=3, padx=10, pady=8, sticky="ew")

ctk.CTkLabel(formulario, text="Destinatário:").grid(
    row=2, column=0, padx=10, pady=8, sticky="w"
)
campo_destinatario = ctk.CTkComboBox(
    formulario,
    values=["ESI", "EPA", "ESI e EPA"],
    state="readonly"
)
campo_destinatario.grid(row=2, column=1, padx=10, pady=8, sticky="ew")
campo_destinatario.set("ESI")

botoes_cadastro = ctk.CTkFrame(frame_cadastro, fg_color="transparent")
botoes_cadastro.pack(pady=10)

ctk.CTkButton(
    botoes_cadastro,
    text="Salvar instrução",
    command=salvar_instrucao,
    width=180,
    height=38
).pack(side="left", padx=8)

ctk.CTkButton(
    botoes_cadastro,
    text="Limpar campos",
    command=limpar_campos,
    width=150,
    height=38,
    fg_color="gray"
).pack(side="left", padx=8)


# ============================================================
# EXPORTAÇÃO SEMANAL
# ============================================================

frame_planilha = ctk.CTkFrame(container_principal)
frame_planilha.pack(fill="x", padx=10, pady=10)

ctk.CTkLabel(
    frame_planilha,
    text="GERAR PLANILHA SEMANAL",
    font=ctk.CTkFont(size=18, weight="bold")
).pack(pady=12)

ctk.CTkLabel(
    frame_planilha,
    text="Informe uma data da semana desejada. A planilha incluirá de segunda a sexta-feira."
).pack(pady=3)

frame_data_planilha = ctk.CTkFrame(frame_planilha, fg_color="transparent")
frame_data_planilha.pack(pady=8)

ctk.CTkLabel(frame_data_planilha, text="Data da semana:").pack(side="left", padx=8)

campo_data_planilha = ctk.CTkEntry(
    frame_data_planilha,
    placeholder_text="DD/MM/AAAA",
    width=150
)
campo_data_planilha.pack(side="left", padx=8)
campo_data_planilha.insert(0, datetime.now().strftime("%d/%m/%Y"))

ctk.CTkButton(
    frame_data_planilha,
    text="Gerar e salvar Excel",
    command=gerar_planilha,
    width=190,
    height=38
).pack(side="left", padx=8)


# ============================================================
# CONSULTA
# ============================================================

frame_consulta = ctk.CTkFrame(container_principal)
frame_consulta.pack(fill="both", expand=True, padx=10, pady=10)

ctk.CTkLabel(
    frame_consulta,
    text="CONSULTAR INSTRUÇÕES",
    font=ctk.CTkFont(size=18, weight="bold")
).pack(pady=12)

frame_filtro = ctk.CTkFrame(frame_consulta, fg_color="transparent")
frame_filtro.pack(fill="x", padx=10, pady=5)

ctk.CTkLabel(frame_filtro, text="Filtrar por data:").pack(side="left", padx=5)

campo_filtro_data = ctk.CTkEntry(
    frame_filtro,
    placeholder_text="DD/MM/AAAA",
    width=150
)
campo_filtro_data.pack(side="left", padx=5)


def formatar_data_filtro(event=None):
    texto = campo_filtro_data.get()

    # Mantém somente os números
    numeros = "".join(caractere for caractere in texto if caractere.isdigit())

    # Limita a 8 números: DDMMYYYY
    numeros = numeros[:8]

    # Insere as barras automaticamente
    if len(numeros) >= 5:
        texto_formatado = numeros[:2] + "/" + numeros[2:4] + "/" + numeros[4:]
    elif len(numeros) >= 3:
        texto_formatado = numeros[:2] + "/" + numeros[2:]
    else:
        texto_formatado = numeros

    if texto_formatado != texto:
        campo_filtro_data.delete(0, tk.END)
        campo_filtro_data.insert(0, texto_formatado)
        campo_filtro_data.icursor(tk.END)


campo_filtro_data.bind("<KeyRelease>", formatar_data_filtro)


ctk.CTkButton(
    frame_filtro,
    text="Filtrar",
    command=atualizar_consulta,
    width=100
).pack(side="left", padx=5)

ctk.CTkButton(
    frame_filtro,
    text="Mostrar todas",
    command=limpar_filtro,
    width=120,
    fg_color="gray"
).pack(side="left", padx=5)


# Tabela de consulta
frame_tabela = ctk.CTkFrame(frame_consulta)
frame_tabela.pack(fill="both", expand=True, padx=10, pady=10)

colunas = (
    "id",
    "data",
    "tempos",
    "tema",
    "instrutor",
    "destinatario"
)

tabela_consulta = ttk.Treeview(
    frame_tabela,
    columns=colunas,
    show="headings",
    height=12
)

tabela_consulta.heading("id", text="ID")
tabela_consulta.heading("data", text="Data")
tabela_consulta.heading("tempos", text="Tempos")
tabela_consulta.heading("tema", text="Tema da instrução")
tabela_consulta.heading("instrutor", text="Instrutor")
tabela_consulta.heading("destinatario", text="Destinatário")

tabela_consulta.column("id", width=50, anchor="center")
tabela_consulta.column("data", width=100, anchor="center")
tabela_consulta.column("tempos", width=70, anchor="center")
tabela_consulta.column("tema", width=300, anchor="w")
tabela_consulta.column("instrutor", width=180, anchor="w")
tabela_consulta.column("destinatario", width=130, anchor="center")

scroll_vertical = ttk.Scrollbar(
    frame_tabela,
    orient="vertical",
    command=tabela_consulta.yview
)
tabela_consulta.configure(yscrollcommand=scroll_vertical.set)

tabela_consulta.pack(side="left", fill="both", expand=True)
scroll_vertical.pack(side="right", fill="y")


# Botões de edição e exclusão
frame_botoes_consulta = ctk.CTkFrame(frame_consulta, fg_color="transparent")
frame_botoes_consulta.pack(pady=10)

ctk.CTkButton(
    frame_botoes_consulta,
    text="Editar instrução selecionada",
    command=editar_instrucao,
    width=220,
    height=38
).pack(side="left", padx=8)

ctk.CTkButton(
    frame_botoes_consulta,
    text="Excluir instrução selecionada",
    command=excluir_instrucao,
    width=220,
    height=38,
    fg_color="#B22222",
    hover_color="#8B0000"
).pack(side="left", padx=8)

ctk.CTkButton(
    frame_botoes_consulta,
    text="Atualizar consulta",
    command=atualizar_consulta,
    width=150,
    height=38,
    fg_color="gray"
).pack(side="left", padx=8)


# Rodapé
rodape = ctk.CTkLabel(
    container_principal,
    text="SAIM — Sistema de Automação de Instruções Militares",
    font=ctk.CTkFont(size=11),
    text_color="gray"
)
rodape.pack(pady=10)


# ============================================================
# INICIALIZAÇÃO
# ============================================================

inicializar_banco()
atualizar_consulta()

app.mainloop()