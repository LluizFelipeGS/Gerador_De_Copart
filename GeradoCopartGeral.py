import flet as ft
import pdfplumber
import pandas as pd
import re
import os
import sys


def resource_path(relative_path):
    base_path = getattr(sys, "_MEIPASS", os.path.abspath("."))
    return os.path.join(base_path, relative_path)


def fechar_dialogo(dialog, page):
    dialog.open = False
    page.update()


def gerar_excel(pdf_path, excel_path, competencia, operadora, page):
    try:
        text = ""
        with pdfplumber.open(pdf_path) as pdf:
            for page_ in pdf.pages:
                txt = page_.extract_text()
                if txt:
                    text += txt + "\n"

        text_clean = re.sub(r'\n+', '\n', text).strip()

        # Ajustes para operações SB Saúde
        if operadora == "SB Saúde":
            # Ajustar espaçamento caso precise
            text_clean = re.sub(r"(\d{2}/\d{2}/\d{4})([A-Z])", r"\1 \2", text_clean)
            text_clean = re.sub(r"([A-ZÀ-ÿ\s\.]+)(-?)(\d{6,})", r"\1 - \3", text_clean)
            text_clean = re.sub(r"\s+(R\$*\s*[\d.,]+)", r"\n\1", text_clean)

        dados = []

        if operadora == "Nordeste Saude":

            titular_match = re.search(
                r"Titular:\s*\((\d+\-\d+)\)\s*(.*?)\s*(?=\s*Titular|Empresa|$)",
                text_clean, re.IGNORECASE
            )
            numero_id = titular_match.group(1).strip() if titular_match else ""
            titular_nome = titular_match.group(2).strip() if titular_match else ""
            beneficiario_nome = titular_nome

            padrao = re.compile(
                r"(\d{2}/\d{2}/\d{4})\s+"
                r"(\d{2}/\d{2}/\d{4})\s+"
                r"CL[IÍ]NICA\s*:\s*(.*?)\s+"
                r"(\d{6,})\s+"
                r"([A-Z\s]+?)\s+"
                r"(\(\d{1,2}\.\d{1,2}\.\d{1,2}\.\d{1,3}-\d\).*?)\s+"
                r"([0-9]+[.,][0-9]{2})(?![\d])",
                re.IGNORECASE
            )

            for match in padrao.finditer(text_clean):
                dt_liberacao = match.group(1)
                dt_realizacao = match.group(2)
                local = match.group(3).strip()
                senha = match.group(4)
                especialidade = match.group(5).strip()
                procedimento_completo = match.group(6).strip()
                valor = match.group(7).replace(",", ".")

                numero_proc = re.match(r"\((.*?)\)", procedimento_completo)
                numero_procedimento = numero_proc.group(1) if numero_proc else ""
                nome_procedimento = (
                    procedimento_completo.replace(f"({numero_procedimento})", "").strip()
                    if numero_procedimento
                    else procedimento_completo.strip()
                )

                tipo_contrato = "TITULAR" if beneficiario_nome == titular_nome else "DEPENDENTE"

                dados.append({
                    "Titular": titular_nome,
                    "Número de Identificação": numero_id,
                    "Tipo de Contrato": tipo_contrato,
                    "Dt. Liberação": dt_liberacao,
                    "Dt. Realização": dt_realizacao,
                    "Local": local,
                    "Senha": senha,
                    "Especialidade": especialidade,
                    "Competência": competencia,
                    "Número de Identificação de Procedimento": numero_procedimento,
                    "Procedimento": nome_procedimento,
                    "Vr. Copart": float(valor)
                })

        elif operadora == "SB Saúde":

            # Captura titular/empresa
            titular_match = re.search(r"Empresa:\s*(.*)", text_clean, re.IGNORECASE)
            titular_nome = titular_match.group(1).strip() if titular_match else "Titular não encontrado"

            competencia_match = re.search(r"Referencia:\s*(\d{2}/\d{4})", text_clean, re.IGNORECASE)
            competencia_doc = competencia_match.group(1) if competencia_match else competencia

            # Quebrar o texto em linhas e formar registros (blocos) de atendimento
            lines = text_clean.splitlines()
            registros = []
            temp_registro = []

            for line in lines:
                line = line.strip()
                if re.match(r"^\d{2}/\d{2}/\d{4}", line):
                    if temp_registro:
                        registros.append(temp_registro)
                    temp_registro = [line]
                else:
                    if temp_registro:
                        temp_registro.append(line)
                    else:
                        # ignorar linhas antes do primeiro registro
                        continue
            if temp_registro:
                registros.append(temp_registro)

            for bloco in registros:
                # ignorar blocos vazios
                if not bloco:
                    continue

                data_match = re.match(r"(\d{2}/\d{2}/\d{4})(.*)", bloco[0])
                if not data_match:
                    continue

                data_atendimento = data_match.group(1)
                tipo_procedimento = data_match.group(2).strip()

                prestador = bloco[1] if len(bloco) > 1 else ""

                codigo = ""
                conta = ""
                quantidade = 1
                valor = 0.0
                beneficiario = ""

                for linha in bloco[2:]:
                    linha = linha.strip()

                    # Ignorar linhas de total e rodapé
                    if linha.lower().startswith(("total", "página", "soma")):
                        continue

                    # Beneficiário geralmente vem seguido de 'Soma' - capturar iten anterior e limpar a palavra
                    if "soma" in linha.lower():
                        # Como beneficiário vem antes da palavra soma, tentar buscar no registro anterior
                        continue

                    # Por vezes beneficiário pode estar concatenado com soma - vamos separar:
                    if re.match(r"^[A-ZÇÑÀ-ÿ\s]+$", linha):
                        # Linha provavelmente é beneficiário
                        beneficiario = linha.strip()
                        continue

                    if re.fullmatch(r"\d{6,}", linha):
                        codigo = linha
                        continue

                    if re.fullmatch(r"\d{1,6}", linha) and not conta:
                        conta = linha
                        continue

                    if re.fullmatch(r"\d+", linha):
                        val_int = int(linha)
                        # Quantidade frequentemente é número entre 1 e 9 e diferente do conta
                        if 0 < val_int < 10 and val_int != int(conta or 0):
                            quantidade = val_int
                            continue

                    vr_match = re.match(r"R\$\s*([\d.,]+)", linha)
                    if vr_match:
                        try:
                            valor = float(vr_match.group(1).replace(".", "").replace(",", "."))
                        except ValueError:
                            valor = 0.0
                        continue

                # Se beneficiario não aparece, usar titular
                if not beneficiario:
                    beneficiario = titular_nome

                dados.append({
                    "Nome do Titular": titular_nome,
                    "Beneficiário": beneficiario,
                    "Data Atendimento": data_atendimento,
                    "Tipo Procedimento": tipo_procedimento,
                    "Prestador": prestador,
                    "Código de Procedimento": codigo,
                    "Contas": conta,
                    "Quantidade": quantidade,
                    "Competência": competencia_doc,
                    "Vr. Copart": valor
                })

            if not dados:
                print("⚠️ Nenhum dado encontrado para SB Saúde no documento.")

        else:
            raise ValueError(f"Operadora não reconhecida: {operadora}")

        # Gravação do Excel e mensagens dialog
        if dados:
            try:
                df = pd.DataFrame(dados)
                df.to_excel(excel_path, index=False, engine='openpyxl')

                dialog = ft.AlertDialog(
                    title=ft.Text("Sucesso"),
                    content=ft.Text("Relatório gerado com sucesso!"),
                    actions=[ft.TextButton("Fechar", on_click=lambda e: fechar_dialogo(dialog, page))],
                    actions_alignment="end"
                )

            except PermissionError:
                dialog = ft.AlertDialog(
                    title=ft.Text("Erro"),
                    content=ft.Text(
                        "Permissão negada ao salvar o arquivo. Verifique se ele está aberto ou se há restrições de acesso:\n" + excel_path
                    ),
                    actions=[ft.TextButton("Fechar", on_click=lambda e: fechar_dialogo(dialog, page))],
                    actions_alignment="end"
                )

        else:
            dialog = ft.AlertDialog(
                title=ft.Text("Aviso"),
                content=ft.Text("Nenhum dado foi encontrado."),
                actions=[ft.TextButton("Fechar", on_click=lambda e: fechar_dialogo(dialog, page))],
                actions_alignment="end"
            )

        page.dialog = dialog
        dialog.open = True
        page.add(dialog)
        page.update()

    except Exception as e:
        dialog = ft.AlertDialog(
            title=ft.Text("Erro"),
            content=ft.Text(str(e)),
            actions=[ft.TextButton("Fechar", on_click=lambda e: fechar_dialogo(dialog, page))],
            actions_alignment="end"
        )
        page.dialog = dialog
        dialog.open = True
        page.add(dialog)
        page.update()


def main(page: ft.Page):
    page.title = "Gerador de Coparticipação"
    page.theme_mode = "light"
    page.window_width = 600
    page.window_height = 400
    page.vertical_alignment = "start"

    icon_path = resource_path("LOGO-QUALI.ico")
    page.window_icon = icon_path

    logo_empresa = ft.Image(src=icon_path, width=120)

    pdf_path = ft.TextField(label="Selecionar um arquivo PDF", width=450)
    excel_path = ft.TextField(label="Salvar como (.xlsx)", width=450)
    competencia = ft.TextField(label="Competência (MM/AAAA)", width=200)

    operadora_dropdown = ft.Dropdown(
        label="Selecione a Operadora",
        width=300,
        options=[
            ft.dropdown.Option("Nordeste Saude"),
            ft.dropdown.Option("SB Saúde")
        ]
    )

    file_picker = ft.FilePicker()
    save_picker = ft.FilePicker()
    page.overlay.extend([file_picker, save_picker])

    cabecalho = ft.Column(
        controls=[
            logo_empresa,
            ft.Text("Gerador de Relatório de Coparticipação", size=22, weight="bold", text_align="center")
        ],
        horizontal_alignment="center"
    )

    def selecionar_pdf(e):
        def on_result(ev):
            if ev.files:
                pdf_path.value = ev.files[0].path
                page.update()
        file_picker.on_result = on_result
        file_picker.pick_files(allow_multiple=False)

    def salvar_excel_dialog(e):
        def on_result(ev):
            if ev.path:
                path = ev.path
                if not path.lower().endswith(".xlsx"):
                    path += ".xlsx"
                excel_path.value = path
                page.update()
        save_picker.on_result = on_result
        save_picker.save_file(
            dialog_title="Salvar Excel",
            file_name="relatorio.xlsx"
        )

    def iniciar(e):
        if not pdf_path.value or not excel_path.value or not competencia.value or not operadora_dropdown.value:
            dialog = ft.AlertDialog(
                title=ft.Text("Atenção"),
                content=ft.Text("Preencha todos os campos e selecione a operadora."),
                actions=[ft.TextButton("Fechar", on_click=lambda e: fechar_dialogo(dialog, page))],
                actions_alignment="end"
            )
            page.dialog = dialog
            dialog.open = True
            page.add(dialog)
            page.update()
            return

        gerar_excel(
            pdf_path=pdf_path.value,
            excel_path=excel_path.value,
            competencia=competencia.value,
            operadora=operadora_dropdown.value,
            page=page
        )

    page.add(
        cabecalho,
        ft.Row([
            pdf_path,
            ft.TextButton("Selecionar", on_click=selecionar_pdf)
        ]),
        ft.Row([
            excel_path,
            ft.TextButton("Salvar como", on_click=salvar_excel_dialog)
        ]),
        competencia,
        operadora_dropdown,
        ft.ElevatedButton("Gerar Relatório", on_click=iniciar),
    )


if __name__ == "__main__":
    ft.app(target=main)
