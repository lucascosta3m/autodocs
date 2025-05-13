# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import messagebox
import gspread
from google.oauth2.service_account import Credentials
from gspread.utils import rowcol_to_a1
from docx import Document
from pathlib import Path
import re
import os
import sys
import time  # Opcional para delays

# ==============================================================================
# 1. FUNÇÃO AUXILIAR PARA CAMINHOS (PyInstaller)
# ==============================================================================
def resource_path(relative_path):
    """ Obtém o caminho absoluto para o recurso, funciona para dev e PyInstaller """
    try:
        base_path = Path(sys._MEIPASS)
    except Exception:
        try:
            base_path = Path(__file__).parent
        except NameError:
            base_path = Path(".")
    return base_path / relative_path

# ==============================================================================
# 2. CONFIGURAÇÕES GLOBAIS E CONSTANTES
# ==============================================================================
DEBUG_MODE = True # Mudar para False para menos logs

# --- Caminhos e Arquivos ---
PASTA_SAIDA = "docs_gerados" # Será criada onde o script/exe rodar
PASTA_TEMPLATES_REL = "templates"
CAMINHO_CREDENCIAL_REL = "credenciais.json"

# Converte caminhos relativos para absolutos (funciona em dev e no .exe)
PASTA_TEMPLATES = resource_path(PASTA_TEMPLATES_REL)
CAMINHO_CREDENCIAL = resource_path(CAMINHO_CREDENCIAL_REL)

# --- Nomes Google Sheets ---
PLANILHA_PF_FILENAME = "FORMULÁRIO PESSOA FÍSICA (respostas)"
PLANILHA_PJ_FILENAME = "FORMULÁRIO PESSOA JURÍDICA (respostas)"
PF_TAB_NAME = "Respostas ao formulário PF"
PJ_TAB_NAME = "Respostas ao formulário PJ"

# --- Templates DOCX ---
TEMPLATE_PF = [PASTA_TEMPLATES / f"PF{i}- {name}.docx" for i, name in enumerate(["FICHA DE INSCRICAO", "INSTRUMENTO", "PROPOSTA DE ADMISSAO", "TERMO DE ADESAO"], 1)]
TEMPLATE_PJ = [PASTA_TEMPLATES / f"PJ{i}- {name}.docx" for i, name in enumerate(["FICHA DE INSCRICAO", "INSTRUMENTO", "PROPOSTA DE ADMISSAO", "TERMO DE ADESAO"], 1)]

# --- Colunas Chave Planilha ---
TRIGGER_VALUE = "JÁ CADASTREI MEUS DADOS PESSOAIS, QUERO CADASTRAR OUTRO VEÍCULO"
COL_CADASTRO = "CADASTRO"
COL_PF_ID_TRIGGER = "CPF - (SOMENTE NÚMERO)"
COL_PF_ID_COMPARISON = "CPF (somente número)"
COL_PJ_ID_TRIGGER = "CNPJ - (SOMENTE NÚMERO)"
COL_PJ_ID_COMPARISON = "CNPJ (somente número)"
STATUS_COL = "Status"

# ==============================================================================
# 3. FUNÇÕES AUXILIARES (Formatação, Substituição)
# ==============================================================================
def formatar_cpf(cpf_input):
    """Formata CPF para xxx.xxx.xxx-xx."""
    if cpf_input is None: return ""
    cpf = str(cpf_input).strip().lstrip("'")
    cpf_limpo = re.sub(r'\D', '', cpf).zfill(11)
    if len(cpf_limpo) == 11: return f'{cpf_limpo[:3]}.{cpf_limpo[3:6]}.{cpf_limpo[6:9]}-{cpf_limpo[9:]}'
    if DEBUG_MODE: print(f"DBG: CPF inválido para formatação: '{cpf_input}'")
    return str(cpf_input)

def formatar_cnpj(cnpj_input):
    """Formata CNPJ para xx.xxx.xxx/xxxx-xx."""
    if cnpj_input is None: return ""
    cnpj = str(cnpj_input).strip().lstrip("'")
    cnpj_limpo = re.sub(r'\D', '', cnpj).zfill(14)
    if len(cnpj_limpo) == 14: return f'{cnpj_limpo[:2]}.{cnpj_limpo[2:5]}.{cnpj_limpo[5:8]}/{cnpj_limpo[8:12]}-{cnpj_limpo[12:]}'
    if DEBUG_MODE: print(f"DBG: CNPJ inválido para formatação: '{cnpj_input}'")
    return str(cnpj_input)

def substituir_placeholders(document, dados):
    """Substitui placeholders {CHAVE} no documento preservando formatação (em runs)."""
    formatadores = {
        COL_PF_ID_COMPARISON.upper(): formatar_cpf,
        COL_PJ_ID_COMPARISON.upper(): formatar_cnpj,
        COL_PF_ID_TRIGGER.upper(): formatar_cpf,
        COL_PJ_ID_TRIGGER.upper(): formatar_cnpj,
    }

    placeholders = {}
    for chave, valor in dados.items():
        if str(chave).lower() in ["tipo", "linha"]:
            continue
        chave_fmt = str(chave).strip().upper()
        valor_fmt = str(valor).strip() if valor is not None else ""
        formatador = formatadores.get(chave_fmt)
        if formatador:
            valor_fmt = formatador(valor_fmt)
        elif not valor_fmt.isdigit():
            valor_fmt = valor_fmt.upper()
        placeholders[f"{{{chave_fmt}}}"] = valor_fmt

    def substituir_em_runs(runs):
        for run in runs:
            for ph, val in placeholders.items():
                if ph in run.text:
                    run.text = run.text.replace(ph, val)

    # Substituição em parágrafos
    for paragraph in document.paragraphs:
        substituir_em_runs(paragraph.runs)

    # Substituição em tabelas
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    substituir_em_runs(paragraph.runs)


# ==============================================================================
# 4. FUNÇÃO DE PRÉ-PREENCHIMENTO
# ==============================================================================
def preencher_e_atualizar_planilha(sheet, headers, dados_originais, id_col_trigger, id_col_comparison, col_cadastro, trigger_value):
    """Preenche dados baseados em cadastros anteriores e atualiza planilha."""
    if not dados_originais:
        print(f"Aviso Pré-preenchimento: Planilha '{sheet.title}' vazia.")
        return []

    print(f"\n--- Iniciando pré-preenchimento para: {sheet.title} ---")
    try:
        # Cabeçalhos já são passados como argumento, valida se colunas existem
        colunas_essenciais = [col_cadastro, id_col_trigger, id_col_comparison]
        for col in colunas_essenciais:
             if col not in headers: raise ValueError(f"Coluna essencial '{col}' não encontrada.")

        idx_col_comparison = headers.index(id_col_comparison) + 1
        col_indices = {name: i + 1 for i, name in enumerate(headers)}
        cols_to_fill = { name: index for name, index in col_indices.items()
                         if name not in [col_cadastro, id_col_trigger, id_col_comparison, STATUS_COL, "Timestamp", "Carimbo de data/hora"] }
        if DEBUG_MODE: print(f" Colunas a preencher da fonte (se vazias): {list(cols_to_fill.keys())}")

    except ValueError as e:
        messagebox.showerror("Erro Config Pré-preenchimento", f"{e} na planilha '{sheet.title}'.")
        return dados_originais # Retorna dados originais se houve erro nos headers

    source_data_map = {}; dados_modificados = []; linhas_para_buscar = []
    print(" Mapeando dados de origem e identificando alvos...")
    for i, row_dict in enumerate(dados_originais):
        clean_row = {str(k).strip(): str(v).strip() if v is not None else "" for k, v in row_dict.items() if k}
        dados_modificados.append(clean_row)
        row_num = i + 2
        cadastro_status = clean_row.get(col_cadastro, "")
        comp_id = clean_row.get(id_col_comparison, "")
        trig_id = clean_row.get(id_col_trigger, "")

        if cadastro_status != trigger_value and comp_id: # Linha fonte
            if comp_id not in source_data_map: source_data_map[comp_id] = clean_row
        elif cadastro_status == trigger_value and trig_id: # Linha alvo
            linhas_para_buscar.append((i, trig_id, row_num))
        elif cadastro_status == trigger_value and not trig_id:
             print(f" Aviso Pré-preenchimento: Linha {row_num} ({sheet.title}) com trigger mas sem ID em '{id_col_trigger}'.")

    batch_updates = []
    print(f" Processando {len(linhas_para_buscar)} linha(s) alvo...")
    for list_idx, trigger_id, row_num in linhas_para_buscar:
        if list_idx >= len(dados_modificados): continue # Segurança
        target_row = dados_modificados[list_idx]
        updated = False
        # Passo 1: Preencher ID de Comparação se vazio
        if not target_row.get(id_col_comparison, "") and trigger_id:
            target_row[id_col_comparison] = trigger_id
            cell_a1 = rowcol_to_a1(row_num, idx_col_comparison)
            batch_updates.append({'range': cell_a1, 'values': [[str(trigger_id)]]})
            if DEBUG_MODE: print(f"  L{row_num}: Preenchendo '{id_col_comparison}' com '{trigger_id}'")
            updated = True
        # Passo 2: Preencher outros campos da fonte
        if trigger_id in source_data_map:
            source_row = source_data_map[trigger_id]
            for col_name, col_idx in cols_to_fill.items():
                if not target_row.get(col_name, ""): # Só preenche se vazio
                    source_value = source_row.get(col_name, "")
                    if source_value: # Só preenche se fonte não for vazia
                        target_row[col_name] = source_value
                        cell_a1 = rowcol_to_a1(row_num, col_idx)
                        batch_updates.append({'range': cell_a1, 'values': [[str(source_value)]]})
                        if DEBUG_MODE: print(f"  L{row_num}: Preenchendo '{col_name}' com valor da fonte.")
                        updated = True
        elif DEBUG_MODE: print(f"  L{row_num}: Fonte não encontrada para ID '{trigger_id}'.")
        if not updated and DEBUG_MODE: print(f"  L{row_num}: Nenhuma atualização necessária.")

    if batch_updates:
        print(f" Enviando {len(batch_updates)} atualizações de pré-preenchimento para '{sheet.title}'...")
        try:
            sheet.batch_update(batch_updates, value_input_option='USER_ENTERED')
            print(" Pré-preenchimento salvo com sucesso!")
        except Exception as e:
             print(f"ERRO ao salvar pré-preenchimento em '{sheet.title}': {e}")
             messagebox.showerror("Erro API Google (Pré-preenchimento)", f"Falha ao salvar pré-preenchimento em '{sheet.title}':\n{e}")
    else: print(f" Nenhuma atualização de pré-preenchimento necessária para '{sheet.title}'.")
    print(f"--- Fim do pré-preenchimento para: {sheet.title} ---\n")
    return dados_modificados

# ==============================================================================
# 5. FUNÇÕES DE INTERAÇÃO COM PLANILHAS (Carregar)
# ==============================================================================
def autenticar_google():
    """Autentica com a API do Google Sheets e retorna o cliente gspread."""
    print("Autenticando com Google...")
    scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    if not CAMINHO_CREDENCIAL.exists():
         raise FileNotFoundError(f"Arquivo de credenciais não encontrado em: {CAMINHO_CREDENCIAL}")
    credenciais = Credentials.from_service_account_file(CAMINHO_CREDENCIAL, scopes=scopes)
    gc = gspread.authorize(credenciais)
    print("Autenticado.")
    return gc

def carregar_planilha(gc, filename, tab_name):
    """Abre uma planilha e aba específica, lê cabeçalho e dados."""
    print(f"\nAbrindo arquivo: {filename}")
    try:
        workbook = gc.open(filename)
        print(f"Tentando acessar aba: '{tab_name}'")
        sheet = workbook.worksheet(tab_name)
        print(f"Aba '{sheet.title}' (ID: {sheet.id}) acessada.")
        print(f"Lendo cabeçalho e dados da aba '{sheet.title}'...")
        headers = [str(h).strip() for h in sheet.row_values(1)]
        if not headers: raise ValueError("Cabeçalho vazio ou não lido.")
        # Validar colunas essenciais globais (Status)
        if STATUS_COL not in headers: raise ValueError(f"Coluna '{STATUS_COL}' não encontrada.")
        data = sheet.get_all_records(head=1)
        print(f"Lidos {len(data)} registros de '{sheet.title}'.")
        return sheet, headers, data
    except gspread.exceptions.SpreadsheetNotFound:
        messagebox.showerror("Erro Crítico", f"Arquivo '{filename}' não encontrado. Verifique o nome e permissões.")
        sys.exit()
    except gspread.exceptions.WorksheetNotFound:
        messagebox.showerror("Erro Crítico", f"Aba '{tab_name}' não encontrada em '{filename}'. Verifique o nome exato.")
        sys.exit()
    except ValueError as e:
         messagebox.showerror("Erro Cabeçalho", f"Erro no cabeçalho da aba '{tab_name}' em '{filename}': {e}")
         sys.exit()
    except Exception as e:
        messagebox.showerror("Erro ao Carregar Planilha", f"Erro ao carregar '{filename}' / Aba '{tab_name}':\n{type(e).__name__}: {e}")
        sys.exit()

# ==============================================================================
# 6. FUNÇÕES PRINCIPAIS DA APLICAÇÃO (Gerar Docs, Excluir)
# ==============================================================================

# Variáveis globais para acesso pelas funções de comando dos botões
# (Alternativa seria usar uma classe App, mas mantendo funcional por ora)
root = None
checkboxes_pf = []
checkboxes_pj = []
dados_pf_processados = []
dados_pj_processados = []
sheet_pf = None
sheet_pj = None
col_index_status_pf = -1
col_index_status_pj = -1

def gerar_documentos_cmd():
    """Função chamada pelo botão 'Gerar Documentos'."""
    global checkboxes_pf, checkboxes_pj, sheet_pf, sheet_pj, col_index_status_pf, col_index_status_pj, root

    selecionados_tuplas = [(p, var) for p, var in checkboxes_pf + checkboxes_pj if var.get()]
    if not selecionados_tuplas:
        messagebox.showwarning("Aviso", "Nenhum registro selecionado.")
        return

    print(f"\n--- Gerando Docs para {len(selecionados_tuplas)} registro(s) ---")
    try:
        os.makedirs(PASTA_SAIDA, exist_ok=True)
    except OSError as e:
         messagebox.showerror("Erro ao Criar Pasta", f"Não foi possível criar a pasta '{PASTA_SAIDA}': {e}")
         return

    erros_template = []; erros_geracao = []; erros_atualizacao = []
    updates_status_batch = []

    for pessoa_original, var_tk in selecionados_tuplas:
        dados_pessoa = {k.strip().upper(): v for k, v in pessoa_original.items()}
        tipo = dados_pessoa.get('TIPO')
        linha = pessoa_original.get("linha")

        if tipo == "PF":
            templates, ws, col_idx = TEMPLATE_PF, sheet_pf, col_index_status_pf
        elif tipo == "PJ":
            templates, ws, col_idx = TEMPLATE_PJ, sheet_pj, col_index_status_pj
        else:
            erros_geracao.append(f"L{linha}: Tipo '{tipo}' inválido."); continue

        nome_base = str(dados_pessoa.get("NOME COMPLETO", dados_pessoa.get("RAZÃO SOCIAL", f"Reg_L{linha}"))).strip().replace(" ", "_")
        placa_base = str(dados_pessoa.get("PLACA", "SemPlaca")).strip().replace('-', '')
        if not isinstance(linha, int) or linha < 2:
             erros_geracao.append(f"{nome_base}: Linha inválida ({linha})."); continue

        if DEBUG_MODE: print(f"\n=== PROC REG: {nome_base} (L{linha}, {tipo}) ===")
        pasta_destino = Path(PASTA_SAIDA) / nome_base
        pasta_destino.mkdir(parents=True, exist_ok=True)
        todos_templates_ok = True

        for template_path_obj in templates:
            if not template_path_obj.exists():
                msg = f"Template não encontrado: {template_path_obj.name}"; erros_template.append(msg); todos_templates_ok = False; continue
            if DEBUG_MODE: print(f"-- Proc Template: {template_path_obj.name}")
            try:
                doc = Document(template_path_obj)
                substituir_placeholders(doc, dados_pessoa)
                nome_safe = re.sub(r'[\\/*?:"<>|]', "", nome_base)[:50]
                placa_safe = re.sub(r'[\\/*?:"<>|]', "", placa_base)
                prefixo = template_path_obj.stem.split('-', 1)[0].strip()
                nome_doc = f"{prefixo}_{nome_safe}_{placa_safe}.docx"
                caminho_saida = pasta_destino / nome_doc
                doc.save(caminho_saida)
                if DEBUG_MODE: print(f"  >> Salvo: {nome_doc}")
            except Exception as e:
                msg = f"L{linha} ({nome_base}): Erro ao gerar '{template_path_obj.name}': {e}"; print(f"❌ {msg}"); erros_geracao.append(msg); todos_templates_ok = False

        if todos_templates_ok and col_idx > 0:
            if DEBUG_MODE: print(f"  >> Docs OK. Preparando update status '{ws.title}' L{linha} C{col_idx}")
            try:
                cell_a1 = rowcol_to_a1(linha, col_idx)
                updates_status_batch.append({'range': cell_a1, 'values': [['GERADO']], 'worksheet': ws})
            except Exception as e: msg = f"L{linha} ({nome_base}): Erro ao preparar A1 para status: {e}"; print(f"❌ {msg}"); erros_atualizacao.append(msg)
        elif todos_templates_ok and col_idx <= 0:
            msg = f"L{linha} ({nome_base}): Docs gerados, mas coluna STATUS não encontrada em '{ws.title}'."; print(f"⚠️ {msg}"); erros_atualizacao.append(msg)
        elif DEBUG_MODE: print(f"  >> Geração incompleta. Status NÃO será atualizado.")

    # Envio dos Status em Lote
    status_atualizado_count = 0
    if updates_status_batch:
        print(f"\n--- Enviando {len(updates_status_batch)} updates de status ---")
        updates_agrupados = {}
        for info in updates_status_batch: # Agrupa pelo objeto worksheet
            ws = info['worksheet']
            if ws not in updates_agrupados: updates_agrupados[ws] = []
            updates_agrupados[ws].append({'range': info['range'], 'values': info['values']})

        for sheet_obj, updates_list in updates_agrupados.items():
            title = sheet_obj.title
            print(f"\n>>> Enviando {len(updates_list)} updates para '{title}'...")
            try:
                response = sheet_obj.batch_update(updates_list, value_input_option='USER_ENTERED')
                print(f"    ✅ API call para '{title}' CONCLUÍDA sem exceções.")
                if DEBUG_MODE: print(f"    Resposta API: {response}")
                status_atualizado_count += len(updates_list)
            except Exception as e:
                msg = f"    ❌ ERRO ao atualizar status em '{title}': {e}"; print(msg); erros_atualizacao.append(msg)
            print(f"<<< Finalizado envio para '{title}'")
        print("\n--- Fim do envio de status ---")
    else: print("\nNenhum update de status a enviar.")

    # Monta Relatório Final
    msg_final = [f"Processo Concluído.", f"Selecionados: {len(selecionados_tuplas)}", f"Status 'GERADO' atualizado (API OK): {status_atualizado_count}"]
    if erros_template: msg_final.extend(["\n--- Templates Não Encontrados ---"] + list(set(erros_template)))
    if erros_geracao: msg_final.extend([f"\n--- Erros Geração/Salvar Docs ({len(erros_geracao)}) ---"] + erros_geracao[:5] + ["(... ver console)"] if len(erros_geracao) > 5 else erros_geracao)
    if erros_atualizacao: msg_final.extend([f"\n--- Erros Preparação/Envio Status ({len(erros_atualizacao)}) ---"] + erros_atualizacao[:5] + ["(... ver console)"] if len(erros_atualizacao) > 5 else erros_atualizacao)
    messagebox.showinfo("Relatório Final", "\n".join(msg_final))

    print("Recarregando interface...")
    try:
        if root and root.winfo_exists(): root.destroy()
    except tk.TclError: print("Aviso: Janela já destruída (geração).")

def excluir_entradas_cmd():
    """Função chamada pelo botão 'Excluir da Planilha'."""
    global checkboxes_pf, checkboxes_pj, sheet_pf, sheet_pj, root

    selecionados_tuplas = [(p, var) for p, var in checkboxes_pf + checkboxes_pj if var.get()]
    if not selecionados_tuplas: messagebox.showwarning("Aviso", "Nenhum registro selecionado."); return
    confirm = messagebox.askyesno("Confirmar Exclusão", f"Tem certeza que deseja excluir {len(selecionados_tuplas)} registro(s)?\n\nESTA AÇÃO NÃO PODE SER DESFEITA.")
    if not confirm: return

    print(f"\n--- Excluindo {len(selecionados_tuplas)} registro(s) ---")
    excluidos_count = 0; erros_exclusao = []; requests_pf = []; requests_pj = []
    # Ordena por linha DECRESCENTE para evitar problemas de índice na exclusão em lote
    selecionados_ordenados = sorted(selecionados_tuplas, key=lambda item: item[0].get("linha", 0), reverse=True)
    try: sheet_id_pf = sheet_pf.id; sheet_id_pj = sheet_pj.id
    except Exception as e: messagebox.showerror("Erro Interno", f"Erro ao obter IDs das abas para exclusão: {e}"); return

    for pessoa_dict, _ in selecionados_ordenados:
        linha = pessoa_dict.get("linha"); tipo = pessoa_dict.get("tipo")
        nome = str(pessoa_dict.get("NOME COMPLETO", pessoa_dict.get("RAZÃO SOCIAL", "Reg Desconhecido"))).strip()
        if not isinstance(linha, int) or linha < 2: msg = f"Exclusão: Linha inválida ({linha}) para {nome}."; print(f"⚠️ {msg}"); erros_exclusao.append(msg); continue
        start_idx = linha - 1; end_idx = linha
        delete_req = {'deleteDimension': {'range': {'sheetId': None, 'dimension': 'ROWS', 'startIndex': start_idx, 'endIndex': end_idx}}}
        if tipo == "PF": delete_req['deleteDimension']['range']['sheetId'] = sheet_id_pf; requests_pf.append(delete_req)
        elif tipo == "PJ": delete_req['deleteDimension']['range']['sheetId'] = sheet_id_pj; requests_pj.append(delete_req)
        else: msg = f"Exclusão: Tipo '{tipo}' desconhecido L{linha}."; print(f"⚠️ {msg}"); erros_exclusao.append(msg)

    def executar_exclusao(sheet_obj, requests, tipo_label):
        count = 0
        if requests:
            print(f" Executando {len(requests)} exclusões na planilha {tipo_label} ('{sheet_obj.title}')...")
            try:
                sheet_obj.spreadsheet.batch_update({'requests': requests}) # Exclui usando o spreadsheet
                print(f"  > Exclusões {tipo_label} concluídas."); count = len(requests)
            except Exception as e: msg = f"Erro ao excluir {tipo_label}: {e}"; print(f"❌ {msg}"); erros_exclusao.append(msg)
        return count

    excluidos_count += executar_exclusao(sheet_pf, requests_pf, "PF")
    excluidos_count += executar_exclusao(sheet_pj, requests_pj, "PJ")

    msg_final = [f"{excluidos_count} registro(s) efetivamente excluído(s)."]
    if erros_exclusao: msg_final.extend(["\n--- Ocorrências Durante Exclusão ---"] + erros_exclusao)
    messagebox.showinfo("Relatório de Exclusão", "\n".join(msg_final))

    print("Recarregando interface após exclusão...")
    try:
        if root and root.winfo_exists(): root.destroy()
    except tk.TclError: print("Aviso: Janela já destruída (exclusão).")

# ==============================================================================
# 7. FUNÇÕES DA INTERFACE GRÁFICA (Tkinter)
# ==============================================================================
def adicionar_checkbox(pessoa_data, parent_frame, checkboxes_list):
    """Cria e adiciona um checkbox para uma pessoa/empresa."""
    var = tk.BooleanVar()
    nome = str(pessoa_data.get("NOME COMPLETO", pessoa_data.get("RAZÃO SOCIAL", "N/A"))).strip()
    placa = str(pessoa_data.get("PLACA", "N/A")).strip()
    max_len_nome = 35
    nome_display = f"{nome[:max_len_nome]}..." if len(nome) > max_len_nome else nome
    texto = f"{nome_display} - {placa}"
    cb = tk.Checkbutton(parent_frame, text=texto, variable=var, anchor="w", justify="left", wraplength=350)
    cb.pack(fill="x", padx=5, pady=1) # Menor padding vertical
    checkboxes_list.append((pessoa_data, var))

def criar_interface(root_window, dados_pf, dados_pj):
    """Cria todos os elementos da interface gráfica."""
    global checkboxes_pf, checkboxes_pj # Permite que esta função popule as listas globais

    root_window.title("Autodocs - Gerador de Documentos v1.1") # Exemplo de versão
    root_window.geometry("850x650")

    # --- Área Rolável ---
    frame_scroll_container = tk.Frame(root_window)
    frame_scroll_container.pack(pady=10, fill="both", expand=True)
    canvas = tk.Canvas(frame_scroll_container)
    scrollbar = tk.Scrollbar(frame_scroll_container, orient=tk.VERTICAL, command=canvas.yview)
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    frame_content = tk.Frame(canvas) # Frame interno que rola
    canvas.create_window((0, 0), window=frame_content, anchor="nw")
    def on_frame_configure(event): canvas.configure(scrollregion=canvas.bbox("all"))
    frame_content.bind("<Configure>", on_frame_configure)

    # --- Frames PF e PJ (Dentro da Área Rolável) ---
    frame_pf = tk.LabelFrame(frame_content, text=f"PESSOA FÍSICA", padx=45, pady=5)
    frame_pf.pack(side="left", fill="none", expand=True, padx=10, pady=45, anchor="nw") # Anchor pode ajudar
    frame_pj = tk.LabelFrame(frame_content, text=f"PESSOA JURÍDICA", padx=45, pady=5)
    frame_pj.pack(side="right", fill="none", expand=True, padx=10, pady=45, anchor="ne") # Anchor pode ajudar

    # --- Populando Checkboxes ---
    checkboxes_pf.clear(); checkboxes_pj.clear() # Limpa listas caso recarregue
    print(f"\nPopulando Interface (Ignorando Status='GERADO')")
    count_pf_gerado, count_pj_gerado = 0, 0
    for pessoa in dados_pf:
        if str(pessoa.get(STATUS_COL, "")).strip().upper() != "GERADO":
            adicionar_checkbox(pessoa, frame_pf, checkboxes_pf)
        else: count_pf_gerado += 1
    for pessoa in dados_pj:
        if str(pessoa.get(STATUS_COL, "")).strip().upper() != "GERADO":
            adicionar_checkbox(pessoa, frame_pj, checkboxes_pj)
        else: count_pj_gerado += 1
    print(f" PF: {len(checkboxes_pf)} para exibir ({count_pf_gerado} já 'GERADO').")
    print(f" PJ: {len(checkboxes_pj)} para exibir ({count_pj_gerado} já 'GERADO').")

    # --- Botões de Seleção (Fixos na parte de baixo) ---
    frame_botoes_selecao = tk.Frame(root_window)
    frame_botoes_selecao.pack(pady=(0,5), fill="x", side=tk.BOTTOM) # (top, bottom) padding
    def selecionar(lista): [var.set(True) for _, var in lista]
    def desmarcar(lista): [var.set(False) for _, var in lista]
    tk.Button(frame_botoes_selecao, text="Selecionar todos - PF", command=lambda: selecionar(checkboxes_pf)).pack(side=tk.LEFT, padx=5)
    tk.Button(frame_botoes_selecao, text="Desmarcar todos - PF", command=lambda: desmarcar(checkboxes_pf)).pack(side=tk.LEFT, padx=5)
    # Espaçador simples
    tk.Label(frame_botoes_selecao, text="   ").pack(side=tk.LEFT)
    tk.Button(frame_botoes_selecao, text="Selecionar todos - PJ", command=lambda: selecionar(checkboxes_pj)).pack(side=tk.LEFT, padx=5)
    tk.Button(frame_botoes_selecao, text="Desmarcar todos - PJ", command=lambda: desmarcar(checkboxes_pj)).pack(side=tk.LEFT, padx=5)


    # --- Botões Principais (Fixos, acima dos de seleção) ---
    frame_botoes_principais = tk.Frame(root_window)
    frame_botoes_principais.pack(pady=5, fill="x", side=tk.BOTTOM)
    tk.Button(frame_botoes_principais, text="Gerar Documentos", command=gerar_documentos_cmd, width=20).pack(side=tk.BOTTOM, padx=10, expand=False, fill='none')
    # tk.Button(frame_botoes_principais, text="Excluir da Planilha", command=excluir_entradas_cmd, fg="red", width=20).pack(side=tk.LEFT, padx=20, expand=False, fill='x')

    # --- Label (Fixo, no fundo) ---
    label_dev = tk.Label(root_window, text='Desenvolvido por Lucas Costa')
    label_dev.pack(pady=(0,5), side=tk.BOTTOM)


# ==============================================================================
# 8. EXECUÇÃO PRINCIPAL
# ==============================================================================
if __name__ == "__main__":
    try:
        # Autentica e carrega dados
        client_gspread = autenticar_google()
        sheet_pf, headers_pf, dados_pf_originais = carregar_planilha(client_gspread, PLANILHA_PF_FILENAME, PF_TAB_NAME)
        sheet_pj, headers_pj, dados_pj_originais = carregar_planilha(client_gspread, PLANILHA_PJ_FILENAME, PJ_TAB_NAME)

        # Calcula índices das colunas de status uma vez
        try: col_index_status_pf = headers_pf.index(STATUS_COL) + 1
        except ValueError: messagebox.showerror("Erro Fatal", f"Coluna Status '{STATUS_COL}' não encontrada em PF."); sys.exit()
        try: col_index_status_pj = headers_pj.index(STATUS_COL) + 1
        except ValueError: messagebox.showerror("Erro Fatal", f"Coluna Status '{STATUS_COL}' não encontrada em PJ."); sys.exit()


        # Executa pré-preenchimento
        dados_pf_preenchidos = preencher_e_atualizar_planilha(
            sheet_pf, headers_pf, dados_pf_originais,
            COL_PF_ID_TRIGGER, COL_PF_ID_COMPARISON, COL_CADASTRO, TRIGGER_VALUE
        )
        dados_pj_preenchidos = preencher_e_atualizar_planilha(
            sheet_pj, headers_pj, dados_pj_originais,
            COL_PJ_ID_TRIGGER, COL_PJ_ID_COMPARISON, COL_CADASTRO, TRIGGER_VALUE
        )

        # Processa dados (adiciona tipo/linha) - usa os dados retornados pelo pré-preenchimento
        dados_pf_processados = [dict(row, tipo="PF", linha=i+2) for i, row in enumerate(dados_pf_preenchidos)]
        dados_pj_processados = [dict(row, tipo="PJ", linha=i+2) for i, row in enumerate(dados_pj_preenchidos)]

        # Cria e executa a interface gráfica
        root = tk.Tk()
        criar_interface(root, dados_pf_processados, dados_pj_processados)
        print("\nInterface pronta. Executando mainloop...")
        root.mainloop()

    except Exception as main_error:
        # Captura erros gerais que podem ocorrer antes da interface iniciar
        print(f"ERRO FATAL NA EXECUÇÃO PRINCIPAL: {main_error}")
        # Tenta mostrar um messagebox, mas pode falhar se o Tkinter não iniciou
        try:
            messagebox.showerror("Erro Fatal Inesperado", f"Ocorreu um erro crítico:\n{type(main_error).__name__}: {main_error}\n\nVerifique o console para detalhes.")
        except Exception:
             pass # Ignora erro ao mostrar messagebox se o Tkinter falhou
        sys.exit(1) # Sai com código de erro

    print("Aplicação finalizada.")