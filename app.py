import streamlit as st
import openpyxl
from openpyxl import Workbook
import csv
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm
import os
import tempfile
import shutil
from pathlib import Path
import zipfile
import sys
import pandas as pd
import requests
from requests.auth import HTTPBasicAuth

st.set_page_config(page_title="Gerador de Relatórios SEPE", layout="wide")

st.title("🏗️ Gerador de Relatórios de Vistoria")
st.markdown("---")

# Adicionar tabs para escolher fonte de dados
tab1, tab2 = st.tabs(["📡 Conectar ao ODK Central", "📁 Upload de Arquivo CSV"])

with tab1:
    st.subheader("🔗 Conexão com ODK Central")

    col_odk1, col_odk2 = st.columns(2)

    with col_odk1:
        odk_url = st.text_input(
            "URL do ODK Central",
            value="https://levantamentos.dflegal.df.gov.br",
            help="URL base do servidor ODK Central"
        )

        odk_project_id = st.text_input(
            "ID do Projeto",
            value="4",
            help="ID numérico do projeto"
        )

        odk_form_id = st.text_input(
            "ID do Formulário",
            value="aWX8oXGiD9zmcpDX7KtAer",
            help="ID do formulário"
        )

    with col_odk2:
        odk_email = st.text_input(
            "Email",
            help="Email de usuário do ODK Central"
        )

        odk_password = st.text_input(
            "Senha",
            type="password",
            help="Senha do usuário"
        )

        baixar_anexos = st.checkbox(
            "Baixar anexos (imagens)",
            value=True,
            help="Baixa automaticamente as imagens enviadas no formulário"
        )

    if st.button("🔄 Conectar e Buscar Dados", type="primary", use_container_width=True):
        if not odk_email or not odk_password:
            st.error("❌ Por favor, preencha email e senha!")
        else:
            try:
                with st.spinner("Conectando ao ODK Central..."):
                    base_url = f"{odk_url}/v1/projects/{odk_project_id}/forms/{odk_form_id}"
                    auth = HTTPBasicAuth(odk_email, odk_password)

                    st.info("Buscando dados do formulário...")

                    csv_url = f"{base_url}/submissions.csv.zip"

                    response = requests.get(csv_url, auth=auth)
                    response.raise_for_status()

                    import io
                    from zipfile import ZipFile

                    zip_buffer = io.BytesIO(response.content)

                    with ZipFile(zip_buffer, 'r') as zip_file:
                        csv_filename = [f for f in zip_file.namelist() if f.endswith('.csv')][0]
                        csv_content = zip_file.read(csv_filename).decode('utf-8')

                    st.session_state['csv_data'] = csv_content
                    st.session_state['data_source'] = 'odk'
                    st.session_state['odk_credentials'] = {
                        'base_url': base_url,
                        'auth': auth
                    }

                    num_linhas = len(csv_content.split('\n')) - 1

                    # Gravar cópia do CSV em C:\temp\sepe
                    try:
                        pasta_sepe = r'C:\temp\sepe'
                        os.makedirs(pasta_sepe, exist_ok=True)
                        csv_local_path = os.path.join(pasta_sepe, 'dados_odk.csv')
                        with open(csv_local_path, 'w', encoding='utf-8') as f:
                            f.write(csv_content)
                        st.success(f"💾 Planilha salva em: {csv_local_path}")
                    except Exception as e:
                        st.warning(f"⚠️ Não foi possível salvar planilha em C:\\temp\\sepe: {e}")

                    if baixar_anexos:
                        st.info("Baixando anexos (imagens)...")

                        local_media_dir = None
                        if os.name == 'nt' or os.path.exists('C:/'):
                            try:
                                local_media_dir = 'C:/arquivos_sepe/media'
                                os.makedirs(local_media_dir, exist_ok=True)
                            except:
                                local_media_dir = None

                        temp_media_dir = os.path.join(tempfile.gettempdir(), 'odk_media')
                        os.makedirs(temp_media_dir, exist_ok=True)

                        # Pasta C:\temp\sepe\midia — destino principal das fotos
                        pasta_midia_sepe = r'C:\temp\sepe\midia'
                        try:
                            os.makedirs(pasta_midia_sepe, exist_ok=True)
                        except Exception as e:
                            st.warning(f"⚠️ Não foi possível criar C:\\temp\\sepe\\midia: {e}")
                            pasta_midia_sepe = None

                        try:
                            import io

                            csv_lines = csv_content.split('\n')
                            csv_reader_header = csv.reader([csv_lines[0]])
                            header_cols = next(csv_reader_header)

                            id_cols = [col for col in header_cols if
                                       'id' in col.lower() or 'ID' in col or 'mero' in col]

                            csv_reader = csv.DictReader(io.StringIO(csv_content))

                            id_map = {}
                            primeira_linha = None

                            for idx, row in enumerate(csv_reader):
                                if idx == 0:
                                    primeira_linha = row

                                instance_id_raw = row.get('KEY') or row.get('InstanceID') or row.get('meta-instanceID')

                                if not instance_id_raw:
                                    continue

                                id_projeto = row.get('details-N_mero_ID')

                                if not id_projeto or not str(id_projeto).strip():
                                    row_values = list(row.values())
                                    if len(row_values) > 4:
                                        id_projeto = row_values[4]

                                if id_projeto:
                                    id_projeto = str(id_projeto).strip()

                                if id_projeto:
                                    instance_ids = []
                                    instance_ids.append(instance_id_raw)

                                    if instance_id_raw.startswith('uuid:'):
                                        instance_ids.append(instance_id_raw[5:])

                                    if not instance_id_raw.startswith('uuid:'):
                                        instance_ids.append(f"uuid:{instance_id_raw}")

                                    for iid in instance_ids:
                                        id_map[iid] = id_projeto

                            num_registros = len(set(id_map.values()))
                            st.success(f"✅ {num_registros} registros mapeados com ID do projeto")

                            submissions_url = f"{base_url}/submissions"
                            submissions_response = requests.get(submissions_url, auth=auth)
                            submissions_response.raise_for_status()
                            submissions_data = submissions_response.json()

                            total_anexos = 0
                            anexos_baixados = []

                            for idx, submission in enumerate(submissions_data):
                                instance_id = submission.get('instanceId')
                                id_projeto = id_map.get(instance_id)

                                attachments_url = f"{base_url}/submissions/{instance_id}/attachments"
                                att_response = requests.get(attachments_url, auth=auth)

                                if att_response.status_code == 200:
                                    attachments = att_response.json()

                                    for attachment in attachments:
                                        att_name = attachment.get('name')

                                        att_download_url = f"{attachments_url}/{att_name}"
                                        file_response = requests.get(att_download_url, auth=auth)

                                        if file_response.status_code == 200:
                                            # Verificar que o conteúdo é realmente uma imagem
                                            content_type = file_response.headers.get('Content-Type', '')
                                            if 'image' not in content_type.lower():
                                                print(
                                                    f"Anexo ignorado (não é imagem): {att_name} - Content-Type: {content_type}")
                                                continue

                                            if id_projeto:
                                                novo_nome = f"foto_{id_projeto}_{att_name}"
                                            else:
                                                novo_nome = f"foto_{att_name}"

                                            # Salvar em C:\temp\sepe\midia (destino principal)
                                            if pasta_midia_sepe:
                                                try:
                                                    file_path_sepe = os.path.join(pasta_midia_sepe, att_name)
                                                    with open(file_path_sepe, 'wb') as f:
                                                        f.write(file_response.content)
                                                except Exception as e:
                                                    st.warning(
                                                        f"Erro ao salvar {att_name} em C:\\temp\\sepe\\midia: {e}")

                                            if local_media_dir:
                                                try:
                                                    file_path_local = os.path.join(local_media_dir, att_name)
                                                    with open(file_path_local, 'wb') as f:
                                                        f.write(file_response.content)
                                                except Exception as e:
                                                    st.warning(f"Erro ao salvar {att_name} localmente: {e}")

                                            file_path_temp = os.path.join(temp_media_dir, att_name)
                                            with open(file_path_temp, 'wb') as f:
                                                f.write(file_response.content)

                                            anexos_baixados.append({
                                                'nome_original': att_name,
                                                'nome_com_prefixo': novo_nome,
                                                'path_temp': file_path_temp,
                                                'data': file_response.content
                                            })

                                            total_anexos += 1

                            st.session_state['anexos_baixados'] = anexos_baixados

                            # Resumo dos destinos das fotos
                            msgs_destino = []
                            if pasta_midia_sepe and os.path.exists(pasta_midia_sepe):
                                n = len([f for f in os.listdir(pasta_midia_sepe)])
                                if n > 0:
                                    msgs_destino.append(f"C:\\temp\\sepe\\midia ({n} arquivos)")
                            if local_media_dir and os.path.exists(local_media_dir):
                                arquivos_salvos = os.listdir(local_media_dir)
                                if arquivos_salvos:
                                    msgs_destino.append(f"C:\\arquivos_sepe\\media ({len(arquivos_salvos)} arquivos)")

                            if msgs_destino:
                                st.success(f"✅ {total_anexos} fotos salvas em: {' | '.join(msgs_destino)}")
                            else:
                                st.success(f"✅ {total_anexos} fotos baixadas (somente em memória)")
                        except Exception as e:
                            st.warning(f"⚠️ Aviso ao baixar anexos: {str(e)}")

                    st.success(f"✅ Conectado com sucesso! {num_linhas} registros encontrados.")
                    st.rerun()

            except Exception as e:
                st.error(f"❌ Erro ao conectar: {str(e)}")
                st.exception(e)

with tab2:
    st.subheader("📄 Upload Manual de CSV")
    csv_file_upload = st.file_uploader("Selecione o arquivo CSV", type=['csv'])

    if csv_file_upload:
        st.session_state['csv_data'] = csv_file_upload.getvalue().decode('utf-8')
        st.session_state['data_source'] = 'upload'
        st.success("✅ Arquivo CSV carregado com sucesso!")

st.markdown("---")

# Verificar se há dados carregados (de qualquer fonte)
csv_file = None
if 'csv_data' in st.session_state:
    from io import StringIO, BytesIO

    csv_file = BytesIO(st.session_state['csv_data'].encode('utf-8'))

    fonte = "ODK Central" if st.session_state.get('data_source') == 'odk' else "Upload Manual"
    st.info(f"📊 Dados carregados de: **{fonte}**")

    if 'anexos_baixados' in st.session_state and len(st.session_state['anexos_baixados']) > 0:
        st.subheader("📥 Download de Imagens")

        col_img1, col_img2 = st.columns(2)

        with col_img1:
            st.info(f"**{len(st.session_state['anexos_baixados'])} imagens** disponíveis para download")

        with col_img2:
            if st.button("📦 Baixar Todas as Imagens (ZIP)", use_container_width=True):
                try:
                    zip_buffer = BytesIO()
                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                        for anexo in st.session_state['anexos_baixados']:
                            zip_file.writestr(anexo['nome_com_prefixo'], anexo['data'])

                    zip_buffer.seek(0)

                    st.download_button(
                        label="⬇️ Download ZIP com Imagens Renomeadas",
                        data=zip_buffer,
                        file_name="imagens_odk_com_prefixo_foto.zip",
                        mime="application/zip",
                        use_container_width=True
                    )

                    st.success("✅ ZIP criado com sucesso! As imagens foram renomeadas com prefixo 'foto_ID_'.")
                except Exception as e:
                    st.error(f"❌ Erro ao criar ZIP: {str(e)}")

        st.markdown("---")


# Criar diretórios temporários
@st.cache_resource
def criar_diretorios_temp():
    temp_dir = tempfile.mkdtemp()
    dirs = {
        'base': temp_dir,
        'xlsx': os.path.join(temp_dir, 'arquivo_xlsx'),
        'relatorios': os.path.join(temp_dir, 'relatorios_pdf'),
        'media': os.path.join(temp_dir, 'media'),
        'sem_media': os.path.join(temp_dir, 'sem_media'),
        'modelo': os.path.join(temp_dir, 'modelo_relatorio')
    }
    for d in dirs.values():
        os.makedirs(d, exist_ok=True)
    return dirs


def converter_csv_para_xlsx(csv_file, xlsx_path):
    """Converte CSV para XLSX com coluna de numeração"""
    wb = Workbook()
    ws = wb.active
    ws.title = 'dados_vistoria'

    csv_content = csv_file.getvalue().decode('utf-8').splitlines()
    csv_reader = csv.reader(csv_content, delimiter=',')

    for row_index, row in enumerate(csv_reader, start=1):
        ws.cell(row=row_index, column=1, value=row_index)
        for col_index, value in enumerate(row, start=2):
            ws.cell(row=row_index, column=col_index, value=value)

    wb.save(xlsx_path)
    return xlsx_path


def validar_imagem(path):
    """
    Valida se o arquivo no caminho dado é uma imagem real e legível.
    Retorna True se válida, False caso contrário.
    """
    try:
        from PIL import Image as PILImage
        with PILImage.open(path) as img:
            img.verify()
        return True
    except Exception as e:
        print(f"Arquivo não é imagem válida ({path}): {e}")
        return False


def processar_imagem(doc, valor_imagem, dirs):
    """
    Processa uma imagem para inserção no relatório DOCX.
    - Se valor_imagem é vazio/None: tenta usar imagem padrão.
    - Se valor_imagem tem nome de arquivo: tenta encontrá-lo nos caminhos conhecidos.
    - Valida cada candidato com PIL antes de passar ao python-docx.
    - Retorna None se nenhuma imagem válida for encontrada.
    """

    def fazer_inline(path, tamanho_cm):
        """Cria InlineImage apenas se o arquivo for uma imagem válida."""
        if not os.path.exists(path):
            return None
        if not validar_imagem(path):
            return None
        try:
            return InlineImage(doc, path, Cm(tamanho_cm))
        except Exception as e:
            print(f"Erro ao criar InlineImage ({path}): {e}")
            return None

    def obter_imagem_padrao():
        """
        Tenta obter uma imagem padrão 'sem imagem disponível':
        1. Caminho local (Windows)
        2. Download da internet (com validação de Content-Type)
        3. None se tudo falhar
        """
        # 1. Caminho local
        local = 'C:/arquivos_sepe/xxx.jpg'
        resultado = fazer_inline(local, 3)
        if resultado:
            return resultado

        # 2. Download da internet
        try:
            url = (
                "https://st2.depositphotos.com/12694644/47297/v/380/"
                "depositphotos_472972706-stock-illustration-image-available-sign-isolated-white.jpg"
            )
            temp_path = os.path.join(tempfile.gettempdir(), 'no_image_default.jpg')

            if not os.path.exists(temp_path):
                r = requests.get(url, timeout=10)
                r.raise_for_status()

                content_type = r.headers.get('Content-Type', '')
                if 'image' not in content_type.lower():
                    raise ValueError(
                        f"URL da imagem padrão retornou Content-Type inválido: {content_type}"
                    )

                with open(temp_path, 'wb') as f:
                    f.write(r.content)

            resultado = fazer_inline(temp_path, 3)
            if resultado:
                return resultado

        except Exception as e:
            print(f"Não foi possível obter imagem padrão da internet: {e}")

        return None  # Sem imagem disponível

    # --- Sem nome de arquivo: usar imagem padrão ---
    if not valor_imagem or str(valor_imagem).strip() == '':
        return obter_imagem_padrao()

    # --- Com nome de arquivo: procurar nos caminhos possíveis ---
    nome = str(valor_imagem).strip()
    caminhos_possiveis = [
        rf'C:\temp\sepe\midia\{nome}',
        f'C:/arquivos_sepe/media/{nome}',
        os.path.join(tempfile.gettempdir(), 'odk_media', nome),
        os.path.join(dirs.get('media', ''), nome),
    ]

    for path in caminhos_possiveis:
        resultado = fazer_inline(path, 12)
        if resultado:
            return resultado

    # Imagem referenciada não encontrada ou inválida → usar padrão
    print(f"Imagem '{nome}' não encontrada ou inválida em nenhum caminho. Usando imagem padrão.")
    return obter_imagem_padrao()


def processar_relatorios(xlsx_path, modelo_path, dirs, indices_selecionados=None):
    """Processa e gera os relatórios em DOCX"""

    workbook = openpyxl.load_workbook(xlsx_path)
    sheet = workbook['dados_vistoria']
    list_values = list(sheet.values)

    relatorios_gerados = []
    progress_bar = st.progress(0)
    status_text = st.empty()

    if indices_selecionados:
        dados_filtrados = [list_values[0]] + [list_values[i] for i in indices_selecionados if i < len(list_values)]
    else:
        dados_filtrados = list_values

    total = len(dados_filtrados[1:])

    for idx, valores in enumerate(dados_filtrados[1:], 1):
        status_text.text(f"Processando relatório {idx} de {total}: {valores[0]}")
        progress_bar.progress(idx / total)

        doc = DocxTemplate(modelo_path)

        # Processar imagens com validação
        imagem1 = processar_imagem(doc, valores[18], dirs)
        imagem2 = processar_imagem(doc, valores[19], dirs)
        imagem3 = processar_imagem(doc, valores[20], dirs)
        imagem4 = processar_imagem(doc, valores[21], dirs)
        imagem5 = processar_imagem(doc, valores[22], dirs)

        # Formatar data
        data_formatada = valores[2]
        if valores[2] and isinstance(valores[2], str):
            try:
                from datetime import datetime
                if 'T' in valores[2]:
                    dt = datetime.fromisoformat(valores[2].replace('Z', '+00:00'))
                else:
                    dt = datetime.strptime(valores[2], '%Y-%m-%d')
                data_formatada = dt.strftime('%d-%m-%Y')
            except:
                data_formatada = valores[2]

        doc.render({
            'possui_placa': valores[9],
            'plano_trabalho': valores[10],
            'relatorio': valores[0],
            'id_proj': valores[5],
            'id_tipo_rel': valores[8],
            'meta': valores[25],
            'data': data_formatada,
            'processo_sei': valores[6],
            'cidade': valores[12],
            'responsavel': valores[27],
            'lat': valores[13],
            'long': valores[14],
            'observacao': valores[23],
            'tipo_proj': valores[17],
            'imagem_1': imagem1,
            'imagem_2': imagem2,
            'imagem_3': imagem3,
            'imagem_4': imagem4,
            'imagem_5': imagem5
        })

        doc_name = os.path.join(dirs['relatorios'], f"{valores[0]}.docx")
        doc.save(doc_name)
        relatorios_gerados.append(doc_name)

    progress_bar.empty()
    status_text.empty()

    return relatorios_gerados


def criar_zip(arquivos, zip_path):
    """Cria um arquivo ZIP com os relatórios"""
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for arquivo in arquivos:
            zipf.write(arquivo, os.path.basename(arquivo))


# Interface principal
dirs = criar_diretorios_temp()

col1, col2 = st.columns(2)

with col1:
    st.subheader("📄 Modelo do Relatório")
    modelo_file = st.file_uploader("Upload do modelo DOCX (formulario.docx)", type=['docx'])

with col2:
    st.subheader("📁 Diretórios de Imagens")

    local_exists = os.path.exists('C:/arquivos_sepe')

    if local_exists:
        st.info("**Imagem padrão:** `C:/arquivos_sepe/xxx.jpg`")
        st.info("**Imagens do projeto:** `C:/arquivos_sepe/media/`")

        if os.path.exists('C:/arquivos_sepe/xxx.jpg'):
            st.success("✅ Imagem padrão local encontrada")
        else:
            st.info("ℹ️ Usando imagem padrão da internet")

        if os.path.exists('C:/arquivos_sepe/media'):
            num_imagens = len(
                [f for f in os.listdir('C:/arquivos_sepe/media') if f.lower().endswith(('.jpg', '.jpeg', '.png'))])
            st.success(f"✅ Diretório de imagens encontrado ({num_imagens} imagens)")
        else:
            st.warning("⚠️ Diretório de imagens não encontrado")
    else:
        st.info("**🌐 Modo Cloud**")
        st.success("✅ Imagem padrão: Internet (depositphotos)")
        st.success("✅ Imagens do projeto: Download do ODK")
        st.caption("Marque '✓ Baixar anexos' ao conectar ao ODK")

st.markdown("---")

# Seção de preview e seleção de relatórios
if csv_file is not None:
    st.subheader("📋 Visualizar e Selecionar Relatórios")

    csv_file.seek(0)
    csv_text = csv_file.read().decode('utf-8')

    lines = csv_text.strip().split('\n')

    if len(lines) > 1:
        import csv

        header_reader = csv.reader([lines[0]])
        original_cols = next(header_reader)

        seen = {}
        unique_cols = []

        for col in original_cols:
            col = col.strip()
            if col not in seen:
                seen[col] = 0
                unique_cols.append(col)
            else:
                seen[col] += 1
                unique_cols.append(f"{col}_dup{seen[col]}")

        data_reader = csv.reader(lines[1:])
        data_rows = [row for row in data_reader if row]

        df = pd.DataFrame(data_rows, columns=unique_cols)

        final_cols = []
        for i, col in enumerate(df.columns):
            if df.columns.tolist().count(col) > 1:
                final_cols.append(f"{col}_idx{i}")
            else:
                final_cols.append(col)

        df.columns = final_cols

        df.insert(0, '#', range(1, len(df) + 1))

        header = df.columns.tolist()

        if len(header) != len(set(header)):
            st.error(f"🔴 AINDA HÁ DUPLICATAS: {[h for h in header if header.count(h) > 1]}")
            st.stop()

        submission_date_cols = [col for col in df.columns if
                                'SubmissionDate' in col and not col.endswith(tuple('0123456789'))]
        if submission_date_cols:
            try:
                col_name = submission_date_cols[0]
                df[col_name] = pd.to_datetime(df[col_name], errors='coerce')
                df[col_name] = df[col_name].dt.strftime('%d-%m-%Y')
            except:
                pass

    if len(df) > 0:
        st.info(f"📊 Total de relatórios disponíveis: **{len(df)}**")

        col_sel1, col_sel2 = st.columns([1, 3])

        with col_sel1:
            selecao_tipo = st.radio(
                "Tipo de seleção:",
                ["Todos os relatórios", "Selecionar específicos"],
                key="tipo_selecao"
            )

        with col_sel2:
            # Função auxiliar para montar lista de colunas para exibição
            def montar_colunas_display(header, df, incluir_tipo_proj=False):
                colunas = ['#']

                id_proj_cols = [c for c in header if 'N_mero_ID' in c or 'Numero_ID' in c or 'details-N' in c]
                if id_proj_cols:
                    colunas.append(id_proj_cols[0])
                elif len(header) > 1:
                    colunas.append(header[1])

                tipo_relat_cols = [c for c in header if 'Tipo_Relat' in c or 'Tipo_Relatorio' in c]
                if tipo_relat_cols:
                    colunas.append(tipo_relat_cols[0])

                submission_cols = [c for c in header if 'SubmissionDate' in c]
                if submission_cols:
                    colunas.append(submission_cols[0])
                elif len(header) > 3:
                    colunas.append(header[3])

                cidade_cols = [c for c in header if 'cidade' in c.lower() or 'regiao' in c.lower()]
                if cidade_cols:
                    colunas.append(cidade_cols[0])
                elif len(header) > 7:
                    colunas.append(header[7])

                processo_cols = [c for c in header if 'processo' in c.lower() or 'sei' in c.lower()]
                if processo_cols:
                    colunas.append(processo_cols[0])
                elif len(header) > 6:
                    colunas.append(header[6])

                if incluir_tipo_proj:
                    tipo_proj_cols = [c for c in header if 'tipo' in c.lower() and 'proj' in c.lower()]
                    if tipo_proj_cols:
                        colunas.append(tipo_proj_cols[0])
                    elif len(header) > 12:
                        colunas.append(header[12])

                # Remover duplicatas preservando ordem e garantindo que existem no df
                vistos = []
                for c in colunas:
                    if c not in vistos and c in df.columns:
                        vistos.append(c)
                return vistos


            if selecao_tipo == "Selecionar específicos":
                colunas_display_unique = montar_colunas_display(header, df)
                df_display = df[colunas_display_unique].copy().reset_index(drop=True)
                st.dataframe(df_display, width="stretch", height=400)

                numeros_selecionados = st.text_input(
                    "Digite os números dos relatórios (separados por vírgula):",
                    placeholder="Ex: 1, 3, 5, 7-10",
                    help="Você pode usar vírgulas para separar números individuais ou hífen para intervalos"
                )

                indices_selecionados = []
                if numeros_selecionados:
                    try:
                        partes = numeros_selecionados.split(',')
                        for parte in partes:
                            parte = parte.strip()
                            if '-' in parte:
                                inicio, fim = map(int, parte.split('-'))
                                indices_selecionados.extend(range(inicio, fim + 1))
                            else:
                                indices_selecionados.append(int(parte))

                        indices_selecionados = sorted(set(indices_selecionados))
                        st.success(
                            f"✅ {len(indices_selecionados)} relatórios selecionados: {', '.join(map(str, indices_selecionados))}")
                    except:
                        st.error("❌ Formato inválido. Use números separados por vírgula ou intervalos com hífen.")
            else:
                colunas_display_unique = montar_colunas_display(header, df, incluir_tipo_proj=True)
                df_display = df[colunas_display_unique].copy().reset_index(drop=True)
                st.dataframe(df_display, width="stretch", height=400)
                st.caption(f"📊 Mostrando todos os {len(df)} relatórios")
                indices_selecionados = list(range(1, len(df) + 1))
    else:
        st.warning("⚠️ O arquivo CSV está vazio.")
        indices_selecionados = []
else:
    indices_selecionados = []

st.markdown("---")

# Botão de gerar com validação de seleção
botao_habilitado = csv_file is not None and modelo_file is not None and len(indices_selecionados) > 0

if not botao_habilitado and csv_file is not None and modelo_file is not None:
    st.warning("⚠️ Nenhum relatório selecionado. Por favor, selecione ao menos um relatório.")

if st.button("🚀 Gerar Relatórios", type="primary", use_container_width=True, disabled=not botao_habilitado):

    if not csv_file:
        st.error("❌ Por favor, faça upload do arquivo CSV ou conecte ao ODK Central!")
    elif not modelo_file:
        st.error("❌ Por favor, faça upload do modelo DOCX!")
    else:
        try:
            with st.spinner("Processando..."):

                modelo_path = os.path.join(dirs['modelo'], 'formulario.docx')
                with open(modelo_path, 'wb') as f:
                    f.write(modelo_file.getbuffer())

                st.info("Convertendo CSV para XLSX...")
                xlsx_path = os.path.join(dirs['xlsx'], 'dados.xlsx')
                converter_csv_para_xlsx(csv_file, xlsx_path)

                st.info("Gerando relatórios...")
                relatorios = processar_relatorios(xlsx_path, modelo_path, dirs, indices_selecionados)

                zip_path = os.path.join(dirs['base'], 'relatorios.zip')
                criar_zip(relatorios, zip_path)

                st.success(f"✅ {len(relatorios)} relatórios gerados com sucesso!")

                with open(zip_path, 'rb') as f:
                    st.download_button(
                        label="📥 Download de Todos os Relatórios DOCX (ZIP)",
                        data=f,
                        file_name="relatorios_vistoria.zip",
                        mime="application/zip",
                        use_container_width=True
                    )

        except Exception as e:
            st.error(f"❌ Erro ao processar: {str(e)}")
            st.exception(e)

st.markdown("---")
st.caption("Desenvolvido para SEPE - Sistema de Geração de Relatórios de Vistoria - versão 1.4")
