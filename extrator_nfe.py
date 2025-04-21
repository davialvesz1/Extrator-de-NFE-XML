import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import zipfile
import io

st.set_page_config(page_title="Extrator de NFe", layout="centered")

ns = {'ns': 'http://www.portalfiscal.inf.br/nfe'}

def extrair_dados_nfe(xml_content):
    tree = ET.ElementTree(ET.fromstring(xml_content))
    root = tree.getroot()

    emit = root.find('.//ns:emit', ns)
    dest = root.find('.//ns:dest', ns)
    total = root.find('.//ns:total/ns:ICMSTot', ns)
    ide = root.find('.//ns:ide', ns)

    dados = {
        'nNF': ide.findtext('ns:nNF', default='', namespaces=ns),
        'data_emissao': ide.findtext('ns:dhEmi', default='', namespaces=ns)[:10],
        'emitente_nome': emit.findtext('ns:xNome', default='', namespaces=ns),
        'destinatario_nome': dest.findtext('ns:xNome', default='', namespaces=ns),
        'valor_total': total.findtext('ns:vNF', default='', namespaces=ns),
    }

    produtos = []
    for det in root.findall('.//ns:det', ns):
        prod = det.find('ns:prod', ns)
        produtos.append({
            'nNF': dados['nNF'],
            'descricao': prod.findtext('ns:xProd', default='', namespaces=ns),
            'quantidade': prod.findtext('ns:qCom', default='', namespaces=ns),
            'valor_unitario': prod.findtext('ns:vUnCom', default='', namespaces=ns),
            'valor_total': prod.findtext('ns:vProd', default='', namespaces=ns)
        })

    return dados, produtos

st.title("📄 Extrator de NFe em XML")

uploaded_files = st.file_uploader("Envie arquivos XML de NFe", type="xml", accept_multiple_files=True)

if uploaded_files:
    lista_nfe = []
    lista_produtos = []
    arquivos_zip = io.BytesIO()
    with zipfile.ZipFile(arquivos_zip, mode="w") as zf:
        for file in uploaded_files:
            content = file.read()
            try:
                dados, produtos = extrair_dados_nfe(content)
                lista_nfe.append(dados)
                lista_produtos.extend(produtos)
                zf.writestr(file.name, content)
            except Exception as e:
                st.error(f"Erro ao processar {file.name}: {e}")

    df_nfe = pd.DataFrame(lista_nfe)
    df_prod = pd.DataFrame(lista_produtos)

    st.success("Notas processadas com sucesso!")

    with st.expander("🔍 Visualizar dados extraídos"):
        st.subheader("Notas Fiscais")
        st.dataframe(df_nfe)

        st.subheader("Produtos")
        st.dataframe(df_prod)

    # Criar Excel para download
    output_excel = io.BytesIO()
    with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
        df_nfe.to_excel(writer, sheet_name='Notas', index=False)
        df_prod.to_excel(writer, sheet_name='Produtos', index=False)
    output_excel.seek(0)

    st.download_button("📥 Baixar Excel", output_excel, file_name="notas_fiscais.xlsx")

    arquivos_zip.seek(0)
    st.download_button("📥 Baixar XMLs ZIP", arquivos_zip, file_name="xmls.zip")
