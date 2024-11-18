# Pasta onde estão os arquivos XML
pasta_xml = r'C:\Users\yasmin.thums\Desktop\Leitor de xml\XMLs'

# Listas para armazenar os dados de todos os XMLs
dados_nfse = []
dados_nfe = []
dados_cte = []

def extract_elements(d, parent_key='', sep='_'):
    """ Extrai todos os elementos de forma recursiva do dicionário XML """
    items = {}
    if isinstance(d, dict):
        for k, v in d.items():
            new_key = f"{parent_key}{sep}{k}" if parent_key else k
            items.update(extract_elements(v, new_key, sep))
    elif isinstance(d, list):
        for i, v in enumerate(d):
            new_key = f"{parent_key}{sep}{i}" if parent_key else str(i)
            items.update(extract_elements(v, new_key, sep))
    else:
        items[parent_key] = d
    return items

# Itera sobre os arquivos na pasta
for filename in os.listdir(pasta_xml):
    if filename.endswith('.xml'):
        print(filename)
        file_path = os.path.join(pasta_xml, filename)
        with open(file_path, 'r', encoding='utf-8') as file:
            xml_content = file.read()
            obj = xmltodict.parse(xml_content)

            # Verifica se é NFse
            if 'ConsultarNfseResposta' in obj:
                try:
                    nfse_info = obj['ConsultarNfseResposta']['ListaNfse']['CompNfse']['Nfse']['InfNfse']
                    # Extrai dados da NFse e adiciona à lista
                    data_nfse = extract_elements(nfse_info)
                    dados_nfse.append(data_nfse)
                    print(f"Dados NFse encontrados em {filename}")
                except KeyError as e:
                    print(f"Chave NFse não encontrada em {filename}: {e}")

            # Verifica se é NFe
            if 'nfeProc' in obj:
                try:
                    nfe_info = obj['nfeProc']['NFe']['infNFe']
                    # Extrai dados da NFe e adiciona à lista
                    data_nfe = extract_elements(nfe_info)
                    dados_nfe.append(data_nfe)
                    print(f"Dados NFe encontrados em {filename}")
                except KeyError as e:
                    print(f"Chave NFe não encontrada em {filename}: {e}")

            # Verifica se é CTe
            if 'cteProc' in obj:
                try:
                    cte_info = obj['cteProc']['CTe']['infCte']
                    # Extrai dados do CTe e adiciona à lista
                    data_cte = extract_elements(cte_info)
                    dados_cte.append(data_cte)
                    print(f"Dados CTe encontrados em {filename}")
                except KeyError as e:
                    print(f"Chave CTe não encontrada em {filename}: {e}")

# Convertendo os dados para DataFrames
df_nfse = pd.DataFrame(dados_nfse)
df_nfe = pd.DataFrame(dados_nfe)
df_cte = pd.DataFrame(dados_cte)

# Exportando para Excel
with pd.ExcelWriter('dados_fiscais.xlsx') as writer:
    df_nfse.to_excel(writer, sheet_name='NFse', index=False)
    df_nfe.to_excel(writer, sheet_name='NFe', index=False)
    df_cte.to_excel(writer, sheet_name='CTe', index=False)

print("Dados exportados para Excel com sucesso!")
