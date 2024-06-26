import os
import xmltodict
import pandas as pd

# Pasta onde estão os arquivos XML
pasta_xml = r'C:\Users\yasmin.thums\Desktop\Leitor de xml\XMLs'

# Lista para armazenar os dados de todos os XMLs
dados = []

# Itera sobre os arquivos na pasta
for filename in os.listdir(pasta_xml):
    if filename.endswith('.xml'):
        file_path = os.path.join(pasta_xml, filename)
        with open(file_path, 'r', encoding='utf-8') as file:
            xml_content = file.read()
            obj = xmltodict.parse(xml_content)
            print(obj)
            try:
                nfse_info = obj['ConsultarNfseResposta']['ListaNfse']['CompNfse']['Nfse']['InfNfse']
                data = {
                    'codigo_verificacao': nfse_info.get('CodigoVerificacao', ''),
                    'competencia': nfse_info.get('Competencia', ''),
                    'data_emissao': nfse_info.get('DataEmissao', ''),
                    'incentivador_cultural': nfse_info.get('IncentivadorCultural', ''),
                    'natureza_operacao': nfse_info.get('NaturezaOperacao', ''),
                    'numero': nfse_info.get('Numero', ''),
                    'optante_simples_nacional': nfse_info.get('OptanteSimplesNacional', ''),
                    'orgao_gerador_codigo_municipio': nfse_info.get('OrgaoGerador', {}).get('CodigoMunicipio', ''),
                    'orgao_gerador_uf': nfse_info.get('OrgaoGerador', {}).get('Uf', ''),
                    'prestador_contato_email': nfse_info.get('PrestadorServico', {}).get('Contato', {}).get('Email', ''),
                    'prestador_contato_telefone': nfse_info.get('PrestadorServico', {}).get('Contato', {}).get('Telefone', ''),
                    'prestador_endereco_bairro': nfse_info.get('PrestadorServico', {}).get('Endereco', {}).get('Bairro', ''),
                    'prestador_endereco_cep': nfse_info.get('PrestadorServico', {}).get('Endereco', {}).get('Cep', ''),
                    'prestador_endereco_codigo_municipio': nfse_info.get('PrestadorServico', {}).get('Endereco', {}).get('CodigoMunicipio', ''),
                    'prestador_endereco_complemento': nfse_info.get('PrestadorServico', {}).get('Endereco', {}).get('Complemento', ''),
                    'prestador_endereco_endereco': nfse_info.get('PrestadorServico', {}).get('Endereco', {}).get('Endereco', ''),
                    'prestador_endereco_numero': nfse_info.get('PrestadorServico', {}).get('Endereco', {}).get('Numero', ''),
                    'prestador_endereco_uf': nfse_info.get('PrestadorServico', {}).get('Endereco', {}).get('Uf', ''),
                    'prestador_cnpj': nfse_info.get('PrestadorServico', {}).get('IdentificacaoPrestador', {}).get('Cnpj', ''),
                    'prestador_inscricao_municipal': nfse_info.get('PrestadorServico', {}).get('IdentificacaoPrestador', {}).get('InscricaoMunicipal', ''),
                    'prestador_razao_social': nfse_info.get('PrestadorServico', {}).get('RazaoSocial', ''),
                    'servico_codigo_municipio': nfse_info.get('Servico', {}).get('CodigoMunicipio', ''),
                    'servico_codigo_tributacao_municipio': nfse_info.get('Servico', {}).get('CodigoTributacaoMunicipio', ''),
                    'servico_discriminacao': nfse_info.get('Servico', {}).get('Discriminacao', ''),
                    'servico_item_lista_servico': nfse_info.get('Servico', {}).get('ItemListaServico', ''),
                    'servico_valores_aliquota': nfse_info.get('Servico', {}).get('Valores', {}).get('Aliquota', ''),
                    'servico_valores_base_calculo': nfse_info.get('Servico', {}).get('Valores', {}).get('BaseCalculo', ''),
                    'servico_valores_iss_retido': nfse_info.get('Servico', {}).get('Valores', {}).get('IssRetido', ''),
                    'servico_valores_inss': nfse_info.get('Servico', {}).get('Valores', {}).get('ValorInss', ''),
                    'servico_valores_iss': nfse_info.get('Servico', {}).get('Valores', {}).get('ValorIss', ''),
                    'servico_valores_liquido_nfse': nfse_info.get('Servico', {}).get('Valores', {}).get('ValorLiquidoNfse', ''),
                    'servico_valores_servicos': nfse_info.get('Servico', {}).get('Valores', {}).get('ValorServicos', ''),
                    'tomador_contato_email': nfse_info.get('TomadorServico', {}).get('Contato', {}).get('Email', ''),
                    'tomador_contato_telefone': nfse_info.get('TomadorServico', {}).get('Contato', {}).get('Telefone', ''),
                    'tomador_endereco_bairro': nfse_info.get('TomadorServico', {}).get('Endereco', {}).get('Bairro', ''),
                    'tomador_endereco_cep': nfse_info.get('TomadorServico', {}).get('Endereco', {}).get('Cep', ''),
                    'tomador_endereco_codigo_municipio': nfse_info.get('TomadorServico', {}).get('Endereco', {}).get('CodigoMunicipio', ''),
                    'tomador_endereco_complemento': nfse_info.get('TomadorServico', {}).get('Endereco', {}).get('Complemento', ''),
                    'tomador_endereco_endereco': nfse_info.get('TomadorServico', {}).get('Endereco', {}).get('Endereco', ''),
                    'tomador_endereco_numero': nfse_info.get('TomadorServico', {}).get('Endereco', {}).get('Numero', ''),
                    'tomador_endereco_uf': nfse_info.get('TomadorServico', {}).get('Endereco', {}).get('Uf', ''),
                    'tomador_cnpj': nfse_info.get('TomadorServico', {}).get('IdentificacaoTomador', {}).get('CpfCnpj', {}).get('Cnpj', ''),
                    'tomador_inscricao_municipal': nfse_info.get('TomadorServico', {}).get('IdentificacaoTomador', {}).get('InscricaoMunicipal', ''),
                    'tomador_razao_social': nfse_info.get('TomadorServico', {}).get('RazaoSocial', '')
                }
                dados.append(data)
            except KeyError as e:
                print(f"Chave não encontrada em {filename}: {e}")
                continue

# Criar um DataFrame pandas com os dados coletados
df = pd.DataFrame(dados)

# Exportar o DataFrame para um arquivo Excel
excel_file = r'C:\Users\yasmin.thums\Desktop\Leitor de xml\dados_nfse.xlsx'
df.to_excel(excel_file, index=False)

print(f'Dados exportados para {excel_file}')
