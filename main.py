import pdfplumber
import pandas as pd
import re
import glob
import os
from datetime import datetime

# Dicionário para conversão de meses do Nubank
MESES_PT = {
    'JAN': 1, 'FEV': 2, 'MAR': 3, 'ABR': 4, 'MAI': 5, 'JUN': 6,
    'JUL': 7, 'AGO': 8, 'SET': 9, 'OUT': 10, 'NOV': 11, 'DEZ': 12
}

def classificar_fatura(texto):
    """Identifica de qual banco é a fatura baseado no texto da primeira página."""
    texto_upper = texto.upper()
    if "NUBANK" in texto_upper or "OLÁ, DANTON.\nESTA É A SUA FATURA DE" in texto_upper or "OLÁ, LAUREN.\nESTA É A SUA FATURA DE" in texto_upper or "NU PAGAMENTOS" in texto_upper:
        return "Nubank"
    elif "PICPAY" in texto_upper:
        return "PicPay"
    return "Desconhecido"

def extrair_ano_mes_fatura(referencia):
    """Extrai o ano e o mês baseados no nome do arquivo."""
    match_ano = re.search(r'(20\d{2})', referencia)
    ano = int(match_ano.group(1)) if match_ano else datetime.now().year
    
    match_mes_nu = re.search(r'-(\d{2})-', referencia)
    match_mes_pp = re.search(r'_(\d{2})20\d{2}', referencia)
    
    if match_mes_nu:
        mes = int(match_mes_nu.group(1))
    elif match_mes_pp:
        mes = int(match_mes_pp.group(1))
    else:
        mes = datetime.now().month
        
    return ano, mes

def formatar_data(data_str, ano_fatura, mes_fatura):
    """Converte a data extraída para o formato MM/DD/YYYY americano."""
    match_nu = re.match(r'(\d{2})\s+([A-Z]{3})', data_str.upper())
    if match_nu:
        dia = int(match_nu.group(1))
        mes = MESES_PT.get(match_nu.group(2), mes_fatura)
    else:
        match_pp = re.match(r'(\d{2})/(\d{2})', data_str)
        if match_pp:
            dia = int(match_pp.group(1))
            mes = int(match_pp.group(2))
        else:
            return data_str

    ano = ano_fatura
    if mes == 12 and mes_fatura == 1:
        ano -= 1
    elif mes == 1 and mes_fatura == 12:
        ano += 1
        
    return f"{mes:02d}/{dia:02d}/{ano}"

def formatar_valor(valor_str):
    """
    Converte o valor para o padrão americano (ex: -1,261.90),
    garantindo que o sinal negativo seja preservado.
    """

    is_negative = '-' in valor_str
    
    if is_negative:
        print(f"Valor original: {valor_str} - Sinal negativo detectado.") 
    
    v = valor_str.replace('R$', '').replace('-', '').strip()
    
    # Remove os pontos de milhar brasileiros e troca vírgula por ponto
    v = v.replace('.', '').replace(',', '.')
    
    try:
        f_val = float(v)
        if is_negative == False:
            f_val = -f_val
            
        # Formata explicitamente para string americana
        return "{:,.2f}".format(f_val)
    except Exception:
        return valor_str

def extrair_transacoes(caminho_pdf):
    transacoes = []
    
    padrao_nubank = re.compile(r'^(\d{2}\s[A-Z]{3})\s+(.+?)\s+(-?R\$\s?[\d\.]+,\d{2})$')
    padrao_picpay = re.compile(r'^(\d{2}/\d{2})\s+(.+?)\s+(-?[\d\.]+,\d{2})$')

    with pdfplumber.open(caminho_pdf) as pdf:
        primeira_pagina = pdf.pages[0].extract_text()
        banco = classificar_fatura(primeira_pagina)
        
        referencia = os.path.basename(caminho_pdf).replace('.pdf', '')
        ano_fatura, mes_fatura = extrair_ano_mes_fatura(referencia)

        for pagina in pdf.pages:
            texto = pagina.extract_text()
            if not texto:
                continue
                
            linhas = texto.split('\n')
            
            for linha in linhas:
                linha = linha.strip()
                
                linha = linha.replace('−', '-').replace('–', '-').replace('—', '-')
                
                if banco == "Nubank":
                    linha_limpa = re.sub(r'^(\d{2}\s[A-Z]{3})\s+\d{4}\s+', r'\1 ', linha)
                    match = padrao_nubank.search(linha_limpa)
                    
                    if match:
                        data, descricao, valor = match.groups()
                        if "PAGAMENTO RECEBIDO" not in descricao.upper() and "TOTAL" not in descricao.upper():
                            transacoes.append({
                                "Banco/Cartão": banco,
                                "Fatura Referência": referencia,
                                "Data": formatar_data(data.strip(), ano_fatura, mes_fatura),
                                "Estabelecimento / Descrição": descricao.strip(),
                                "Valor": formatar_valor(valor),
                                "Motivo do Estorno": "" # Inicia vazio
                            })
                    else:
                        if transacoes:
                            if "Estorno referente" in linha:
                                motivo_limpo = re.sub(r'^.*?Estorno referente', 'Estorno referente', linha)
                                transacoes[-1]["Motivo do Estorno"] = motivo_limpo.strip()
                            
                            elif transacoes[-1]["Motivo do Estorno"] != "":
                                if not re.match(r'^(\d{2}\s[A-Z]{3}|\d{2}/\d{2}|Página|Pagamento)', linha, re.IGNORECASE) and linha != "":
                                    transacoes[-1]["Motivo do Estorno"] += " " + linha.strip()
                            
                elif banco == "PicPay":
                    match = padrao_picpay.search(linha)
                    if match:
                        data, descricao, valor = match.groups()
                        transacoes.append({
                            "Banco/Cartão": banco,
                            "Fatura Referência": referencia,
                            "Data": formatar_data(data.strip(), ano_fatura, mes_fatura),
                            "Estabelecimento / Descrição": descricao.strip(),
                            "Valor": formatar_valor(valor),
                            "Motivo do Estorno": "" # Mantém a estrutura de colunas simétrica
                        })

    return transacoes

def processar_diretorio(diretorio_origem, arquivo_saida):
    todas_transacoes = []
    caminhos_pdfs = glob.glob(os.path.join(diretorio_origem, "*.pdf"))
    
    if not caminhos_pdfs:
        print(f"Nenhum arquivo PDF encontrado na pasta: {diretorio_origem}")
        return

    print(f"Encontrados {len(caminhos_pdfs)} PDFs. Iniciando extração...")
    
    for caminho in caminhos_pdfs:
        print(f"Processando: {os.path.basename(caminho)}...")
        transacoes_pdf = extrair_transacoes(caminho)
        todas_transacoes.extend(transacoes_pdf)
        
    if todas_transacoes:
        df = pd.DataFrame(todas_transacoes)
        df.to_excel(arquivo_saida, index=False)
        df.to_csv(arquivo_saida.replace('.xlsx', '.csv'), index=False) 
        print(f"\nExtração concluída com sucesso! Salvo em: {arquivo_saida} (e formato .csv)")
    else:
        print("\nNenhuma transação foi encontrada nos PDFs processados.")

if __name__ == "__main__":
    PASTA_PDFS = "./import" 
    ARQUIVO_EXCEL_SAIDA = "faturas_consolidadas.xlsx"
    processar_diretorio(PASTA_PDFS, ARQUIVO_EXCEL_SAIDA)