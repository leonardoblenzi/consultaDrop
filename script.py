import pandas as pd
import requests
import re
from urllib.parse import urlparse

# Função para extrair o slug do URL
def slug_from_url(url):
    path = urlparse(url).path
    parts = path.strip('/').split('/')
    return parts[-1] if parts[-1] else parts[-2]

# Função para obter preço e estoque via API do WooCommerce
def get_product_info(slug):
    api_url = f'https://meudropbrasil.com/wp-json/wc/store/products?slug={slug}'
    try:
        response = requests.get(api_url, timeout=20)
        products = response.json()
        if response.status_code == 200 and products:
            product = products[0]
            # Preço em centavos → converter para reais
            price_cents = int(product['prices']['price'])
            price = price_cents / 100

            # Estoque: texto e quantidade numérica
            stock_text = product['stock_availability']['text']
            stock_qty = None
            if stock_text:
                m = re.search(r'(\d+)', stock_text)
                stock_qty = int(m.group(1)) if m else None

            return price, stock_qty
    except Exception as e:
        print(f'Erro ao acessar {api_url}: {e}')
    return None, None

# Carregar a planilha original, especificando a aba "Produtos"
caminho_planilha = 'produtos.xlsx'
df = pd.read_excel(caminho_planilha, sheet_name='Produtos')

# Para cada URL, extrair slug, consultar API e preencher preço e estoque
for index, row in df.iterrows():
    url_produto = row['URL']
    slug = slug_from_url(url_produto)
    preco, estoque = get_product_info(slug)
    # Atualizando as células de "Preço de Custo" e "Estoque"
    df.loc[index, 'Preço de Custo'] = preco
    df.loc[index, 'Estoque'] = estoque

# Salvar a planilha atualizada, sobrescrevendo os valores antigos
df.to_excel(caminho_planilha, index=False, sheet_name='Produtos')
print('Planilha atualizada com sucesso!')
