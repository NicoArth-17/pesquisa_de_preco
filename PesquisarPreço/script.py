import pandas as pd
from selenium import webdriver # Comandar navegador controlado
from selenium.webdriver.chrome.service import Service # responsável por iniciar e parar o chromedriver
from webdriver_manager.chrome import ChromeDriverManager # Gerenciar arquvio ChromeDriver
from selenium.webdriver.common.by import By # Selecionar elementos
from selenium.webdriver.common.keys import Keys # Utilizar comandos do teclado
import time
import win32com.client as win32

# Criando df com arquivo inicial
produtos_df = pd.read_excel(r'C:\Users\mobishopgamer\Documents\RepositorioGitHub\pesquisa_de_preco\PesquisarPreço\arquivos\buscas.xlsx')

# Instalar arquivo temporário ChromeDrive a partir da versão atual do Chrome
servico = Service(ChromeDriverManager().install())

# Abrir navegador com tela maximizada
options = webdriver.ChromeOptions()
options.add_argument('--start-maximized')

# Criar navegador
nav = webdriver.Chrome(service=servico, options=options)

def pesquisa_google(nav, nome_produto, nomes_banidos, preço_min, preço_max):

    # Entrar google shopping
    nav.get('https://shopping.google.com/?nord=1')

    # Pesquisar produto
    nav.find_element(By.XPATH, '//*[@id="REsRA"]').send_keys(nome_produto)
    nav.find_element(By.XPATH, '//*[@id="REsRA"]').send_keys(Keys.ENTER)

    # Filtrar Valor min
    nav.find_element(By.NAME, 'lower').send_keys(str(preço_min))
    # Filtrar Valor max
    nav.find_element(By.NAME, 'upper').send_keys(str(preço_max))
    # Aplicar filtro de valor
    nav.find_element(By.NAME, 'upper').send_keys(Keys.ENTER)
        
    # Listando elementos com informações do produto pesquisado
    info_produto = nav.find_elements(By.CLASS_NAME, 'sh-dgr__grid-result')

    # Listando termos banidos separadamente
    termos_banidos = nomes_banidos.lower()
    termos_banidos = termos_banidos.split(' ')

    # Listando termos do nome do produto separadamente
    termos_produto = nome_produto.lower()
    termos_produto = termos_produto.split(' ')

    # Criando lista com os nome, preço e link do produto
    resultado_pesquisa = []
    for elemento in info_produto:
        
        # Selecionando elemento com nome do produto
        nome_produto = elemento.find_element(By.CLASS_NAME, 'tAxDx').text.lower()

        # Verificando termo banido
        produto_banido = False
        for termo in termos_banidos:
            if termo in nome_produto:
                produto_banido = True
        
        # Verificando se o nome do produto bate com o pesquisado
        produto_aceito = True
        for termo in termos_produto:
            if termo not in nome_produto:
                produto_aceito = False

        # Filtrando nome do produto
        if not produto_banido and produto_aceito:

            # Selecionando elemento com preço do produto
            preço = elemento.find_element(By.CLASS_NAME, 'a8Pemb').text

            # Selecionando elemento com link do produto
            elemento_link_filho = elemento.find_element(By.CLASS_NAME, 'aULzUe')
            elemento_link_pai = elemento_link_filho.find_element(By.XPATH, '..')
            link = elemento_link_pai.get_attribute('href')
            
            resultado_pesquisa.append((nome_produto, preço, link))

    return resultado_pesquisa

def pesquisa_buscape(nav, nome_produto, nomes_banidos, preço_min, preço_max):

    # Entrar buscapé
    nav.get('https://www.buscape.com.br/')

    # Pesquisar produto
    nav.find_element(By.XPATH, '//*[@id="new-header"]/div[1]/div/div/div[3]/div/div/div[2]/div/div[1]/input').send_keys(nome_produto)
    nav.find_element(By.XPATH, '//*[@id="new-header"]/div[1]/div/div/div[3]/div/div/div[2]/div/div[1]/input').send_keys(Keys.ENTER)

    time.sleep(3)
  
    # Listando elementos com informações do produto pesquisado
    info_produto = nav.find_elements(By.CLASS_NAME, 'SearchCard_ProductCard__1D3ve')

    # Listando termos banidos separadamente
    termos_banidos = nomes_banidos.lower()
    termos_banidos = termos_banidos.split(' ')

    # Listando termos do nome do produto separadamente
    termos_produto = nome_produto.lower()
    termos_produto = termos_produto.split(' ')

    # Criando lista com os nome, preço e link do produto
    resultado_pesquisa = []
    for elemento in info_produto:
        
        # Selecionando elemento com nome do produto
        nome_produto = elemento.find_element(By.CLASS_NAME, 'SearchCard_ProductCard_Name__ZaO5o').text.lower()

        # Verificando termo banido
        produto_banido = False
        for termo in termos_banidos:
            if termo in nome_produto:
                produto_banido = True
        
        # Verificando se o nome do produto bate com o pesquisado
        produto_aceito = True
        for termo in termos_produto:
            if termo not in nome_produto:
                produto_aceito = False

        # Filtrando nome do produto
        if not produto_banido and produto_aceito:

            # Selecionando elemento com preço do produto
            preço = elemento.find_element(By.CLASS_NAME, 'Text_MobileHeadingS__Zxam2').text
            preço_format = preço.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
            preço_format = float(preço_format)

            # Filtrando Valores
            if preço_min <= preço_format <= preço_max:

                # Selecionando elemento com link do produto
                link = elemento.find_element(By.CLASS_NAME, 'SearchCard_ProductCard_Inner__7JhKb').get_attribute('href')
            
                resultado_pesquisa.append((nome_produto, preço, link))

    return resultado_pesquisa

# Criando df separados para cada produto
iphone_df = pd.DataFrame(index=None)
rtx_df = pd.DataFrame(index=None)

# Rodando functions para pesquisar no gshopping e no buscapé percorrendo os produtos da tabela inicial
for linha in produtos_df.index:
    nomes_banidos = produtos_df.loc[linha, 'Termos banidos']
    nome_produto = produtos_df.loc[linha, 'Nome']
    preço_max = produtos_df.loc[linha, 'Preço máximo']
    preço_min = produtos_df.loc[linha, 'Preço mínimo']

    gshop = pesquisa_google(nav, nome_produto, nomes_banidos, preço_min, preço_max)
    if gshop:
        gshop_df = pd.DataFrame(gshop, columns=['PRODUTO', 'PREÇO', 'LINK'])
        
        if 'iphone' in gshop[0][0]:
            iphone_df = pd.concat([iphone_df, gshop_df])

        elif 'rtx' in gshop[0][0]:
            rtx_df = pd.concat([rtx_df, gshop_df])
    else:
        gshop_df = None

    buscape = pesquisa_buscape(nav, nome_produto, nomes_banidos, preço_min, preço_max)
    if buscape:
        buscape_df = pd.DataFrame(buscape, columns=['PRODUTO', 'PREÇO', 'LINK'])

        if 'iphone' in buscape[0][0]:
            iphone_df = pd.concat([iphone_df, buscape_df])

        elif 'rtx' in buscape[0][0]:
            rtx_df = pd.concat([rtx_df, buscape_df])

    else:
        buscape_df = None

# Ordeando preços do menor para o maior
iphone_df = iphone_df.sort_values(['PREÇO'])
rtx_df = rtx_df.sort_values(['PREÇO'])

# Exportando df dos resultados
iphone_df.to_excel(r'C:\Users\mobishopgamer\Documents\RepositorioGitHub\pesquisa_de_preco\PesquisarPreço\arquivos\resultados\iphone.xlsx')
rtx_df.to_excel(r'C:\Users\mobishopgamer\Documents\RepositorioGitHub\pesquisa_de_preco\PesquisarPreço\arquivos\resultados\rtx.xlsx')

# Envio de email

# verificando se existe informaçao dentro da tabela de resultados
if len(iphone_df.index) > 0 and len(rtx_df.index) > 0:
    # vou enviar email
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'nicolarthur17+Teste@hotmail.com'
    mail.Subject = 'Resultado da pesquisa de compra'
    mail.HTMLBody = f"""
    <p>Prezado,</p>
    <p>Foi encontrados alguns produtos dentro da faixa de preço desejada. Segue tabela com detalhes dos 10 resultados mais baratos:</p>
    {iphone_df[:10].to_html(index=False)}
    <br>
    {rtx_df[:10].to_html(index=False)}
    <p>Qualquer dúvida estou à disposição</p>
    <p>Att., Nicolas Arthur</p>
    """

    attachment = r'C:\Users\mobishopgamer\Documents\RepositorioGitHub\pesquisa_de_preco\PesquisarPreço\arquivos\resultados\iphone.xlsx'
    mail.Attachments.Add(str(attachment))
    attachment = r'C:\Users\mobishopgamer\Documents\RepositorioGitHub\pesquisa_de_preco\PesquisarPreço\arquivos\resultados\rtx.xlsx'
    mail.Attachments.Add(str(attachment))

    mail.Send()

nav.quit()  