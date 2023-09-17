from selenium.webdriver.common.by import By

from rotinas_robo import util
from rotinas_robo.google import planilha_google
from time import sleep
from selenium.webdriver.support.ui import Select
import pyautogui
from rotinas_robo.tjpe import rotina_tjpe
from selenium.webdriver.common.action_chains import ActionChains

nome_planilha = 'Processos_return'
aba = 'Pendentes'

def autenticar_planilha(nome_planilha, range):
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
    client = gspread.authorize(creds)
    sheet = client.open(nome_planilha).worksheet(range)
    return sheet

def planilha_por_nome(sheet):
    return sheet.get_all_records()

def autenticar_hack():
    chrome_options = webdriver.ChromeOptions()
    preferences = {"directory_upgrade": True,
                   "safebrowsing.enabled": True,
                   "download.prompt_for_download": False,
                   "plugins.always_open_pdf_externally": True}
    chrome_options.add_experimental_option("prefs", preferences)
    driver = webdriver.Chrome(chrome_options=chrome_options)
    driver.maximize_window()
    wdw = WebDriverWait(driver, 90)
    driver.get('https://pje.tjpe.jus.br/1g/login.seam')
    nome_sistema_locator = (By.CLASS_NAME, 'nomeSistema')
    wdw.until(
        presence_of_element_located(nome_sistema_locator)
    )
    sleep(2)
    try:
        btn_login = driver.find_element(By.ID, 'loginAplicacaoButton')
    except Exception as e:
        driver.switch_to.frame('ssoFrame')
        sleep(1)
        btn_login = driver.find_element(By.ID, 'kc-pje-office')
    try:
        btn_login.click()
        sleep(5)
    except Exception as e:
        print('Autenticando')
        pyautogui.press('1')
        pyautogui.press('2')
        pyautogui.press('3')
        pyautogui.press('4')
        pyautogui.press('enter')
    # Testa token expirado
    try:
        div_certificado_expirado = driver.find_element(By.ID, 'popupAlertaCertificadoProximoDeExpirarContentDiv')
        certificado_expirado = div_certificado_expirado.find_element(By.TAG_NAME, 'i')
        certificado_expirado.click()
        sleep(2)
    except Exception:
        print('Certificado na validade')
    return driver

def pesquisar_processo_pje_pe(driver, nro_processo):
    try:
        print('Entrou na rotina de Pesquisar processo')
        wdw = WebDriverWait(driver, 90)
        driver.get('https://pje.tjpe.jus.br/pje/Processo/ConsultaProcesso/listView.seam')
        sleep(3)
        # print('Depois do wait')
        # try:
        #     print('Antes de pegar a URL')
        #     url_atual = driver.current_url
        #     print(f'URL_ATUAL: {url_atual}')
        #     if 'https://pje.tjpe.jus.br/1g/QuadroAviso' in url_atual:
        #         print('Antes de chamar a página')
        #         driver.get('https://pje.tjpe.jus.br/pje/Processo/ConsultaProcesso/listView.seam')
        #         print('Depois de chamar a página')
        #         sleep(1)
        #     else:
        #         # Testa limpar os campos
        #         print('Antes do botão liompar')
        #         btn_limpar = driver.find_element(By.ID, 'fPP:clearButtonProcessos')
        #         btn_limpar.click()
        #         print('Depois do botão liompar')
        #         sleep(3)
        # except Exception as e:
        #     print('Bug PJE')
        #     print(e)
        imput_processo_locator = (By.ID, 'fPP:numeroProcesso:numeroSequencial')
        wdw.until(
            presence_of_element_located(imput_processo_locator)
        )
        nro_processo_partes = util.get_partes_cnj(nro_processo)
        # sequencial
        input_seq_nro_cnj = driver.find_element(By.ID, 'fPP:numeroProcesso:numeroSequencial')
        input_seq_nro_cnj.send_keys(nro_processo_partes.get('sequencial'))
        sleep(1)
        # digito verificador
        input_digito = driver.find_element(By.ID, 'fPP:numeroProcesso:numeroDigitoVerificador')
        input_digito.send_keys(nro_processo_partes.get('digito'))
        sleep(1)
        # Ano
        input_ano = driver.find_element(By.ID, 'fPP:numeroProcesso:Ano')
        input_ano.send_keys(nro_processo_partes.get('ano'))
        sleep(1)
        input_origem = driver.find_element(By.ID, 'fPP:numeroProcesso:NumeroOrgaoJustica')
        input_origem.send_keys(nro_processo_partes.get('orgao_justica'))
        sleep(1)
        # botão pesquisar
        btn_pesquisar = driver.find_element(By.ID, 'fPP:searchProcessos')
        btn_pesquisar.click()
        sleep(5)
        processos_table_locator = (By.ID, 'fPP:processosTable')
        wdw.until(
            presence_of_element_located(processos_table_locator)
        )
        qtd_resultado_pesquisa = driver.find_element(By.ID, 'fPP:processosTable').find_element(By.CLASS_NAME,
                                                                                               'text-muted').text
        if qtd_resultado_pesquisa == '0 resultados encontrados.' or \
                qtd_resultado_pesquisa == 'resultados encontrados.':
            resp = {
                'status': False,
                'driver': driver,
                'erro': 'Processo não localizado no PJE virstual'
            }
            return resp
        sleep(4)
        resp = {
            'status': True,
            'driver': driver,
            'erro': None
        }
        return resp
    except Exception as e:
        print('Erro ao pesquisar processo')
        print(e)
        resp = {
            'status': False,
            'driver': driver,
            'erro': str(e)
        }
        return resp

def inserir_movimento_hack(sheet, id, movimento, posicao):
    cell = sheet.find(id)
    if posicao == 1:
        sheet.update_acell(f'D{cell.row}', movimento)
    elif posicao == 2:
        sheet.update_acell(f'E{cell.row}', movimento)
    elif posicao == 3:
        sheet.update_acell(f'F{cell.row}', movimento)
    elif posicao == 4:
        sheet.update_acell(f'G{cell.row}', movimento)
    elif posicao == 5:
        sheet.update_acell(f'H{cell.row}', movimento)

def atualizar_status_planilha_by_id(sheet, nro_processo, status):
    cell = sheet.find(nro_processo)
    sheet.update_acell(f'A{cell.row}', status)

def inserir_advogado_hack(sheet, id, texto):
    cell = sheet.find(id)
    sheet.update_acell(f'I{cell.row}', texto)

def transferir_peticao_inicial_hack(processo):
    downloads = f'C:/Users/gabri/Downloads/'
    print(downloads)
    diretorio_downloads = os.listdir(downloads)
    if len(diretorio_downloads) > 0:
        arquivo_origem = f'{downloads}{diretorio_downloads[0]}'
        arquivo_destino = f'C:/Projetos/pp/iniciais/INICIAL_{processo}.pdf'
        os.rename(arquivo_origem, arquivo_destino)
        return True
    else:
        return False

def transferir_fatura_hack(processo):
    downloads = f'C:/Users/gabri/Downloads/'
    print(downloads)
    diretorio_downloads = os.listdir(downloads)
    if len(diretorio_downloads) > 0:
        arquivo_origem = f'{downloads}{diretorio_downloads[0]}'
        arquivo_destino = f'C:/Projetos/pp/faturas/FATURA_{processo}.pdf'
        os.rename(arquivo_origem, arquivo_destino)
        return True
    else:
        return False

def limpar_pasta_downloads_hack():
    downloads = f'C:/Users/gabri/Downloads/'
    diretorio_downloads = os.listdir(downloads)
    if len(diretorio_downloads) > 0:
        for arquivo in diretorio_downloads:
            path_arquivo = f'{downloads}{arquivo}'
            if os.path.isdir(path_arquivo):
                shutil.rmtree(path_arquivo, ignore_errors=True)
            else:
                print(f'Removendo arquivo: {arquivo}')
                os.remove(path_arquivo)

try:
    sheet = autenticar_planilha(nome_planilha, aba)
    planilha = planilha_por_nome(sheet)
    sleep(1)
    print('Autenticado no PJe')
    driver = autenticar_hack()
    for linha in planilha:
        status = linha.get('STATUS')
        nro_processo = linha.get('NRO_PROCESSO')
        id_planilha = linha.get('ID')
        if status == 'Pendente' or \
                status == 'Movimentos Capturados' or \
                status == 'Petição Capturada' or \
                status == 'Advogados Capturados':
            resposta = pesquisar_processo_pje_pe(driver, nro_processo)
            link_processo = driver.find_element(By.ID, 'fPP:processosTable:tb').find_elements(By.TAG_NAME, 'a')[1]
            print('Antes do Clique')
            link_processo.click()
            print('Depois do clique')
            # trocar de janela e pesquisa parte passiva
            util.verifica_alert(driver)
            sleep(2)
            driver.switch_to.window(driver.window_handles[1])
            sleep(1)
            if status == 'Pendente':
                print(f'Capturando movimentos do processo: {id_planilha}')
                status = resposta['status']
                if status == False:
                    erro = resposta['erro']
                    if erro == 'Processo não localizado no PJE virstual':
                        erro = 'Processo não localizado no PJE virtual'
                        print(erro)
                    else:
                        print(erro)
                elif status == True:
                    driver = resposta['driver']
                print('Depois de trocar de aba')
                div_tl = driver.find_element(By.ID, 'divTimeLine:divEventosTimeLine')
                divs = div_tl.find_elements(By.CLASS_NAME, 'media')
                movimentos = []
                última_data = ''
                for div in divs:
                    # Verifica se a div é data
                    classe_div = div.get_attribute('class')
                    if classe_div == 'media data':
                        última_data = div.text
                        continue
                    elif 'media interno' in classe_div:
                        movimento = dict()
                        movimento['data'] = última_data
                        titulo_movimento = div.text
                        if '\n' in titulo_movimento:
                            titulo_movimento = titulo_movimento.split('\n')[0]
                        movimento['titulo'] = titulo_movimento
                        movimentos.append(movimento)
                    else:
                        continue
                print(f'Quantidade de movimentos do processo {nro_processo}: {len(movimentos)} ')
                pos = 1
                for movimento in movimentos[:5]:
                    data = movimento.get('data')
                    movimento = movimento.get('titulo')
                    texto_planilha = f'Data: {data}\n Movimento: {movimento}'
                    inserir_movimento_hack(sheet, id_planilha, texto_planilha, pos)
                    pos += 1
                    sleep(5)
                atualizar_status_planilha_by_id(sheet, id_planilha, 'Movimentos Capturados')
                sleep(1)
                status = 'Movimentos Capturados'
            if status == 'Movimentos Capturados':
                qtd_advogados = 0
                advogados = ''
                sleep(1)
                btn_mais_detalhes = driver.find_element(By.CLASS_NAME, 'mais-detalhes')
                btn_mais_detalhes.click()
                sleep(1)
                div_polo_ativo = driver.find_element(By.ID, 'poloAtivo')
                linhas_polo_ativo = div_polo_ativo.find_elements(By.TAG_NAME, 'li')
                for linha_polo_ativo in linhas_polo_ativo:
                    conteudo = linha_polo_ativo.text
                    if '(ADVOGADO)' in conteudo:
                        nome = conteudo[:conteudo.index('- OAB')].strip()
                        oab = conteudo[conteudo.index('- OAB') + 5:conteudo.index('- CPF')].strip()
                        cpf = conteudo[conteudo.index('- CPF') + 6:conteudo.index('(ADVOGADO)')].strip()
                        qtd_advogados += 1
                        if nome != '' and oab != '' and cpf != '':
                            texto = f'Nome: {nome}, oab: {oab}, cpf: {cpf} \n'
                            advogados = advogados + texto
                div_polo_passivo = driver.find_element(By.ID, 'poloPassivo')
                linhas_polo_passivo = div_polo_passivo.find_elements(By.TAG_NAME, 'li')
                for linha_polo_passivo in linhas_polo_passivo:
                    conteudo = linha_polo_passivo.text
                    if '(ADVOGADO)' in conteudo:
                        nome = conteudo[:conteudo.index('- OAB')].strip()
                        oab = conteudo[conteudo.index('- OAB') + 5:conteudo.index('- CPF')].strip()
                        cpf = conteudo[conteudo.index('- CPF') + 6:conteudo.index('(ADVOGADO)')].strip()
                        qtd_advogados += 1
                        if nome != '' and oab != '' and cpf != '':
                            if nome != '' and oab != '' and cpf != '':
                                texto = f'Nome: {nome}, oab: {oab}, cpf: {cpf} \n'
                                advogados = advogados + texto
                inserir_advogado_hack(sheet, id_planilha, advogados)
                atualizar_status_planilha_by_id(sheet, id_planilha, 'Advogados Capturados')
                sleep(1)
            if status == 'Advogados Capturados':
                util.limpar_pasta_downloads_hack()
                btn_autos_processo = driver.find_element(By.CLASS_NAME, 'filtros-download')
                btn_autos_processo.click()
                sleep(1)
                ddl_tipo_documento = Select(driver.find_element(By.ID, 'navbar:cbTipoDocumento'))
                try:
                    ddl_tipo_documento.select_by_visible_text(
                        'Petição Inicial (Outras)')
                    sleep(1)
                except Exception as e:
                    erro = 'Erro ao selecionar tipo de documento dos autos do processo'
                    print(erro)
                div_download_autos_processo = driver.find_element(By.ID, 'navbar:botoesDownload')
                btn_download_autos_processo = div_download_autos_processo.find_element(By.XPATH,
                                                                                       "//input[@value='Download']")
                sleep(1)
                try:
                    btn_download_autos_processo.click()
                except:
                    for i in range(9):
                        pyautogui.hotkey('tab')
                    pyautogui.hotkey('space')
                    sleep(2)
                sleep(10)
                peticao_encontrada = checa_peticao_baixada_hack()
                if peticao_encontrada == True:
                    arquivo_transferido = False
                    if peticao_encontrada == True:
                        arquivo_transferido = transferir_peticao_inicial_hack(nro_processo)
                else:
                    icone_menu_abas = driver.find_element(By.CLASS_NAME, 'icone-menu-abas')
                    icone_menu_abas.click()
                    sleep(1)
                    menu_documentos = driver.find_element(By.ID, 'navbar:linkAbaDocumentos')
                    menu_documentos.click()
                    sleep(3)
                    peticao_encontrada = False
                    possui_paginacao = False
                    try:
                        div_paginacao = driver.find_element(By.CLASS_NAME, 'rich-inslider-track-decor-2')
                        tabela_paginacao = driver.find_element(By.XPATH,
                                                               '/html/body/div[1]/div[2]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div/div/div/div[1]/div/div[2]/span/div/div/form/table')
                        colunas = tabela_paginacao.find_elements(By.TAG_NAME, 'td')
                        coluna = colunas[2].text
                        range_paginacao = int(coluna)
                        possui_paginacao = True
                        sleep(2)
                    except Exception:
                        range_paginacao = 1
                    for i in range(range_paginacao):
                        if i == 0:
                            if possui_paginacao == True:
                                input_paginacao = driver.find_element(By.CLASS_NAME, 'rich-inslider-field-right')
                                actions = ActionChains(driver)
                                actions.double_click(on_element=input_paginacao)
                                actions.send_keys('3')
                                actions.perform()
                                sleep(5)
                        if i == 1:
                            input_paginacao = driver.find_element(By.CLASS_NAME, 'rich-inslider-field-right')
                            actions = ActionChains(driver)
                            actions.double_click(on_element=input_paginacao)
                            actions.send_keys('1')
                            actions.perform()
                            sleep(5)
                        elif i == 2:
                            input_paginacao = driver.find_element(By.CLASS_NAME, 'rich-inslider-field-right')
                            actions = ActionChains(driver)
                            actions.double_click(on_element=input_paginacao)
                            actions.send_keys('2')
                            actions.perform()
                            sleep(5)
                        tabela_documentos = driver.find_element(By.ID, 'processoDocumentoGridList')
                        linhas_tabela_documentos = tabela_documentos.find_elements(By.TAG_NAME, 'tr')
                        for linha_tabela_documentos in linhas_tabela_documentos:
                            colunas = linha_tabela_documentos.find_elements(By.TAG_NAME, 'td')
                            if len(colunas) > 5:
                                coluna_tipo_documento = colunas[6].text
                                if ('PETIÇÃO INICIAL' in coluna_tipo_documento) or \
                                        ('Petição Inicial' in coluna_tipo_documento):
                                    limpar_pasta_downloads_hack()
                                    btn_download = linha_tabela_documentos.find_elements(By.TAG_NAME, 'a')[0]
                                    if btn_download.get_attribute('title') == 'Visualizar':
                                        btn_download.click()
                                        sleep(3)
                                        driver.switch_to.window(driver.window_handles[2])
                                        btn_download_detalhe = driver.find_element(By.ID, 'detalhesDocumento:download')
                                        btn_download_detalhe.click()
                                        sleep(2)
                                        driver.switch_to.alert.accept()
                                        sleep(5)
                                        driver.close()
                                        driver.switch_to.window(driver.window_handles[1])
                                        sleep(1)
                                    else:
                                        btn_download.click()
                                    sleep(5)
                                    peticao_encontrada = checa_peticao_baixada_hack()
                                    break
                    if peticao_encontrada == False:
                        erro = 'Petição Inicial não encontrada'
                        print(erro)
                    arquivo_transferido = False
                    if peticao_encontrada == True:
                        arquivo_transferido = transferir_peticao_inicial_hack(nro_processo)
                    if arquivo_transferido == True:
                        print('Inicial capturada')
                    else:
                        print('Erro ao baixar peticao inicial tjpe')
                atualizar_status_planilha_by_id(sheet, id_planilha, 'Petição Capturada')
                sleep(1)
                status = 'Petição Capturada'
            if status == 'Petição Capturada':
                limpar_pasta_downloads_hack()
                icone_menu_abas = driver.find_element(By.CLASS_NAME, 'icone-menu-abas')
                icone_menu_abas.click()
                sleep(1)
                menu_documentos = driver.find_element(By.ID, 'navbar:linkAbaDocumentos')
                menu_documentos.click()
                sleep(3)
                fatura_encontrada = False
                possui_paginacao = False
                try:
                    div_paginacao = driver.find_element(By.CLASS_NAME, 'rich-inslider-track-decor-2')
                    tabela_paginacao = driver.find_element(By.XPATH,
                                                           '/html/body/div[1]/div[2]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div/div/div/div[1]/div/div[2]/span/div/div/form/table')
                    colunas = tabela_paginacao.find_elements(By.TAG_NAME, 'td')
                    coluna = colunas[2].text
                    range_paginacao = int(coluna)
                    possui_paginacao = True
                    sleep(2)
                except Exception:
                    range_paginacao = 1
                for i in range(range_paginacao):
                    if i == 0:
                        if possui_paginacao == True:
                            input_paginacao = driver.find_element(By.CLASS_NAME, 'rich-inslider-field-right')
                            actions = ActionChains(driver)
                            actions.double_click(on_element=input_paginacao)
                            actions.send_keys('3')
                            actions.perform()
                            sleep(5)
                    if i == 1:
                        input_paginacao = driver.find_element(By.CLASS_NAME, 'rich-inslider-field-right')
                        actions = ActionChains(driver)
                        actions.double_click(on_element=input_paginacao)
                        actions.send_keys('1')
                        actions.perform()
                        sleep(5)
                    elif i == 2:
                        input_paginacao = driver.find_element(By.CLASS_NAME, 'rich-inslider-field-right')
                        actions = ActionChains(driver)
                        actions.double_click(on_element=input_paginacao)
                        actions.send_keys('2')
                        actions.perform()
                        sleep(5)
                    tabela_documentos = driver.find_element(By.ID, 'processoDocumentoGridList')
                    linhas_tabela_documentos = tabela_documentos.find_elements(By.TAG_NAME, 'tr')
                    for linha_tabela_documentos in linhas_tabela_documentos:
                        colunas = linha_tabela_documentos.find_elements(By.TAG_NAME, 'td')
                        if len(colunas) > 5:
                            coluna_documento = colunas[6].text
                            if 'FATURAS' in coluna_documento:
                                coluna_tipo_documento = colunas[7].text
                                if 'Documento de Comprovação' in coluna_tipo_documento:
                                    util.limpar_pasta_downloads_hack()
                                    btn_download = linha_tabela_documentos.find_elements(By.TAG_NAME, 'a')[0]
                                    if btn_download.get_attribute('title') == 'Visualizar':
                                        btn_download.click()
                                        sleep(3)
                                        driver.switch_to.window(driver.window_handles[2])
                                        btn_download_detalhe = driver.find_element(By.ID, 'detalhesDocumento:download')
                                        btn_download_detalhe.click()
                                        sleep(2)
                                        driver.switch_to.alert.accept()
                                        sleep(5)
                                        driver.close()
                                        driver.switch_to.window(driver.window_handles[1])
                                        sleep(1)
                                    else:
                                        btn_download.click()
                                    sleep(5)
                                    fatura_encontrada = checa_peticao_baixada_hack()
                                    arquivo_transferido = False
                                    if fatura_encontrada == True:
                                        arquivo_transferido = transferir_fatura_hack(nro_processo)
                                    if arquivo_transferido == True:
                                        print('Fatura capturada')
                                    else:
                                        print('Erro ao baixar fatura')
                                atualizar_status_planilha_by_id(sheet, id_planilha, 'Fatura Capturada')
                                sleep(1)

except Exception as e:
    print(e)
    print('Erro ao capturar processos')