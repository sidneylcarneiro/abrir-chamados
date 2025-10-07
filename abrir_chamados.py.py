import pyautogui as gui
import time
import subprocess
import pyperclip

# --- CONFIGURAÇÕES GLOBAIS ---
URL_SERVICENOW = "https://afya.service-now.com/esc?id=applications"
# Verifique se este é o caminho correto para o seu Google Chrome
CHROME_PATH = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
CONFIDENCE_LEVEL = 0.9

# --- DADOS DO FORMULÁRIO ---
NUMERO_TELEFONE = "21994186239" 
TEXTO_DESEJO = "Análise de performance"

# --- NOMES DOS ARQUIVOS ---
ARQUIVO_CHAMADOS = "chamados.txt"
ARQUIVO_CHAMADOS_ABERTOS = "chamados_abertos.txt"

def abrir_navegador(url):
    """Abre o navegador CHROME na URL especificada e maximiza a janela."""
    print(f"Abrindo navegador CHROME e acessando: {url}")
    try:
        subprocess.Popen([CHROME_PATH, url])
        time.sleep(5)
        gui.hotkey('win', 'up')
        time.sleep(1)
        print("Navegador aberto e maximizado.")
        return True
    except FileNotFoundError:
        print(f"ERRO: Navegador Chrome não encontrado no caminho: {CHROME_PATH}")
        print("Verifique a variável 'CHROME_PATH' no código.")
        return False
    except Exception as e:
        print(f"ERRO ao abrir o navegador: {e}")
        return False

def esperar_imagem_aparecer(image_path, timeout=15, description=""):
    """Espera ativamente até que uma imagem apareça na tela ou o tempo limite seja atingido."""
    print(f"Aguardando '{description}' ({image_path}) aparecer...")
    start_time = time.time()
    while time.time() - start_time < timeout:
        try:
            if gui.locateOnScreen(image_path, confidence=CONFIDENCE_LEVEL):
                print(f"Elemento '{description}' encontrado.")
                return True
        except gui.PyAutoGUIException:
            pass
        time.sleep(0.5)
    print(f"ERRO: Tempo esgotado. O elemento '{description}' não apareceu em {timeout} segundos.")
    return False

def encontrar_e_clicar(image_path, description=""):
    """Localiza uma imagem e clica nela de forma 'humanizada'."""
    print(f"Procurando e clicando em: '{description}' ({image_path})...")
    location = gui.locateOnScreen(image_path, confidence=CONFIDENCE_LEVEL)
    if location:
        center = gui.center(location)
        gui.moveTo(center, duration=0.25)
        gui.mouseDown()
        time.sleep(0.1)
        gui.mouseUp()
        print(f"Clicou em: '{description}'.")
        return True
    print(f"ERRO: Não foi possível encontrar '{description}' para clicar.")
    return False

def interagir_com_campo_texto(label_image_path, text_to_type, description=""):
    """Localiza o label e COLA o texto no campo abaixo dele."""
    if not esperar_imagem_aparecer(label_image_path, description=f"Label '{description}'"):
        return False
    
    label_location = gui.locateOnScreen(label_image_path, confidence=CONFIDENCE_LEVEL)
    if label_location:
        x = label_location.left + (label_location.width / 2)
        y = label_location.top + label_location.height + 25
        
        gui.moveTo(x, y, duration=0.25)
        gui.click()
        time.sleep(0.5)
        
        pyperclip.copy(text_to_type)
        gui.hotkey('ctrl', 'v')
        
        time.sleep(0.5)
        print(f"COLOU '{text_to_type}' no campo '{description}'.")
        return True
    return False

def copiar_numero_chamado(ritm_image_path, description="Número do Chamado"):
    """Localiza o número RITM, clica duas vezes para selecionar e copia."""
    if not esperar_imagem_aparecer(ritm_image_path, timeout=20, description=description):
        return None
    location = gui.locateOnScreen(ritm_image_path, confidence=CONFIDENCE_LEVEL, grayscale=True)
    if location:
        center = gui.center(location)
        gui.moveTo(center, duration=0.25)
        gui.doubleClick()
        time.sleep(0.5)
        gui.hotkey('ctrl', 'c')
        time.sleep(0.5)
        chamado_copiado = pyperclip.paste()
        print(f"Chamado copiado: {chamado_copiado}")
        if "RITM" in chamado_copiado:
            return chamado_copiado
    print("ERRO: Padrão RITM não encontrado ou falha ao copiar.")
    return None

def main():
    """Orquestra todo o processo de abertura de chamados."""
    if not abrir_navegador(URL_SERVICENOW):
        return False

    try:
        with open(ARQUIVO_CHAMADOS, 'r', encoding='utf-8') as f:
            chamados_pendentes = [line for line in f.read().splitlines() if line.strip()]
        if not chamados_pendentes:
            print(f"Arquivo '{ARQUIVO_CHAMADOS}' está vazio. Encerrando.")
            return True
    except FileNotFoundError:
        print(f"ERRO: Arquivo '{ARQUIVO_CHAMADOS}' não encontrado.")
        return False

    # Inicia o loop para processar cada chamado do arquivo
    for i, linha_chamado in enumerate(chamados_pendentes):
        print(f"\n--- Processando chamado {i+1}/{len(chamados_pendentes)}: '{linha_chamado}' ---")

        try:
            # ETAPA 1: NAVEGAÇÃO PELOS MENUS
            if not esperar_imagem_aparecer('alvo_01.png', description="Menu TI"): raise Exception("Menu TI não encontrado")
            if not encontrar_e_clicar('alvo_01.png', "Menu TI"): raise Exception("Falha ao clicar no Menu TI")
            
            if not esperar_imagem_aparecer('alvo_02.png', description="Menu Infraestrutura"): raise Exception("Menu Infraestrutura não encontrado")
            if not encontrar_e_clicar('alvo_02.png', "Menu Infraestrutura"): raise Exception("Falha ao clicar no Menu Infraestrutura")

            if not esperar_imagem_aparecer('alvo_03.png', description="Menu Computador e Acessórios"): raise Exception("Menu Computador e Acessórios não encontrado")
            if not encontrar_e_clicar('alvo_03.png', "Menu Computador e Acessórios"): raise Exception("Falha ao clicar no Menu Computador e Acessórios")

            # ETAPA 2: PREENCHIMENTO DO FORMULÁRIO
            if not interagir_com_campo_texto('campo_telefone.png', NUMERO_TELEFONE, "Número de Telefone"): raise Exception("Falha ao preencher o telefone")
            if not interagir_com_campo_texto('campo_o_que_deseja.png', TEXTO_DESEJO, "O que deseja?"): raise Exception("Falha ao preencher 'O que deseja?'")
            if not esperar_imagem_aparecer('opcao_analise_performance.png', description="Opção 'Análise de performance'"): raise Exception("Opção 'Análise de performance' não apareceu")
            if not encontrar_e_clicar('opcao_analise_performance.png', "Opção 'Análise de performance'"): raise Exception("Falha ao selecionar 'Análise de performance'")
            if not interagir_com_campo_texto('campo_detalhes.png', linha_chamado, "Detalhes sobre sua exigência"): raise Exception("Falha ao preencher os detalhes")

            # ETAPA 3: ENVIO E CAPTURA DO NÚMERO
            if not encontrar_e_clicar('botao_enviar.png', "Botão Enviar"): raise Exception("Falha ao clicar em Enviar")

            chamado_aberto = copiar_numero_chamado('ritm_padrao.png')
            if chamado_aberto:
                with open(ARQUIVO_CHAMADOS_ABERTOS, 'a', encoding='utf-8') as f:
                    f.write(chamado_aberto + "\n")
                print(f"Chamado '{chamado_aberto}' salvo com SUCESSO.")
            else:
                raise Exception("Falha ao copiar o número do chamado (RITM)")

        except Exception as e:
            print(f"!!! ERRO AO PROCESSAR O CHAMADO: '{linha_chamado}' !!!")
            print(f"Motivo: {e}")
            with open(ARQUIVO_CHAMADOS_ABERTOS, 'a', encoding='utf-8') as f:
                f.write(f"FALHA NO CHAMADO: {linha_chamado} | Motivo: {e}\n")
            # Continua para o próximo chamado em vez de parar o script todo

        # ETAPA 4: PREPARAÇÃO PARA O PRÓXIMO LOOP
        if i < len(chamados_pendentes) - 1:
            print("Retornando à página inicial para o próximo chamado...")
            # Navega para a URL inicial para garantir um estado limpo
            gui.hotkey('ctrl', 'l') # Foca na barra de endereço
            time.sleep(0.5)
            pyperclip.copy(URL_SERVICENOW) # Usa o método de colar para evitar problemas de layout
            gui.hotkey('ctrl', 'v')
            gui.press('enter')
            time.sleep(4) # Espera a página carregar

    print("\n--- Todos os chamados foram processados. ---")
    return True

if __name__ == "__main__":
    # Configurações de segurança e pausa do PyAutoGUI
    gui.FAILSAFE = True  # Move o mouse para o canto superior esquerdo (0,0) para abortar
    gui.PAUSE = 0.2      # Pequena pausa global entre todas as ações

    # Executa a função principal
    sucesso = main()

    # Mostra um alerta final para o usuário
    if sucesso:
        gui.alert(text="Automação de chamados concluída!", title="Status da Automação")
    else:
        gui.alert(text="Automação falhou. Verifique o console para identificar o erro.", title="Status da Automação")