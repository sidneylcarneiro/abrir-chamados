import pyautogui as gui
import time
import subprocess
import pyperclip
import os
import traceback

# --- CONFIGURAÇÕES GLOBAIS ---
URL_SERVICENOW = "https://afya.service-now.com/esc?id=applications"
# Verifique se este é o caminho correto para o seu Google Chrome
CHROME_PATH = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
CONFIDENCE_LEVEL = 0.8

# --- DADOS DO FORMULÁRIO ---
NUMERO_TELEFONE = "21994186239"
TEXTO_DESEJO = "Análise de performance"

# --- NOMES DOS ARQUIVOS DE IMAGEM E TEXTO ---
ARQUIVO_CHAMADOS = "chamados.txt"
ARQUIVO_CHAMADOS_ABERTOS = "chamados_abertos.txt"
IMAGEM_MENU_CHROME = 'menu_tres_pontos.png'
IMAGEM_ZOOM_100 = 'zoom_100_porcento.png'


def verificar_resolucao_tela(largura_desejada=1920, altura_desejada=1080):
    """Verifica se a resolução da tela corresponde à desejada."""
    largura_atual, altura_atual = gui.size()
    print(f"Resolução detectada: {largura_atual}x{altura_atual}")
    if (largura_atual, altura_atual) == (largura_desejada, altura_desejada):
        print("Resolução da tela está correta (1920x1080).")
        return True
    else:
        mensagem_erro = f"ERRO: A resolução da tela é {largura_atual}x{altura_atual}, mas o script requer {largura_desejada}x{altura_desejada}. Ajuste e tente novamente."
        print(mensagem_erro)
        gui.alert(mensagem_erro)
        return False


def verificar_zoom_chrome():
    """Clica no menu do Chrome e verifica se o zoom está em 100%."""
    print("Verificando o zoom do Chrome...")
    if not os.path.exists(IMAGEM_MENU_CHROME):
        gui.alert(f"ERRO: A imagem '{IMAGEM_MENU_CHROME}' não foi encontrada.")
        return False

    try:
        location = gui.locateOnScreen(IMAGEM_MENU_CHROME, confidence=CONFIDENCE_LEVEL, grayscale=True)
        if not location:
            screenshot_name = 'debug_screenshot.png'
            gui.screenshot(screenshot_name)
            print(f"ERRO: Menu do Chrome não localizado. Screenshot salvo como '{screenshot_name}'.")
            return False

        gui.click(gui.center(location))
        time.sleep(1)

        if gui.locateOnScreen(IMAGEM_ZOOM_100, confidence=CONFIDENCE_LEVEL):
            print("Zoom está correto (100%).")
            gui.press('esc')
            return True
        else:
            gui.press('esc')
            gui.alert("ERRO: O zoom do Chrome não está em 100%. Ajuste (Ctrl + 0) e tente novamente.")
            return False

    except Exception as e:
        print(f"!!! ERRO INESPERADO AO VERIFICAR O ZOOM: {e} !!!")
        traceback.print_exc()
        gui.alert("Ocorreu um erro inesperado ao verificar o zoom. Verifique o console.")
        return False


def abrir_navegador(url):
    """Abre o navegador CHROME na URL especificada."""
    print(f"Abrindo navegador CHROME e acessando: {url}")
    try:
        subprocess.Popen([CHROME_PATH, url])
        time.sleep(5)
        print("Navegador aberto")
        return True
    except FileNotFoundError:
        print(f"ERRO: Navegador Chrome não encontrado no caminho: {CHROME_PATH}")
        return False
    except Exception as e:
        print(f"ERRO ao abrir o navegador: {e}")
        return False


def esperar_imagem_aparecer(image_path, timeout=15, description=""):
    """Espera ativamente até que uma imagem apareça na tela."""
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
    try:
        location = gui.locateOnScreen(image_path, confidence=CONFIDENCE_LEVEL)
        if location:
            center = gui.center(location)
            gui.moveTo(center, duration=0.25)
            gui.mouseDown()
            time.sleep(0.1)
            gui.mouseUp()
            print(f"Clicou em: '{description}'.")
            return True
    except Exception as e:
        print(f"ERRO ao tentar encontrar e clicar em '{description}': {e}")
    print(f"ERRO: Não foi possível encontrar '{description}' para clicar.")
    return False


def interagir_com_campo_texto(label_image_path, text_to_type, description="", y_offset=40):
    """Localiza o label, clica a uma distância Y abaixo dele e COLA o texto."""
    if not esperar_imagem_aparecer(label_image_path, description=f"Label '{description}'"):
        return False

    try:
        label_location = gui.locateOnScreen(label_image_path, confidence=CONFIDENCE_LEVEL)
        if label_location:
            x = label_location.left + (label_location.width / 2)
            y = label_location.top + label_location.height + y_offset

            print(f"Label '{description}' encontrado. Clicando em (x={int(x)}, y={int(y)}) para inserir o texto.")
            gui.moveTo(x, y, duration=0.25)
            gui.click()
            time.sleep(0.5)

            gui.hotkey('ctrl', 'a')
            gui.press('delete')
            time.sleep(0.2)

            pyperclip.copy(text_to_type)
            gui.hotkey('ctrl', 'v')

            time.sleep(0.5)
            print(f"COLOU '{text_to_type}' no campo '{description}'.")
            return True
    except Exception as e:
        print(f"ERRO ao interagir com o campo de texto '{description}': {e}")
    return False


def copiar_numero_chamado(ritm_image_path, description="Número do Chamado"):
    """Localiza o número RITM, clica duas vezes para selecionar e copia."""
    if not esperar_imagem_aparecer(ritm_image_path, timeout=20, description=description):
        return None
    try:
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
    except Exception as e:
        print(f"ERRO ao copiar número do chamado: {e}")
    print("ERRO: Padrão RITM não encontrado ou falha ao copiar.")
    return None


def main():
    """Orquestra todo o processo de abertura de chamados."""
    print("Configurando a tela para 'Duplicar' (Espelhar)...")
    gui.hotkey('win', 'p')
    time.sleep(1)
    gui.press('home')
    time.sleep(0.5)
    gui.press('down')
    time.sleep(0.5)
    gui.press('enter')
    time.sleep(0.5)
    gui.press('esc')
    time.sleep(5)
    print("Tela configurada para 'Duplicar'.")

    if not verificar_resolucao_tela():
        return False
    if not abrir_navegador(URL_SERVICENOW):
        return False
    if not verificar_zoom_chrome():
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

    for i, linha_chamado in enumerate(chamados_pendentes):
        print(f"\n--- Processando chamado {i + 1}/{len(chamados_pendentes)}: '{linha_chamado}' ---")
        try:
            if not esperar_imagem_aparecer('alvo_01.png', description="Menu TI"): raise Exception("Menu TI não encontrado")
            if not encontrar_e_clicar('alvo_01.png', "Menu TI"): raise Exception("Falha ao clicar no Menu TI")

            if not esperar_imagem_aparecer('alvo_02.png', description="Menu Infraestrutura"): raise Exception("Menu Infraestrutura não encontrado")
            if not encontrar_e_clicar('alvo_02.png', "Menu Infraestrutura"): raise Exception("Falha ao clicar no Menu Infraestrutura")

            if not esperar_imagem_aparecer('alvo_03.png', description="Menu Computador e Acessórios"): raise Exception("Menu Computador e Acessórios não encontrado")
            if not encontrar_e_clicar('alvo_03.png', "Menu Computador e Acessórios"): raise Exception("Falha ao clicar no Menu Computador e Acessórios")

            if not interagir_com_campo_texto('campo_telefone.png', NUMERO_TELEFONE, "Número de Telefone", y_offset=40): raise Exception("Falha ao preencher o telefone")
            if not interagir_com_campo_texto('campo_o_que_deseja.png', TEXTO_DESEJO, "O que deseja?", y_offset=40): raise Exception("Falha ao preencher 'O que deseja?'")

            if not esperar_imagem_aparecer('opcao_analise_performance.png', description="Opção 'Análise de performance'"): raise Exception("Opção 'Análise de performance' não apareceu")
            if not encontrar_e_clicar('opcao_analise_performance.png', "Opção 'Análise de performance'"): raise Exception("Falha ao selecionar 'Análise de performance'")

            if not interagir_com_campo_texto('campo_detalhes.png', linha_chamado, "Detalhes sobre sua exigência", y_offset=40): raise Exception("Falha ao preencher os detalhes")

            # --- NOVA ABORDAGEM: BUSCA POR TEXTO (CTRL + F) ---
            print("Iniciando envio do formulário via busca de texto (Ctrl+F)...")
            time.sleep(1)

            # 1. Pressiona "Ctrl + F" para abrir a busca
            gui.hotkey('ctrl', 'f')
            time.sleep(1)

            # 2. Cola "Enviar" na barra de busca (mais confiável que digitar)
            pyperclip.copy("Enviar")
            gui.hotkey('ctrl', 'v')
            time.sleep(1)

            # 3. Pressiona Enter para encontrar o texto
            gui.press('enter')
            time.sleep(1)

            # 4. Pressiona Esc para fechar a busca (o foco permanece no botão)
            gui.press('esc')
            time.sleep(1)

            # 5. Pressiona Enter para "clicar" no botão em foco
            print("Pressionando 'Enter' final para enviar o formulário...")
            gui.press('enter')
            # --- FIM DA NOVA ABORDAGEM ---

            print("Verificando se número do chamado foi copiado corretamente...")
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

        if i < len(chamados_pendentes) - 1:
            print("Retornando à página inicial para o próximo chamado...")
            gui.hotkey('ctrl', 'l')
            time.sleep(0.5)
            pyperclip.copy(URL_SERVICENOW)
            gui.hotkey('ctrl', 'v')
            gui.press('enter')
            if not esperar_imagem_aparecer('alvo_01.png', timeout=20, description="Página inicial (Menu TI)"):
                print("ERRO: Não foi possível recarregar a página inicial. Abortando.")
                return False

    print("\n--- Todos os chamados foram processados. ---")
    return True


if __name__ == "__main__":
    gui.FAILSAFE = True
    gui.PAUSE = 0.2
    sucesso = main()

    if sucesso:
        gui.alert(text="Automação de chamados concluída!", title="Status da Automação")
    else:
        gui.alert(text="Automação falhou. Verifique o console para identificar o erro.", title="Status da Automação")

    print("\nRestaurando a configuração de tela para 'Estender'...")
    gui.hotkey('win', 'p')
    time.sleep(1)
    gui.press('home')
    time.sleep(0.5)
    gui.press('down')
    time.sleep(0.5)
    gui.press('down')
    time.sleep(0.5)
    gui.press('enter')
    time.sleep(0.5)
    gui.press('esc')
    time.sleep(5)
    print("Configuração de tela restaurada para 'Estender'.")