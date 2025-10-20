import pyautogui as gui
import time
import subprocess
import pyperclip
import os

# --- CONFIGURAÇÕES GLOBAIS ---
URL_WORKSPACE_V4 = "https://afya.service-now.com/now/workspace/agent/home/sub/non_record/layout/params/list-title/My%20work/table/task/query/active%3Dtrue%5Eassigned_toDYNAMIC90d1921e5f510100a9ad2572f2b477fe/workspace-config-id/7b24ceae5304130084acddeeff7b12a3/word-wrap/false/disable-quick-edit/true"
CHROME_PATH = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
CONFIDENCE_LEVEL = 0.9

# --- NOMES DOS ARQUIVOS ---
ARQUIVO_CHAMADOS_FECHAR = "chamados_abertos.txt"
ARQUIVO_LOG_FECHAMENTO = "log_fechamento_v5.txt"

# --- CONFIGURAÇÕES DE TEMPO (em segundos) ---
TEMPO_ABRIR_NAVEGADOR = 6
TEMPO_APOS_PESQUISA = 6
TEMPO_APOS_CLIQUE_CHAMADO = 6
TEMPO_APOS_RESOLVER_1 = 15
TEMPO_APOS_RESOLVER_2 = 6


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


def verificar_e_ajustar_zoom(imagem_menu='menu_tres_pontos.png', imagem_zoom_100='zoom_100_porcento.png', imagem_zoom_67='zoom_67_porcento.png'):
    """
    Verifica o zoom do Chrome e o ajusta para 67%.
    Começa resetando para 100% para ter um ponto de partida confiável.
    """
    print("--- Iniciando verificação e ajuste de zoom para 67% ---")
    
    print(f"Procurando pelo menu do Chrome ('{imagem_menu}')...")
    location = gui.locateOnScreen(imagem_menu, confidence=0.8)
    if not location:
        print("ERRO: Não foi possível encontrar o menu do Chrome (menu_tres_pontos.png).")
        gui.alert("ERRO: A imagem 'menu_tres_pontos.png' não foi encontrada para ajustar o zoom.")
        return False
    gui.click(gui.center(location))
    time.sleep(1)

    print(f"Verificando se o zoom já está em 67% ('{imagem_zoom_67}')...")
    if gui.locateOnScreen(imagem_zoom_67, confidence=CONFIDENCE_LEVEL):
        print("Zoom já está correto (67%).")
        gui.press('esc')
        return True

    print("Zoom não está em 67%. Resetando para 100% (Ctrl + 0)...")
    gui.hotkey('ctrl', '0')
    time.sleep(1.5)

    print(f"Confirmando o reset para 100% ('{imagem_zoom_100}')...")
    if not gui.locateOnScreen(imagem_zoom_100, confidence=CONFIDENCE_LEVEL):
        print("ERRO: Não foi possível confirmar o reset do zoom para 100%. O ajuste automático falhou.")
        gui.press('esc')
        return False
    print("Zoom resetado para 100% com sucesso.")

    print("Reduzindo o zoom para 67% (pressionando Ctrl + '-' 3 vezes)...")
    for _ in range(3):
        gui.hotkey('ctrl', '-')
        time.sleep(0.5)

    time.sleep(1)
    print(f"Verificando se o zoom final é 67% ('{imagem_zoom_67}')...")
    if gui.locateOnScreen(imagem_zoom_67, confidence=CONFIDENCE_LEVEL):
        print("SUCESSO: Zoom ajustado para 67%.")
        gui.press('esc')
        return True
    else:
        print("ERRO FINAL: Após as tentativas, o zoom não foi ajustado para 67%.")
        gui.alert("ERRO: Falha ao ajustar o zoom do navegador para 67%. Por favor, ajuste manualmente e tente novamente.")
        gui.press('esc')
        return False


def abrir_navegador(url):
    """Abre o navegador CHROME na URL especificada e maximiza a janela."""
    print(f"Abrindo navegador CHROME e acessando o Workspace...")
    try:
        subprocess.Popen([CHROME_PATH, url])
        print(f"Aguardando {TEMPO_ABRIR_NAVEGADOR} segundos para a página carregar...")
        time.sleep(TEMPO_ABRIR_NAVEGADOR)
        print("Navegador aberto.")
        return True
    except Exception as e:
        print(f"ERRO ao abrir o navegador: {e}")
        return False


def esperar_imagem_aparecer(image_path, timeout=40, description=""):
    """Espera ativamente até que uma imagem apareça na tela."""
    print(f"Aguardando '{description}' ({image_path}) aparecer (até {timeout}s)...")
    start_time = time.time()
    while time.time() - start_time < timeout:
        try:
            if gui.locateOnScreen(image_path, confidence=CONFIDENCE_LEVEL):
                print(f"Elemento '{description}' encontrado.")
                return True
        except gui.PyAutoGUIException: pass
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
        gui.mouseDown(); time.sleep(0.1); gui.mouseUp()
        print(f"Clicou em: '{description}'.")
        return True
    print(f"ERRO: Não foi possível encontrar '{description}' para clicar.")
    return False


def main():
    """Orquestra o processo de fechamento de chamados."""
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
    time.sleep(2)
    print("Tela configurada para 'Duplicar'.")

    if not verificar_resolucao_tela():
        return False

    try:
        with open(ARQUIVO_CHAMADOS_FECHAR, 'r', encoding='utf-8') as f:
            chamados_para_fechar = [line for line in f.read().splitlines() if line.strip()]
        if not chamados_para_fechar:
            print(f"Arquivo '{ARQUIVO_CHAMADOS_FECHAR}' está vazio. Nada a fazer.")
            return True
    except FileNotFoundError:
        print(f"ERRO: Arquivo de entrada '{ARQUIVO_CHAMADOS_FECHAR}' não encontrado.")
        return False

    if not abrir_navegador(URL_WORKSPACE_V4):
        return False

    if not verificar_e_ajustar_zoom():
        print("Encerrando o script devido à falha no ajuste de zoom.")
        return False

    for i, chamado_id in enumerate(chamados_para_fechar):
        print(f"\n--- Processando chamado {i + 1}/{len(chamados_para_fechar)}: '{chamado_id}' ---")
        log_message = f"Chamado {chamado_id}: "

        try:
            # Pesquisar chamado
            if not esperar_imagem_aparecer('lupa_pesquisa.png', description="Ícone de Lupa"): raise Exception("Ícone de lupa não encontrado.")
            if not encontrar_e_clicar('lupa_pesquisa.png', "Ícone de Lupa"): raise Exception("Falha ao clicar na lupa.")
            time.sleep(1)
            pyperclip.copy(chamado_id)
            gui.hotkey('ctrl', 'v')
            gui.press('enter')
            print(f"Pesquisou por: {chamado_id}")

            print(f"Aguardando {TEMPO_APOS_PESQUISA} segundos após a pesquisa...")
            time.sleep(TEMPO_APOS_PESQUISA)

            if not esperar_imagem_aparecer('fechar_pesquisa.png', description="Botão 'Fechar Pesquisa'"): raise Exception("Botão para fechar a pesquisa não foi encontrado.")
            if not encontrar_e_clicar('fechar_pesquisa.png', "Botão 'Fechar Pesquisa'"): raise Exception("Falha ao clicar para fechar a pesquisa.")

            if not esperar_imagem_aparecer('computador_acessorios.png', description="'Computador e Acessórios'"): raise Exception("Resultado da busca 'Computador e Acessórios' não encontrado.")
            if not encontrar_e_clicar('computador_acessorios.png', "'Computador e Acessórios'"): raise Exception("Falha ao clicar em 'Computador e Acessórios'.")
            print(f"Aguardando {TEMPO_APOS_CLIQUE_CHAMADO} segundos após clicar no chamado...")
            time.sleep(TEMPO_APOS_CLIQUE_CHAMADO)

            if not esperar_imagem_aparecer('botao_resolver.png', description="Botão 'Resolver'"): raise Exception("Botão 'Resolver' não foi encontrado.")
            if not encontrar_e_clicar('botao_resolver.png', "Botão 'Resolver' (1º clique)"): raise Exception("Falha no primeiro clique em 'Resolver'.")

            print(f"Aguardando {TEMPO_APOS_RESOLVER_1} segundos após o 1º clique...")
            time.sleep(TEMPO_APOS_RESOLVER_1)

            if esperar_imagem_aparecer('botao_resolver.png', timeout=15, description="Verificação para 2º clique"):
                if not encontrar_e_clicar('botao_resolver.png', "Botão 'Resolver' (2º clique)"): raise Exception("Falha no segundo clique em 'Resolver'.")
                print(f"Aguardando {TEMPO_APOS_RESOLVER_2} segundos após o 2º clique...")
                time.sleep(TEMPO_APOS_RESOLVER_2)
            else:
                print("Botão 'Resolver' não apareceu para o 2º clique.")

            if esperar_imagem_aparecer('botao_resolver.png', timeout=15, description="Verificação final"):
                log_message += "FALHA - Botão 'Resolver' ainda visível após todas as tentativas."
                print("!!! O chamado pode não ter sido resolvido. Botão ainda na tela. !!!")
            else:
                log_message += "Resolvido com SUCESSO."
                print(log_message)

        except Exception as e:
            log_message += f"FALHA - Motivo: {e}"
            print(f"!!! {log_message} !!!")

        with open(ARQUIVO_LOG_FECHAMENTO, 'a', encoding='utf-8') as f:
            f.write(log_message + "\n")

        # --- Bloco de confirmação MODIFICADO com temporizador ---
        if i < len(chamados_para_fechar) - 1:
            print("Aguardando 5 segundos para a confirmação do usuário...")
            # A função alert retorna o texto do botão se clicado, ou 'Timeout' se o tempo esgotar
            resposta = gui.alert(
                text=f'Chamado {chamado_id} processado.\nO robô continuará para o próximo em 5 segundos.\n\nClique em "Parar" para interromper.',
                title='Continuar Automação?',
                button='Parar Automação',
                timeout=5000  # 5000 milissegundos = 5 segundos
            )

            # Se o usuário clicou no botão "Parar Automação"
            if resposta == 'Parar Automação':
                print("Usuário escolheu parar o programa.")
                break
            else: # Se o tempo esgotou (resposta é 'Timeout')
                print("Tempo esgotado. Continuando para o próximo chamado automaticamente.")
        # --- Fim do bloco modificado ---
        else:
            print("Não há mais chamados na lista.")

    print("\n--- Automação finalizada. ---")
    return True


if __name__ == "__main__":
    gui.FAILSAFE = True
    gui.PAUSE = 0.2

    sucesso = main()

    if sucesso:
        gui.alert(text="Automação de fechamento de chamados concluída!", title="Status da Automação")
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
    time.sleep(2)
    print("Configuração de tela restaurada para 'Estender'.")