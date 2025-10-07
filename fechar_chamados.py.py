import pyautogui as gui
import time
import subprocess
import pyperclip

# --- CONFIGURAÇÕES GLOBAIS ---
URL_WORKSPACE_V4 = "https://afya.service-now.com/now/workspace/agent/home/sub/non_record/layout/params/list-title/My%20work/table/task/query/active%3Dtrue%5Eassigned_toDYNAMIC90d1921e5f510100a9ad2572f2b477fe/workspace-config-id/7b24ceae5304130084acddeeff7b12a3/word-wrap/false/disable-quick-edit/true"
CHROME_PATH = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
CONFIDENCE_LEVEL = 0.9

# --- NOMES DOS ARQUIVOS ---
ARQUIVO_CHAMADOS_FECHAR = "chamados_abertos.txt" 
ARQUIVO_LOG_FECHAMENTO = "log_fechamento_v5.txt" # Novo arquivo de log

# --------------------------------------------------------------------
# --- CONFIGURAÇÕES DE TEMPO (em segundos) ---
# Altere os valores aqui para ajustar as esperas do robô
# --------------------------------------------------------------------
TEMPO_ABRIR_NAVEGADOR = 20      # Espera para o navegador abrir e a página inicial carregar
TEMPO_APOS_PESQUISA = 20        # Pausa após pesquisar o chamado
TEMPO_APOS_CLIQUE_CHAMADO = 20  # Pausa após abrir o chamado
TEMPO_APOS_RESOLVER_1 = 20      # Pausa após o 1º clique em "Resolver"
TEMPO_APOS_RESOLVER_2 = 20      # Pausa após o 2º clique em "Resolver"


def abrir_navegador(url):
    """Abre o navegador CHROME na URL especificada e maximiza a janela."""
    print(f"Abrindo navegador CHROME e acessando o Workspace...")
    try:
        subprocess.Popen([CHROME_PATH, url])
        print(f"Aguardando {TEMPO_ABRIR_NAVEGADOR} segundos para a página carregar...")
        time.sleep(TEMPO_ABRIR_NAVEGADOR)
        print("Navegador aberto e maximizado.")
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
    """Orquestra o processo de fechamento de chamados com interação do usuário."""
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

    for i, chamado_id in enumerate(chamados_para_fechar):
        print(f"\n--- Processando chamado {i+1}/{len(chamados_para_fechar)}: '{chamado_id}' ---")
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

            # --- NOVA ETAPA ADICIONADA AQUI ---
            if not esperar_imagem_aparecer('fechar_pesquisa.png', description="Botão 'Fechar Pesquisa'"): raise Exception("Botão para fechar a pesquisa não foi encontrado.")
            if not encontrar_e_clicar('fechar_pesquisa.png', "Botão 'Fechar Pesquisa'"): raise Exception("Falha ao clicar para fechar a pesquisa.")
            # ------------------------------------

            # Clicar no resultado
            if not esperar_imagem_aparecer('computador_acessorios.png', description="'Computador e Acessórios'"): raise Exception("Resultado da busca 'Computador e Acessórios' não encontrado.")
            if not encontrar_e_clicar('computador_acessorios.png', "'Computador e Acessórios'"): raise Exception("Falha ao clicar em 'Computador e Acessórios'.")
            print(f"Aguardando {TEMPO_APOS_CLIQUE_CHAMADO} segundos após clicar no chamado...")
            time.sleep(TEMPO_APOS_CLIQUE_CHAMADO)

            # Lógica de resolver
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
        
        if i < len(chamados_para_fechar) - 1:
            resposta = gui.confirm(
                text=f'Chamado {chamado_id} processado.\nDeseja continuar?',
                title='Continuar Automação?',
                buttons=['Buscar próximo', 'Fechar programa']
            )
            if resposta == 'Fechar programa':
                print("Usuário escolheu fechar o programa.")
                break
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
