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
ARQUIVO_LOG_FECHAMENTO = "log_fechamento_v4.txt" # Novo arquivo de log

CARREGAR = 10

def abrir_navegador(url):
    """Abre o navegador CHROME na URL especificada e maximiza a janela."""
    print(f"Abrindo navegador CHROME e acessando o Workspace...")
    try:
        subprocess.Popen([CHROME_PATH, url])
        print("Aguardando 20 segundos para a página carregar...")
        time.sleep(CARREGAR) # Espera fixa conforme solicitado
        time.sleep(1)
        print("Navegador aberto e maximizado.")
        return True
    except Exception as e: 
        print(f"ERRO ao abrir o navegador: {e}")
        return False

def esperar_imagem_aparecer(image_path, timeout=25, description=""):
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
            # ETAPA 3 e 4: Pesquisar chamado
            if not esperar_imagem_aparecer('lupa_pesquisa.png', description="Ícone de Lupa"): raise Exception("Ícone de lupa não encontrado.")
            if not encontrar_e_clicar('lupa_pesquisa.png', "Ícone de Lupa"): raise Exception("Falha ao clicar na lupa.")
            time.sleep(1)
            pyperclip.copy(chamado_id)
            gui.hotkey('ctrl', 'v')
            gui.press('enter')
            print(f"Pesquisou por: {chamado_id}")
            
            # ETAPA 5: Aguardar
            print("Aguardando 20 segundos após a pesquisa...")
            time.sleep(CARREGAR)

            # ETAPA 6: Clicar no resultado
            if not esperar_imagem_aparecer('computador_acessorios.png', description="'Computador e Acessórios'"): raise Exception("Resultado da busca 'Computador e Acessórios' não encontrado.")
            if not encontrar_e_clicar('computador_acessorios.png', "'Computador e Acessórios'"): raise Exception("Falha ao clicar em 'Computador e Acessórios'.")
            print("Aguardando 20 segundos após clicar no chamado...")
            time.sleep(CARREGAR)

            # ETAPA 7 a 10: Lógica de resolver
            if not esperar_imagem_aparecer('botao_resolver.png', description="Botão 'Resolver'"): raise Exception("Botão 'Resolver' não foi encontrado.")
            if not encontrar_e_clicar('botao_resolver.png', "Botão 'Resolver' (1º clique)"): raise Exception("Falha no primeiro clique em 'Resolver'.")

            print("Aguardando 20 segundos após o 1º clique...")
            time.sleep(CARREGAR)
            
            # Tenta o segundo clique se o botão ainda existir
            if esperar_imagem_aparecer('botao_resolver.png', timeout=15, description="Verificação para 2º clique"):
                if not encontrar_e_clicar('botao_resolver.png', "Botão 'Resolver' (2º clique)"): raise Exception("Falha no segundo clique em 'Resolver'.")
                print("Aguardando 20 segundos após o 2º clique...")
                time.sleep(CARREGAR)
            else:
                print("Botão 'Resolver' não apareceu para o 2º clique.")

            # Verifica uma última vez
            if esperar_imagem_aparecer('botao_resolver.png', timeout=15, description="Verificação final"):
                log_message += "FALHA - Botão 'Resolver' ainda visível após todas as tentativas."
                print("!!! O chamado pode não ter sido resolvido. Botão ainda na tela. !!!")
            else:
                log_message += "Resolvido com SUCESSO."
                print(log_message)

        except Exception as e:
            log_message += f"FALHA - Motivo: {e}"
            print(f"!!! {log_message} !!!")
        
        # Salva o log da operação
        with open(ARQUIVO_LOG_FECHAMENTO, 'a', encoding='utf-8') as f:
            f.write(log_message + "\n")
        
        # ETAPA 11: Interação com o usuário
        if i < len(chamados_para_fechar) - 1:
            resposta = gui.confirm(
                text=f'Chamado {chamado_id} processado.\nDeseja continuar?',
                title='Continuar Automação?',
                buttons=['Buscar próximo', 'Fechar programa']
            )
            
            if resposta == 'Fechar programa':
                print("Usuário escolheu fechar o programa.")
                break # Encerra o loop for
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