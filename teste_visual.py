import pyautogui
import cv2
import numpy as np
from PIL import Image
import os

# --- CONFIGURAÇÕES DO TESTE ---
# Coloque aqui o nome exato do arquivo de imagem que você quer testar
IMAGEM_PARA_ENCONTRAR = 'menu_tres_pontos.png' 
# Ajuste a confiança para o teste. Comece com um valor mais baixo.
CONFIANCA_DO_TESTE = 0.7 
# Manter True é geralmente mais robusto
USAR_ESCALA_DE_CINZA = True 

# --- INÍCIO DO SCRIPT DE TESTE ---
print(f"--- Iniciando teste para a imagem: '{IMAGEM_PARA_ENCONTRAR}' ---")

# 1. Verificar se o arquivo de imagem existe
if not os.path.exists(IMAGEM_PARA_ENCONTRAR):
    print(f"❌ ERRO CRÍTICO: O arquivo '{IMAGEM_PARA_ENCONTRAR}' não foi encontrado. Teste abortado.")
else:
    try:
        # 2. Tirar um screenshot da tela inteira
        screenshot_pil = pyautogui.screenshot()
        # Converte a imagem para um formato que o OpenCV (cv2) pode usar para desenhar
        screenshot_cv = cv2.cvtColor(np.array(screenshot_pil), cv2.COLOR_RGB2BGR)
        print("Screenshot da tela inteira capturado.")

        # 3. Tentar localizar a imagem DENTRO do screenshot que acabamos de tirar
        print(f"Procurando por '{IMAGEM_PARA_ENCONTRAR}' com confiança >= {CONFIANCA_DO_TESTE}...")
        location = pyautogui.locate(IMAGEM_PARA_ENCONTRAR, screenshot_pil, confidence=CONFIANCA_DO_TESTE, grayscale=USAR_ESCALA_DE_CINZA)

        # 4. Analisar e reportar o resultado
        if location:
            print(f"\n✅ SUCESSO! Imagem encontrada na localização: {location}")
            # Desenha um retângulo verde ao redor da área encontrada
            ponto_inicial = (location.left, location.top)
            ponto_final = (location.left + location.width, location.top + location.height)
            cv2.rectangle(screenshot_cv, ponto_inicial, ponto_final, (0, 255, 0), 3) # Cor BGR: Verde
            
            caminho_saida = 'debug_SUCESSO.png'
            cv2.imwrite(caminho_saida, screenshot_cv)
            print(f"Uma imagem de depuração foi salva como '{caminho_saida}' com a área encontrada marcada em verde.")
        
        else:
            print(f"\n❌ FALHA! Imagem '{IMAGEM_PARA_ENCONTRAR}' não foi encontrada com a confiança mínima de {CONFIANCA_DO_TESTE}.")
            caminho_saida = 'debug_FALHA.png'
            # Salva o screenshot original sem nenhuma marcação
            cv2.imwrite(caminho_saida, screenshot_cv)
            print(f"O screenshot da tela no momento da falha foi salvo como '{caminho_saida}'.")
            print("Verifique este arquivo para garantir que a imagem estava realmente visível e sem obstruções.")

    except Exception as e:
        print(f"\n🚨 Ocorreu um erro inesperado durante o teste: {e}")
        import traceback
        traceback.print_exc()