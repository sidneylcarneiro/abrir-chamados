print("Iniciando teste de importação...")

try:
    import pyperclip
    print("SUCESSO: Biblioteca 'pyperclip' foi importada.")
    
    import pyautogui
    print("SUCESSO: Biblioteca 'pyautogui' foi importada.")
    
    # Subprocess e time são da biblioteca padrão, então devem funcionar sem problemas.
    import subprocess
    import time
    print("SUCESSO: Bibliotecas padrão (subprocess, time) foram importadas.")
    
except Exception as e:
    print("\n--- ERRO! ---")
    print(f"Falha ao importar uma das bibliotecas.")
    print(f"O erro foi: {e}")

print("\nTeste de importação concluído.")
input("Pressione Enter para sair...")