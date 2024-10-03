import pyautogui
import time

print("Mova o mouse para o botão desejado e aguarde 5 segundos...")
time.sleep(5)  # Tempo para você posicionar o mouse

x, y = pyautogui.position()  # Captura a posição do mouse
print(f"As coordenadas do mouse são: {x}, {y}")
