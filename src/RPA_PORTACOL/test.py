import pyautogui
import subprocess
from time import sleep
import pytesseract
from PIL import Image
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
ruta_a_chrome = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
subprocess.Popen([ruta_a_chrome])
sleep(8)
pyautogui.click("C:\\RPA_PORTACOL_INTERACTIVO\\resources\\img\\refresh.PNG")
sleep(4)
#pyautogui.click("C:\\RPA_PORTACOL_INTERACTIVO\\resources\\img\\url.PNG")
#sleep(2)
pyautogui.write("https://www.portabilidadcolombia.com.co/?handler=Check")
sleep(1)
pyautogui.press("Enter")
sleep(6)
pyautogui.click("C:\\RPA_PORTACOL_INTERACTIVO\\resources\\img\\check.PNG")
sleep(4)
pyautogui.click("C:\\RPA_PORTACOL_INTERACTIVO\\resources\\img\\ingresar.PNG")
sleep(4)
pyautogui.hotkey('ctrl', 'a') 
sleep(1)
pyautogui.press('delete')
sleep(1)
pyautogui.write("3224515519")
sleep(2)
#pyautogui.click("C:\\RPA_PORTACOL_INTERACTIVO\\resources\\img\\buscar.PNG")
for _ in range(4):
    pyautogui.press('tab')
    sleep(0.5)  # Espera de 0.5 segundos entre cada Tab
pyautogui.press("Enter")
sleep(8)
try:
    
    img =pyautogui.screenshot()
    # Extrae el texto usando Tesseract
    text = pytesseract.image_to_string(img)
    # Imprime el texto
    print("Texto extraído:\n", text)

    # Normaliza texto (por si hay mayúsculas o espacios raros)
    if "no existen registros" in text.lower():
        print("✅ Resultado: No existen registros")
        # Aquí puedes marcar la fila como gestionada, registrar, etc.
    else:
        print("🔍 Resultado: Sí existen registros o no se detectó bien")

except Exception as e:
    print(f"Error : {e}")        
sleep(3)
