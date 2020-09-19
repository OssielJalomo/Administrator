from pynput.keyboard import Key, Controller
import pyautogui, sys
import win32com.client
import time
import os
shell = win32com.client.Dispatch("WScript.Shell")
keyboard = Controller()

def principal():
    print("ADMINISTRATOR")
    print("Sistema de automatización para NeuroExperimenter")
    print("Version: 1.1")
    print("Creado por: Ossiel Jalomo")
    print("email: cristian_ossi@hotmail.com")
    print("-------------------------------------------------")
    time.sleep(4)
    iteracion = 0
    while iteracion<=240:
        check()
        iteracion+=1


def imagen():
    crea_lienzo()
    time.sleep(0.5)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(0.5)
    shell.AppActivate("Neuro Experimenter Version 6.2")
    time.sleep(2)
    i=0
    while i<3:
        keyboard.press(Key.cmd)
        keyboard.press(Key.up)
        keyboard.release(Key.cmd)
        keyboard.release(Key.up)
        time.sleep(0.5)
        i=i+1
    pyautogui.moveTo(71,75)
    pyautogui.click()
    time.sleep(0.5)
    keyboard.press(Key.alt)
    keyboard.press(Key.print_screen)
    keyboard.release(Key.alt)
    keyboard.release(Key.print_screen)
    time.sleep(0.5)
    shell.AppActivate("Paint")
    time.sleep(0.5)
    pegar()
    time.sleep(0.5)
    guardar()
    cierra()

def grafica():
    crea_nota()
    time.sleep(0.5)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(0.5)
    shell.AppActivate("Neuro Experimenter Version 6.2")
    time.sleep(2)
    i=0
    while i<3:
        keyboard.press(Key.cmd)
        keyboard.press(Key.up)
        keyboard.release(Key.cmd)
        keyboard.release(Key.up)
        time.sleep(0.5)
        i=i+1
    pyautogui.moveTo(153,78)
    pyautogui.click()
    time.sleep(0.5)
    pyautogui.moveTo(1140, 567)
    pyautogui.click()
    pyautogui.click()
    time.sleep(0.5)
    shell.AppActivate("Bloc de notas")
    time.sleep(0.5)
    pegar()
    time.sleep(0.5)
    guardar()
    imagen()
    
   

def ejecuta():
    print("Ejecutando programa, espere por favor...")
    time.sleep(1)
    os.chdir("C:/Program Files (x86)/fhm/NeuroExperimenter/")
    time.sleep(0.5)
    os.popen("EEG.exe")
    time.sleep(10)
    check()


def check():
    lista = os.popen('tasklist /FI "IMAGENAME eq EEG.exe" /FO csv').readlines()
    #print(lista[1][1])
    #print(type(lista))
    if lista[1][1] == 'E':
        print("En ejecución")
        grafica()
    else:
        print("No existe proceso")
        ejecuta()

def pegar():
    keyboard.press(Key.ctrl)
    keyboard.press("v")
    keyboard.release(Key.ctrl)
    keyboard.release("v")
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)

def guardar():
    keyboard.press(Key.ctrl)
    keyboard.press("g")
    keyboard.release(Key.ctrl)
    keyboard.release("g")
    tiempo = str(time.time())
    keyboard.type(tiempo)
    time.sleep(0.5)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)

def cierra():
    time.sleep(1)
    shell.AppActivate("Paint")
    keyboard.press(Key.alt)
    keyboard.press(Key.f4)
    keyboard.release(Key.alt)
    keyboard.release(Key.f4)
    time.sleep(1)
    shell.AppActivate("Bloc de notas")
    keyboard.press(Key.alt)
    keyboard.press(Key.f4)
    keyboard.release(Key.alt)
    keyboard.release(Key.f4)
    time.sleep(1)

def crea_nota():
    time.sleep(1)
    keyboard.press(Key.cmd)
    keyboard.press("r")
    keyboard.release(Key.cmd)
    keyboard.release("r")
    time.sleep(1)
    keyboard.type("notepad")
def crea_lienzo():
    time.sleep(1)
    keyboard.press(Key.cmd)
    keyboard.press("r")
    keyboard.release(Key.cmd)
    keyboard.release("r")
    time.sleep(1)
    keyboard.type("mspaint")

principal()