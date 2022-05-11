# cv2.cvtColor takes a numpy ndarray as an argument
import numpy as nm
import pyautogui
import pytesseract
import time
import turtle
import cv2
from pynput.mouse import Button, Controller

from PIL import ImageGrab
def getCoordinate(location):
    print("Place Cursor on the " + location + " of the box")
    print("3")
    time.sleep(1)
    print("2")
    time.sleep(1)
    print("1")
    time.sleep(1)
    return pyautogui.position()

def getBox():
    x1, y1 = getCoordinate("top left")
    x2, y2 = getCoordinate("bottom right")
    return(x1, y1, x2, y2)

def main():
    pytesseract.pytesseract.tesseract_cmd =r'C:\\Automation\\Tesseract\\tesseract.exe'
    mouse = Controller()
    print("Place Cursor on save button")
    time.sleep(3)
    buttonPos = pyautogui.position()
    time.sleep(2)
    mouse.position = (buttonPos)
    t = turtle.Turtle()
    turtle.tracer(0, 0)
    t.speed(10)
    t.color("red", "red")
    screen = turtle.Screen()
    screen.setup(width=0.333, height=0.333, startx=1500, starty=0)

    
    box1 = getBox()
    box2 = getBox()
    print("Box1:", box1)
    print("Box2", box2)
    
    
    while(True):
        t.begin_fill()
        for _ in range(4):
            t.forward(200)
            t.right(90)
        t.end_fill()
        turtle.update()            
        # ImageGrab-To capture the screen image in a loop.
        # Bbox used to capture a specific area.
        cap1 = ImageGrab.grab(bbox =(box1))
        cap2 = ImageGrab.grab(bbox = (box2))
        # Converted the image to monochrome for it to be easily
        # read by the OCR and obtained the output String.
        box1text = pytesseract.image_to_string(cv2.cvtColor(nm.array(cap1), cv2.COLOR_BGR2GRAY),lang ='eng')
        box2text = pytesseract.image_to_string(cv2.cvtColor(nm.array(cap2), cv2.COLOR_BGR2GRAY),lang ='eng') 
        box1text = box1text.replace("\n", "")
        box2text = box2text.replace("\n", "")
        if(" " in box1text and " " in box2text):
            box1text = box1text.replace(" ", "")
            box2text = box2text.replace(" ", "")
        if(box1text == box2text):
            t.color("green", "green")
            mouse.position = buttonPos
            mouse.press(Button.left)
            mouse.release(Button.left)
            time.sleep(2)
        else:
            t.color("red", "red")
        turtle.update()
        print(box1text)
        print(box2text)

# Calling the function
main()
