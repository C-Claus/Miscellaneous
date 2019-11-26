import pywinauto
import pyautogui
import time
import math
from math import pi
from pywinauto.application import Application

def start_paint():
    app = Application().start("mspaint.exe")
    
def move_mouse(x, y):
    pyautogui.moveTo(x,y,duration=0.5)
    
def drag_circle(radius, integers):

    coordinates = [(round(math.cos(2*pi/integers*x)*radius), round(math.sin(2*pi/integers*x)*radius)) for x in range(0, integers+1)]
    
    for x,y in coordinates:
        print (x,y)
        pyautogui.dragRel((x,y), 0, duration=0.1)
    
   
def drag_half_circle(radius, integers): 
    
    coordinates = [(round(math.cos(2*pi/integers*x)*radius), round(math.sin(2*pi/integers*x)*radius)) for x in range(0, integers+1)]
 
    x = (len(coordinates))

    quarter_list = ((len(coordinates)))/4
    
    del coordinates[0:round(quarter_list)]
    del coordinates[2*round(quarter_list):x]
   
    for x,y in coordinates:
        print (x,y)
        pyautogui.dragRel((x,y), 0, duration=0.1)
        
        
#start paint
start_paint()
time.sleep(2)
move_mouse(x=500, y=200)
time.sleep(2)

#make eyes and face
drag_circle(radius=20, integers=50)
move_mouse(x=550, y=250)
drag_circle(radius=5, integers=50)
move_mouse(x=450, y=250)
drag_circle(radius=5, integers=50)

#make mouth
move_mouse(x=600, y=400)
drag_half_circle(radius=10, integers=50)

#make pupils 
move_mouse(x=430, y=270)
drag_circle(radius=2, integers=50)

move_mouse(x=530, y=270)
drag_circle(radius=2, integers=50)



