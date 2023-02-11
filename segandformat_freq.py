from tkinter import *
from tkinter import filedialog
from GUI import *
import ctypes
import platform
if platform.system=='Windows':
    ctypes.windll.shcore.SetProcessDpiAwareness(1)  
    
          
if __name__ == '__main__':
    root = Tk()
    root.title('Seg/Freq')
    app = GUIDemo(master=root,Freq=True)
    app.mainloop()
