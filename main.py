import tkinter as tk
from tkinter import ttk
from comtypes.client import CreateObject
from comtypes.gen import SpeechLib
import logging
from tkinter import messagebox as mbx
"""sets up initial conversion"""
class MainApplication(tk.Frame):
    def __init__(self,*args, **kwargs):
        tk.Frame.__init__(self, *args, **kwargs)
        """application configuration"""
        menu = tk.Menu()
        root.grid_rowconfigure(2, weight=2)
       
        root.config(menu=menu)
        root.resizable(0,0)
        root.title("my converter")
        root.geometry("200x200")
        menu.add_command(label="exit", command=root.quit)   
        root.configure(background="green")
        menu.configure(background="white")
        self.Linput = ttk.Label(root, text="input", width=20)
        self.Linput.pack()
        self.inputentry = ttk.Entry(root, width=20)
        self.inputentry.pack()
        self.Loutput = ttk.Label(text="output", width=20)
        self.Loutput.grid_rowconfigure(4, weight=2)
        self.Loutput.pack()
        self.OutputEntry = ttk.Entry(width=20)
        self.OutputEntry.pack()
        self.Convert = ttk.Button(text="convert", command=self.ConvertertoAudio)
        self.Convert.pack(side="top")

    def ConvertertoAudio(self):
        try: 
            self.engine = CreateObject("SAPI.SpVoice")
            self.stream = CreateObject("SAPI.SpFileStream")
            self.infile = self.inputentry.get()
            self.outfile = self.OutputEntry.get()
            self.stream.Open(self.outfile, SpeechLib.SSFMCreateForWrite)
            self.engine.AudioOutputStream = self.stream
            with open(self.infile, "r") as text:
                self.content = text.read()
                self.engine.speak(self.content)
                self.stream.Close()
                mbx.showinfo("", f" conversion of {self.infile} {self.outfile} (self.infile, self.outfile)")

        except:
            mbx.showwarning("error", "the file you destination is already converted")
          

root =  tk.Tk()
app = MainApplication()
root.mainloop()








