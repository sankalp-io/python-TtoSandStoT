# Title Spectre Envy
# Module Import
import os  # interact with os of the system
import tkinter as tk  # gui app python standard library
from tkinter import filedialog  # use to create file and directory of the system
# use to display message boxes in the python appliances
from tkinter import messagebox
import speech_recognition as sr  # library to perform speech recognition
import pyttsx3  # text to speech module
# it is library that enables programmer to use Windows API to create Gui based programmes
from win32com.client import constants, Dispatch
# function to translate the speech to text in python


def SpeakText(command):

    engine = pyttsx3.init()
    engine.say(command)  # intialize the engine command to start as argument
    engine.runAndWait()  # wait for user to terminate the conversation


Working_Dir = os.getcwd()  # use to get the current working directory

r = sr.Recognizer()  # uses the machine mic to listen and analyze what has been said
# which microphone is been used while intitializing the conversation
mic = sr.Microphone(device_index=14)

speaker = Dispatch("SAPI.SpVoice")  # text to speech recognition engine


class Application(tk.Frame):  # it is used to organize and group  widgets
    def __init__(self, master=None):
        super().__init__(master=master)
        self.master = master  # it represents the instance of the class
        self.pack()
        self.Main_Frame()

    def Main_Frame(self):  # Driver code of the Dialog Box
        self.Delete_Frame()

        self.Frame_1 = tk.Frame(self)
        self.Frame_1.config(width=400, height=100)  # Size of the dialog box
        self.Frame_1.grid(row=0, column=0, columnspan=2)

        self.Label_1 = tk.Label(self.Frame_1)
        # Main dialog box
        self.Label_1['text'] = 'Convert  Text to Audio And Audio to Text Speech using Python'
        self.Label_1.grid(row=0, column=0, pady=30)

        self.Label_2 = tk.Label(self.Frame_1)
        # statement displayed on the main dialog box
        self.Label_2['text'] = 'Requires an Active Internet Connection'
        self.Label_2.grid(row=1, column=0, pady=10, padx=100)

        self.SpeehToText = tk.Button(self, bg='#FFFF00', fg='black', font=(
            "Times new roman", 14, 'bold'))  # customization of the buttons
        # text written on the buttons
        self.SpeehToText['text'] = 'Speech to Text'
        self.SpeehToText['command'] = self.SpeechToText
        self.SpeehToText.grid(row=1, column=0, pady=80, padx=60)

        self.TextTo_Speech = tk.Button(
            self, bg='#FFFF00', fg='black', font=("Times new roman", 14, 'bold'))
        self.TextTo_Speech['text'] = 'Text to Speech'
        self.TextTo_Speech['command'] = self.TextToSpeech
        self.TextTo_Speech.grid(row=1, column=1, pady=60, padx=60)

    def Delete_Frame(self):
        for widgets in self.winfo_children():
            widgets.destroy()

    def SpeechToText(self):
        self.Delete_Frame()

        self.Listen = tk.Button(self, bg='#FFFF00', fg='black', font=(
            "Times new roman", 18, 'bold'))
        self.Listen['text'] = 'Listen'
        self.Listen['command'] = self.Audio_Recognizer
        self.Listen.grid(row=0, column=0, pady=40)

        self.Back = tk.Button(self, bg='brown', fg='black',
                              font=("Times new roman", 18, 'bold'))
        self.Back['text'] = 'back'
        self.Back['command'] = self.Main_Frame
        self.Back.grid(row=0, column=2)

        self.text = tk.Text(self)
        self.text.configure(width=48, height=10)
        self.text.grid(row=1, column=0, columnspan=3)

    def TextToSpeech(self):
        self.Delete_Frame()
        self.scroll = tk.Scrollbar(self, orient=tk.VERTICAL)
        self.scroll.grid(row=0, column=4, sticky='ns', padx=0)

        self.text = tk.Text(self)
        self.text.configure(width=44, height=12)
        self.text.grid(row=0, column=0, columnspan=3)
        self.text.config(yscrollcommand=self.scroll.set)
        self.scroll.config(command=self.text.yview)

        self.GET_Audio = tk.Button(
            self, bg='#FFFF00', fg='black', font=("Times new roman", 17, 'bold'))
        self.GET_Audio['text'] = 'Get Audio'
        self.GET_Audio['command'] = self.Convert_TextToSpeech
        self.GET_Audio.grid(row=1, column=0, pady=50)

        self.read_file = tk.Button(
            self, bg='#FFFF00', fg='black', font=("Times new roman", 17, 'bold'))
        self.read_file['text'] = 'Read file'
        self.read_file['command'] = self.Read_File
        self.read_file.grid(row=1, column=1)

        self.Clear_Frame = tk.Button(
            self, bg='#FFFF00', fg='black', font=("Times new roman", 17, 'bold'))
        self.Clear_Frame['text'] = 'Clear'
        self.Clear_Frame['command'] = self.Clear_TextBook
        self.Clear_Frame.grid(row=1, column=2)

        self.Back = tk.Button(self, bg='brown', fg='black',
                              font=("Times new roman", 17, 'bold'))
        self.Back['text'] = 'back'
        self.Back['command'] = self.Main_Frame
        self.Back.grid(row=1, column=3)

    def Audio_Recognizer(self):
        self.Clear_TextBook()
        try:
            with sr.Microphone() as source2:
                r.adjust_for_ambient_noise(source2, duration=0.2)
                audio2 = r.listen(source2)
                MyText = r.recognize_google(audio2)
                MyText = MyText.lower()
                self.text.insert('1.0', MyText)
                SpeakText(MyText)

        except:
            self.text.insert('1.0', 'No internet connection.....')

    def Convert_TextToSpeech(self):
        self.msg = self.text.get(1.0, tk.END)
        if self.msg.strip('\n') != '':
            speaker.speak(self.msg)
        else:
            speaker.speak('Write some message first.....')

    def Read_File(self):
        self.filename = filedialog.askopenfilename(initialdir=Working_Dir)

        if (self.filename == '') or (not self.filename.endswith('.txt')):
            messagebox.showerror('Can not load file.....',
                                 'Choose a text file to read....')
        else:
            with open(self.filename) as f:
                text = f.read()
                self.Clear_TextBook()
                self.text.insert('1.0', text)

    def Clear_TextBook(self):
        self.text.delete(1.0, tk.END)


root = tk.Tk()
root.geometry('500x300')
root.wm_title('Text to Speech and Speech to Text converter')

app = Application(master=root)
app['bg'] = '#FFFFFF'
app.mainloop()
