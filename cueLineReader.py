#Written by Stew Towle
#Description: This code is designed to help Original Practice Shakespeare Festival
# actors practice their lines by reading the cues to them.
#Basic edition from Sept 19, 2021

#Notes: I would like to make a pop-up window for common issues such as:
#1) If a line or cue is not loading right sometimes if you go into the document and just hit
# enter at the end of cue lines (sometimes in the cue scripts there isn't a real line break after cue)

import docx
import sys
import gtts
import os
#import pyttsx3
from tkinter import *
from tkinter import filedialog
import tkinter.scrolledtext as scrolledtext

from playsound import playsound

class CueLineReader:

    def __init__(self):
        self.window = Tk()
        # greeting = Label(text="Cue Line Practice Program")
        # greeting.pack()
        self.window.geometry("666x400")
        # Set window background color
        self.window.config(background="gray")
        self.window.title("Cue Line Rehearser")
        self.current_cue = 0
        self.cue_lines = list()
        self.char_lines = list()
        self.label_file_explorer = Label(self.window,
                                    text="Browse to find script as a .docx file then press 'Load Cues' \n(if it is a .doc file "
                                         "open it with a document editor and use save as to save it as .docx)",
                                    #width=20, height=4,
                                    fg="blue")
        self.cue_display = Text(self.window, width=25, height=3, bg="yellow", wrap="word")
        self.line_display = scrolledtext.ScrolledText(self.window, width=70, height=12, bg="violet", wrap="word", font=("Courier", 15))
        #self.engine = pyttsx3.init()


    def test_reader(self, name):
        """
        Given a file path returns a list of each paragraph in the file
        file must be a docx file.
        """
        doc = docx.Document(name)
        return doc.paragraphs
        # with open(name, 'rb') as stream:
        #     while stream.readable():
        #         print(stream.readline())

    def generate_cue_lines(self, list_paras):
        """
        Generates a list of strings that are the cue lines
        :param list_paras: A list of 'paragraphs' from the original docx file
        :return: just the words in the cue lines (which are the lines beginning in a long series of periods)
        """
        cue_lines = list()
        for line in list_paras:
            if "....." in line:
                cue_lines.append(line.lstrip().lstrip(".").lstrip())
        return cue_lines

    def generate_lines(self, the_cues, list_paras):
        """Given a list of cue lines and the list of all paragraphs from the document
        returns a list of the character's lines where each line has the same index
        of the cue that prompts it."""
        cue_counter = 0
        found_lines = False
        lines_list = list()
        for para in list_paras:
            if found_lines:
                if "......" in para:
                    cue_counter += 1
                    found_lines = False
                    lines_list.append(current_line)
                else:
                    current_line += para + "\n"
            if the_cues[cue_counter] in para:
                found_lines = True
                current_line = ""
        if found_lines:
            lines_list.append(current_line)
        if len(lines_list) < len(self.cue_lines):
            # print(self.cue_lines[-1])
            # print(list_paras[-4])
            # print(list_paras[-3])
            # print(list_paras[-2])
            # print(list_paras[-1])
            lines_list.append(list_paras[-1])
        while len(lines_list) < len(self.cue_lines):
            lines_list.append("something went wrong and you have more cues than lines.  Maybe make sure there is"
                              "an actual line break after each cue (hit enter in the document at the end of the cue"
                              "and then re-save)")
        #print(lines_list)
        return lines_list


    def update_window(self):
        """
        Updates the text boxes for displaying the current Cue and Line
        :return: None
        """
        self.cue_display.delete("1.0", "end")
        self.line_display.delete("1.0", "end")

        if self.current_cue > 0:
            self.cue_display.insert("end", self.cue_lines[self.current_cue -1])
            self.line_display.insert("end", self.char_lines[self.current_cue -1])
        elif len(self.cue_lines) > 0:
            self.cue_display.insert("end", self.cue_lines[0])
            self.line_display.insert("end", self.char_lines[0])
        else:
            self.cue_display.insert("end", "Cues will appear here.")
            self.line_display.insert("end", "Lines will appear here.")


        count_tab = "\n\n (cue " + str(self.current_cue) + "/" + str(len(self.cue_lines)) + ")"
        self.line_display.insert("end", count_tab)
        self.cue_display.update_idletasks()
        self.line_display.update_idletasks()



    def browseFiles(self):
        filename = filedialog.askopenfilename(initialdir="/",
                                              title="Select a File",
                                              filetypes=(("Word Doc",
                                                          "*.docx"),
                                                         ("all files",
                                                          "*.*")))

        # Change label contents
        self.label_file_explorer.configure(text="File Opened: " + filename + \
                                                "\nPress 'Load Cues' then use arrow keys or clickable button to play cues.")


    def load_cues(self, file_name = None):
        """Gets the file name from the tkinter window and loads the cue_lines and char_lines
        with the appropriate information."""
        try:
            if file_name is None:
                file_name = self.label_file_explorer.cget("text").lstrip("File Opened: ").replace("\nPress 'Load Cues' then use arrow keys or clickable button to play cues.","")
            #print("loaded", file_name)
            paras = self.test_reader(file_name)
            lines = [paras[i].text for i in range(len(paras))]

            self.cue_lines = self.generate_cue_lines(lines)
            self.char_lines = self.generate_lines(self.cue_lines, lines)
            self.current_cue = 0
            self.update_window()
        except:
            self.cue_lines = ["File did not load"]
            self.char_lines = ["Didst thou select and then load a .docx file?"]
            self.current_cue = 0
            self.update_window()
        # for line in cue_lines:
        #     print(line)
        #     print("-----")

    def prep_cue(self, text):
        """Given a potential cue line removes characters that will make mac os built in say function
        throw an error"""
        result = text.replace("(","")
        result = result.replace(")","")
        result = result.replace("\n","")
        result = result.replace("\t","")
        result = result.replace("'","")

        return result

    def say_text(self, text):
        """Given text to say, says it with built in os say function"""
        # if sys.platform.startswith("darw"):
        #     os.system("/usr/bin/say" + " " + text)
        # elif sys.platform.startswith("win"):
        tts = gtts.gTTS(text)
        tts.save("utter.mp3")
        playsound("utter.mp3")


    def play_next(self, catcher=None, catcher2=None):

        cues_length = len(self.cue_lines)
        if self.current_cue < cues_length:
            text = self.prep_cue(self.cue_lines[self.current_cue])
            self.current_cue += 1
        else:
            text = "FINIS"


        if self.current_cue > cues_length and not self.current_cue == 0:
            self.current_cue = cues_length
        self.update_window()
        self.window.after(500, self.say_text, text)
        # self.engine.say(text)
        # self.engine.runAndWait()

    def replay(self, catcher=None, catcher2=None):
        """Method for replaying current cue."""
        if self.current_cue > 0 and  self.current_cue < len(self.cue_lines) + 1:
            self.current_cue -= 1
            self.play_next()
        else:
            self.update_window()
            self.window.after(500, self.say_text, "Beginning of play.")


    def play_previous(self, catcher=None, catcher2=None):
        """Plays the cue_line before the one most recently played.
        Plays the words 'Beginning of Play' if there is no line before the current"""
        if self.current_cue > 1 and self.current_cue <= len(self.cue_lines):
            self.current_cue -= 2
            self.play_next()
        else:
            self.update_window()
            self.window.after(500, self.say_text, "Beginning of play.")

    def jump_forward(self, catcher=None, catcher2=None):
        """Jumps 10 cues forward (or to end of play if there are not ten more cues)"""
        if self.current_cue + 10 >= len(self.cue_lines):
            self.current_cue = len(self.cue_lines)
        else:
            self.current_cue += 10
        self.update_window()

    def exit(self, catch1=None, catch2=None):
        self.window.destroy()

    def jump_back(self, catcher=None, catcher2=None):
        """Jumps 10 cues backward (or to beginning of play if there are not ten more cues)"""
        if self.current_cue - 10 <= 0:
            self.current_cue = 0
        else:
            self.current_cue -= 10
        self.update_window()

    def run_gui(self):

        button_explore = Button(self.window,
                                text="Browse Files",
                                command=self.browseFiles, fg="black")

        button_exit = Button(self.window,
                             text="Exit",
                             command=self.exit, fg="black")

        button_load = Button(self.window, text="Load Cues", command=self.load_cues, fg="black")

        button_next = Button(self.window, text="Next Cue (right arrow key)", command=self.play_next, fg="black")
        button_previous = Button(self.window, text="Previous Cue (left arrow key)", command=self.play_previous, fg="black")
        button_replay = Button(self.window, text="Replay Cue (spacebar)", command=self.replay, fg="black")
        button_jump_forward = Button(self.window, text="Jump Ahead (10 cues)", command=self.jump_forward, fg="black")
        button_jump_back = Button(self.window, text="Jump Back (10 cues)", command=self.jump_back, fg="black")

        #Here I bind the arrow keys to run cues
        self.window.bind('<Left>', self.play_previous)
        self.window.bind('<space>', self.replay)
        self.window.bind('<Right>', self.play_next)

        # Grid method is chosen for placing
        # the widgets at respective positions
        # in a table like structure by
        # specifying rows and columns
        self.label_file_explorer.grid(column=1, columnspan=3, row=1)

        button_explore.grid(column=1, row=2)
        button_exit.grid(column=3, row=2)
        button_load.grid(column=2, row=2)
        button_next.grid(column=3, row=4)
        button_previous.grid(column=1, row=4)
        button_replay.grid(column=2, row=4)
        button_jump_back.grid(column=1, row=5)
        button_jump_forward.grid(column=3, row=5)
        self.cue_display.grid(column=3, row=6)
        self.line_display.grid(column=1, columnspan=3, row=7)

        # Let the window wait for any events
        self.update_window()
        self.window.mainloop()


if __name__ == '__main__':

    this_program = CueLineReader()
    this_program.run_gui()






