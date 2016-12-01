import re
import datetime
import tkinter

import xlwt
import xlrd


class tkGUI(object):
    """
    Provides an interface for getting the questions asked. It is initiated as an empty widget.
    The function quizscreen opens a frame on the widget with the question.
    """

    def __init__(self):
        self.master = tkinter.Tk()
        self.master.title("Börsenführerschein")
        self.master.attributes("-fullscreen", True)

        self.master.update()

        self.width = self.master.winfo_width()
        self.height = self.master.winfo_height()

        # Timer stuff
        self.done_time = datetime.datetime.now() + datetime.timedelta(seconds=1800) # half hour

        
    def next_click(self):
        # Closes the actual frame to get the next one running
        self.next = 1
        self.frame.quit()
 

    def previous_click(self):
        # Closes the actual frame to get the next one running
        self.next = -1
        self.frame.quit()

        
    def radioclick(self):
        # Stores the result of the radiobuttons
        self.result = self.answer.get()


    def update_clock(self):
        # Let the remaining time tick down
        self.elapsed = self.done_time - datetime.datetime.now()
        m, s = self.elapsed.seconds//60, self.elapsed.seconds%60
        self.timer.configure(text="{}:{}".format(m,s))
        self.frame.after(1000, self.update_clock)


    def quizscreen(self, question, num_quest, *args):
        # Creates a frame containing the radionbuttuns and stuff
        self.answer = tkinter.IntVar()
        self.result = None
        cells = num_quest*2+5
        
        self.frame = tkinter.Frame(master=self.master, background="#D5E88F")
        self.frame.place(width=self.width, height=self.height)

        # The Question
        tkinter.Label(
            self.frame,
            text=question[0],
            background="White"
            ).place(y=self.height/cells,
                    width=self.width,
                    height=self.height/cells)
        
        # The Radiobuttons
        position = 3
        for possible_answer in range(1, num_quest+1):
            tkinter.Radiobutton(
                self.frame,
                text=question[possible_answer],
                value=possible_answer,
                variable=self.answer,
                command=self.radioclick
                ).place(x=self.width/10,
                        y=position*self.height/cells,
                        width=8*self.width/10,
                        height=self.height/cells)
            position += 2

        # The Nextbutton
        submit_button = tkinter.Button(self.frame,
                               text="Nächste Frage",
                               command=self.next_click)
        submit_button.place(x=self.width/10,
                            y=(cells-2)*self.height/cells,
                            width=self.width/5,
                            height=self.height/cells)

        if not "start" in args:
            # The Previousbutton
            submit_button = tkinter.Button(self.frame,
                                   text="Vorherige Frage",
                                   command=self.previous_click)
            submit_button.place(x=4*self.width/10,
                                y=(cells-2)*self.height/cells,
                                width=self.width/5,
                                height=self.height/cells)

        # The timer
        self.timer = tkinter.Label(self.frame,
                                   text="")
        self.timer.place(
            x=7*self.width/10,
            y=(cells-2)*self.height/cells,
            width=self.width/5,
            height=self.height/cells)

        self.update_clock()
        self.frame.mainloop()


    def resultscreen(self, score, result):
        # Creates a final screen that tells your result
        self.frame = tkinter.Frame(master=self.master, background="#D5E88F")
        self.frame.place(width=self.width, height=self.height)

        tkinter.Label(
            self.frame,
            text="Du hast {} von 60 Punkten erreicht!".format(str(score)),
            background="White"
            ).place(y=self.height/7,
                    width=self.width,
                    height=self.height/7)
        
        tkinter.Label(
            self.frame,
            text=result,
            background="White"
            ).place(y=3*self.height/7,
                    width=self.width,
                    height=2*self.height/7)

        self.frame.mainloop()


class Quiz(object):
    """
    This class manages the input and output in excel-files and the asking of the questions.
    """

    def __init__(self, file, gui):
        # Opens a file with the questions
        self.workbook = xlrd.open_workbook(file).sheet_by_index(0)
        self.gui = gui

        # Opens an xls file for the results
        self.resultfile = xlwt.Workbook(encoding="utf-8")
        self.resultsheet = self.resultfile.add_sheet("Ergebnisse", cell_overwrite_ok=True)
        
        self.resultsheet.write(0, 0, "Frage")
        self.resultsheet.write(0, 1, "Richtige Antwort")
        self.resultsheet.write(0, 2, "Antwort")

        # The score of the player
        self.score = 0


    # Pick Questions randomly and ask the questions
    def ask_questions(self):
        #self.gui.startscreen()
        
        self.pointer = 1
        while self.pointer >= 1 and self.pointer < 61:
            question = self.workbook.row_values(self.pointer)
            self.quest(question, self.pointer)
            self.pointer += self.gui.next

        self.save_result()

        self.gui.resultscreen(self.score, self.result)
            
    
    # Ask the question and store the answer directly in the .xls
    def quest(self, question, row):
        if self.pointer == 1:
            self.gui.quizscreen(question, 5, "start")
        else:    
            self.gui.quizscreen(question, 5)
  
        # Store the results in the resultfile
        self.resultsheet.write(row, 0, question[0])
        self.resultsheet.write(row, 1, question[6])
        self.resultsheet.write(row, 2, self.gui.result)

        # Validate the result
        if self.gui.result == int(question[6]):
            self.score += 1


    # Saves the result in a file and prints the score underneath
    def save_result(self):
        self.resultsheet.write(62, 0, self.score)
        if self.score >= 50:
            self.result = "Mit Bravour bestanden!"
        elif self.score >= 30:
            self.result = "Bestanden!"
        else:
            self.result = "Nicht bestanden!"
        self.resultsheet.write(62, 1, self.result)
        self.resultfile.save("result.xls")


if __name__ == "__main__":
    gui = tkGUI()
    a = Quiz("Quiz.xls", gui)
    a.ask_questions()

