import time
import tkinter

import xlwt
import xlrd


class tkGUI(object):

    def __init__(self):
        self.master = tkinter.Tk()
        self.master.title("Börsenführerschein")

        self.answer = tkinter.IntVar()
        self.result = None
        
    def next_click(self):
        # Closes the actual frame to get the next one running
        self.frame.quit()

        
    def radioclick(self):
        # Stores the result of the radiobuttons
        self.result = self.answer.get()

    def quizscreen(self, question):
        # Creates a frame containing the radionbuttuns and stuff
        self.frame = tkinter.Frame(master=self.master)
        self.frame.grid(row=0, column=0, rowspan=10, columnspan=10)

        # The Question
        tkinter.Label(self.frame, text=question[0]).grid(row=1, column=1, columnspan=5)

        # The Radiobuttons
        chooser1 = tkinter.Radiobutton(self.frame, text=question[1], value=1, variable=self.answer, command=self.radioclick)
        chooser1.grid(row=2, column=0)

        chooser2 = tkinter.Radiobutton(self.frame, text=question[2], value=2, variable=self.answer, command=self.radioclick)
        chooser2.grid(row=3, column=0)

        chooser3 = tkinter.Radiobutton(self.frame, text=question[3], value=3, variable=self.answer, command=self.radioclick)
        chooser3.grid(row=4, column=0)

        # The Submitbutton
        submit_button = tkinter.Button(self.frame,
                               text="Nächste Frage",
                               command=self.next_click)
        submit_button.grid(row=5, column=1)

        self.frame.mainloop()




class Quiz(object):

    def __init__(self, file, gui):
        # Opens a file with the questions
        self.workbook = xlrd.open_workbook(file).sheet_by_index(0)
        self.gui = gui

        # Opens an xls file for the results
        self.resultfile = xlwt.Workbook(encoding="utf-8")
        self.resultsheet = self.resultfile.add_sheet("Ergebnisse", cell_overwrite_ok=True)
        
        self.resultsheet.write(0, 0, "Frage")
        self.resultsheet.write(0, 1, "Richtige Antwort")
        self.resultsheet.write(0, 0, "Antwort")

        # The score of the player
        self.score = 0


    # Pick Questions randomly and ask the questions
    def ask_questions(self):
        for row in range(1, self.workbook.nrows): # With random.shuffle, you can sort the list of questions
            question = self.workbook.row_values(row)
            self.quest(question, row) # There is a iterator function to be implemented.
        self.save_result()
            
    
    # Ask the question and store the answer directly in the .xls
    def quest(self, question, row):

        self.gui.quizscreen(question)
  
        # Store the results in the resultfile
        self.resultsheet.write(row, 0, question[0])
        self.resultsheet.write(row, 1, question[4])
        self.resultsheet.write(row, 2, self.gui.result)

        # Validate the result
        if self.gui.result == int(question[4]): # Ugly solution...
            self.score += 1
            print("richtig!") # Only for debugging
        else:
            print("falsch")


    # Saves the result in a file and prints the score underneath
    def save_result(self):
        self.resultsheet.write(62, 0, self.score)
        if self.score >= 50:
            result = "Mit bravour bestanden!"
        elif self.score >= 30:
            result = "Bestanden!"
        else:
            result = "Nicht bestanden!"
        self.resultsheet.write(62, 1, result)
        self.resultfile.save("result.xls")





    # Create a GUI




if __name__ == "__main__":
    gui = tkGUI()
    a = Quiz("Quiz.xls", gui)
    a.ask_questions()

