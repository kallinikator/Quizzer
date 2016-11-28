import random
import tkinter

import xlwt
import xlrd

class Quiz(object):

    def __init__(self, file):
        # Opens a file with the questions
        self.workbook = xlrd.open_workbook(file).sheet_by_index(0)

        # Opens an xls file for the results
        self.resultfile = xlwt.Workbook(encoding="utf-8")
        self.resultsheet = self.resultfile.add_sheet("Ergebnisse", cell_overwrite_ok=True)
        
        self.resultsheet.write(0, 0, "Question")
        self.resultsheet.write(0, 1, "Correct answer")
        self.resultsheet.write(0, 0, "Answer")

        # The score of the player
        self.score = 0


    # Pick Questions randomly and ask the questions
    def ask_questions(self):
        for row in range(1, self.workbook.nrows): # With random.shuffle, you can sort the list of questions
            self.quest(row)
        self.save_result()
            
    
    # Ask the question and store the answer directly in the .xls
    def quest(self, row):
        print(self.workbook.cell_value(row ,0))
        answer = input("1:\t{}\n2:\t{}\n3:\t{}".format(
            self.workbook.cell_value(row ,1),
            self.workbook.cell_value(row ,2),
            self.workbook.cell_value(row ,3),
            ))

        # Store the results in the resultfile
        self.resultsheet.write(row, 0, self.workbook.cell_value(row ,0))
        self.resultsheet.write(row, 1, self.workbook.cell_value(row ,4))
        self.resultsheet.write(row, 2, answer)

        # Validate the result
        if answer == str(int(self.workbook.cell_value(row ,4))): # Ugly solution...
            self.score += 1
            print("richtig!")
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
    a = Quiz("Quiz.xls")
    a.ask_questions()
