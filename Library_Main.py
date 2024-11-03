from openpyxl import Workbook
import openpyxl
import os
from prettytable import PrettyTable

if os.path.isfile("Bookkeeper.xlsx"):
      """Check if the file exists, open the file if it does"""
      workbook = openpyxl.load_workbook("Bookkeeper.xlsx")
      sheet = workbook.active
      print("test")
else:
      """If the file doesn't exist, create a new one and initialize the top row of data"""
      workbook = Workbook()
      workbook.save(filename="Bookkeeper.xlsx")
      workbook = openpyxl.load_workbook("Bookkeeper.xlsx")
      sheet = workbook.active
      a1 = sheet.cell(row=1, column=1)
      a1.value = "Author"
      a2 = sheet.cell(row=1, column=2)
      a2.value = "Title"
      a3 = sheet.cell(row=1, column=3)
      a3.value = "Rating"
      a4 = sheet.cell(row=1, column=4)
      a4.value = "Quote"
      workbook.save("Bookkeeper.xlsx")

def add_Book():
      print("\n")
      print("Please ensure that you enter these correctly, as editing is not currently implemented\n")
      title = input("Enter the name of the book: ")
      author = input("What is the name of the author: ")
      rating = input("Enter a rating for this book between 1-10: ")
      quote = input(
            "If you'd like, enter a favorite quote here, or enter 0 to add your book and return to the main menu: ")
      if quote == 0 or None:
            quote = ""

      entry = (author, title, rating, quote)

      sheet.append(entry)
      workbook.save("BookKeeper.xlsx")
      print("Your book has been added!")
      print("\n")
      return()

def see_Books():
      print("/n")

      """Create the table to be used"""
      myTable = PrettyTable(["Author", "Title", "Rating", "Quotes"])
      row = str(sheet.max_row)
      column = sheet.max_column
      cell = sheet["A1":"D"+row]
      data = [list(elem) for elem in cell]
      cell = []
      row = 2
      column = 1

      """Loop throught the list of cells and add the rows to the table"""
      for i in data:
            for n in range(1, 5):
                  cell.append(sheet.cell(row, column = n))
            if column < 4:
                  column += 1
            row += 1
            myTable.add_row([cell[0].value,cell[1].value,cell[2].value,cell[3].value])
            cell = []
      print(myTable)

def main_Menu():
      selection = ""
      while selection != 0:
            print("Welcome to Bookkeeper, this program will help you track the books you've read by storing them in an excel document\n\n"
                  "Enter 1 to add a new book\n"
                  "Enter 2 to see the books in your collection\n"
                  "Enter 3 add your favorite quotes to a book\n"
                  "Enter 0 to exit the program\n"
                  "Please only enter numbers between 0-3\n")
            selection = input("Enter your selection here: ")

            if selection == "0":
                  exit()
            if selection == "1":
                  add_Book()
            elif selection == "2":
                  see_Books()
      exit()

if __name__ == "__main__":
      main_Menu()


