from Scripts.App import App
from Scripts.Date import Date, Visitors
from Scripts.Student import Student
import os
from pathlib import Path

os.chdir("TestingCenter")
dir = os.path.abspath(os.getcwd())
data_path = Path(dir + "\\Data\\Visitors.xlsx")
courses_path = Path(dir + "\\Data\\Courses.txt")

if __name__ == "__main__":
    app = App(data_path, courses_path, Date, Visitors, Student)
    app.mainloop()

