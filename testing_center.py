import sys
from src.main.python.App import App
from src.test.python.test_app import run_tests as run_app_tests
from src.test.python.test_date import run_tests as run_date_tests
from src.test.python.test_student import run_tests as run_student_tests


if __name__ == "__main__":
    arg1=""
    if len(sys.argv) > 1:
        arg1 = sys.argv[1]

    if arg1=="test_app" or arg1=="ta":
        run_app_tests()
    elif arg1=="test_date" or arg1=="td":
        run_date_tests()
    elif arg1=="test_student" or arg1=="ts":
        run_student_tests()
    else:
        app = App()
        app.mainloop()

