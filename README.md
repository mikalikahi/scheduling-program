# scheduling-program
 Python English teacher schedule management -- from an Excel file, search for available teachers, send automated scheduling emails to all teachers, and look up teachers' monthly schedules

Ideas for updates and revisions:
1.) Open excel files using “with.” ie. (with open("test.txt") as file_handler:)
2.) Make the window size adjustable and make a scrollbar
3.) Raising exceptions in GUI for input errors ie. invalid emails etc.
4.) Set up a status bar which will pop up at the top level of the GUI and run while the emails are being sent -- trying using threading to do this
5.) Find a way to not repeat the GUI variables – or at least minimize their repetition without making them global variables -- try doing this by making subclasses that will inherit the variables from the parent class
6.) Instead of using datetime.datetime just use datetime.date to avoid having the unnecessary digits in the date row in Excel
