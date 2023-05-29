#!/usr/bin/env python3

from tkinter import Tk
from application import Application

# email = "yifefa3387@introace.com"
# username = "InternshalaTask"
# password = "Task@Internshala"

# path = "sellers.xlsx"
# message = "FORMATTABLE_MESSAGE_TEMPALTE"
# url_column = 3
# msg_column = url_column - 1


def main():
    root = Tk()

    # root.title("Techneik Etsy Messenger")
    # root.geometry('350x200')
    
    # filepath_entry = Entry(root, bd =5)
    # filepath_entry.pack(fill=BOTH, )
    # load_button = Button(root, text="Select", command=lambda: setExcelFile(filepath_entry))
    # load_button.pack()
    app = Application(root)

    root.mainloop()

     
if __name__ == "__main__":
    main()

