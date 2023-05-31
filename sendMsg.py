#!/usr/bin/env python3

from tkinter import Tk
from msg_sender import MsgSender


# email = "nixiya4408@soremap.com"
# username = "InternshalaTask"
# password = "Task@Internshala"


def main():
    filepath = input("Excel Filepath: ")
    msg_template = input("Message Template: ")
    start_num = input("Start Num: ")
    num_msg = input("Num Msg: ")
    msg_sender = MsgSender()

     
if __name__ == "__main__":
    main()

