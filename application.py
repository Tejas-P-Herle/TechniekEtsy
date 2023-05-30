
from tkinter import *
from tkinter.filedialog import askopenfilename

from os import path

from msg_sender import MsgSender


class Application(Frame):
    def __init__(self, root):
        super().__init__()
        
        self.row = 0
        self.root = root
        self.status_label = None
        self.initUI()

    def initUI(self):

        # self.master.title("Windows")
        self.master.title("Techneik Etsy Messenger")
        # self.master.geometry('350x200')
    
        self.pack(fill=BOTH, expand=True)

        # self.columnconfigure(0, weight=1)
        # self.columnconfigure(3, pad=7)
        # self.rowconfigure(3, weight=1)
        # self.rowconfigure(5, pad=7)

        col1_ewidth = 16
        
        heading_label = self.add_label("Etsy Auto-Messenger", columnspan=2, font=24)
        status_label = self.add_label("Status: Loaded", columnspan=2)
        self.status_label = status_label

        sel_label = self.add_label("Select Excel File:")
        filepath_entry = self.add_entry()
        load_button = self.add_button("Select", lambda: self.setExcelFile(filepath_entry))
        url_col_label = self.add_label("URL Column", same_row=True)
        url_col_entry = self.add_entry(width=col1_ewidth, column=1,
            num_only=True, default=2)
        msg_label = self.add_label("Enter Message:", columnspan=2)
        
        # message_box = Text(self, height = 5, width = 52)
        # message_box.grid(row=3, column=0, columnspan=2, padx=5, pady=5)
        message_box = self.add_text()

        start_label = self.add_label("Starting URL Offset", same_row=True)
        start_num_entry = self.add_entry(width=col1_ewidth, column=1,
            num_only=True, default=1)
        num_label = self.add_label("Number of Messages", same_row=True)
        num_msg_entry = self.add_entry(width=col1_ewidth, column=1,
            num_only=True, default=3)

        username_label = self.add_label("Email ID(If not logged in)",
            same_row=True)
        username_entry = self.add_entry(width=col1_ewidth, column=1,
            default="yifefa3387@introace.com")
        password_label = self.add_label("Password(If not logged in)",
            same_row=True)
        password_entry = self.add_entry(width=col1_ewidth, column=1, show="*",
            default="Task@Internshala")

        input_fields = {'filepath': filepath_entry, 'url_col': url_col_entry,
            'msg_template': message_box, 'start_num': start_num_entry,
            'num_msg': num_msg_entry, 'username': username_entry,
            'password': password_entry}

        self.row += 1
        send_msg_btn = self.add_button("Send Messages",
            lambda: self.start_sending(input_fields),
            column=0, columnspan=2, padx=5, pady=5)

    def add_label(self, label_text, column=0, same_row=False,
                  font=None, **kwargs):
        
        label = Label(self, text=label_text, font=font)
        label.grid(row=self.row, column=column, **kwargs)
        self.row += not same_row

        return label

    def add_entry(self, width=40, column=0, padx=5, same_row=False,
                  num_only=False, default=None, show=None, **kwargs):

        vcmd = None
        if num_only:
            vcmd = (self.root.register(self.validate),
                            '%d', '%i', '%P', '%s', '%S', '%v', '%V', '%W')
        entry = Entry(self, width=width, validate = 'key', show=show,
            validatecommand=vcmd)
        if default is not None:
            entry.insert(END, str(default))

        entry.grid(row=self.row, column=column, padx=padx, **kwargs)
        self.row += not same_row

        return entry

    def add_button(self, btn_text, command, column=1,
                   same_row=False, **kwargs):
        button = Button(self, text=btn_text, command=command)
        button.grid(row=self.row-1, column=column, sticky="nesw", **kwargs)
        self.row += not same_row

        return button
    
    def add_text(self, height=5, width=52, column=0, columnspan=2,
                 padx=5, pady=5, same_row=False, **kwargs):
        text = Text(self, height=height, width=width)
        text.grid(row=self.row, column=column, columnspan=columnspan,
                  padx=padx, pady=pady)
        self.row += not same_row

        return text
	
    def validate(self, action, index, value_if_allowed,
                       prior_value, text, validation_type, trigger_type, widget_name):
        if value_if_allowed:
            try:
                int(value_if_allowed)
                return True
            except ValueError:
                return False
        else:
            return True

    def setExcelFile(self, excelFileInput):
        excelFilePath = askopenfilename(filetypes=[("Excel files", ".xlsx .xls")])
        excelFileInput.delete(0,END)
        excelFileInput.insert(0,excelFilePath)

    def start_sending(self, input_fields):
        def getValue(widget):
            if isinstance(widget, Text):
                text = widget.get("1.0", END)
                text = text.strip().replace("\n", "\\n").replace("\t", "\\t")
                return text
            else:
                val = widget.get()
                if val.isdigit():
                    return int(val)
                elif val.isdecimal():
                    return float(val)
                return val
        self.status_label.config(text="Status: Sending Messages")
        input_values = {field_name: getValue(widget)
            for field_name, widget in input_fields.items()}
        print("input_values", input_values)
        
        if not input_values["url_col"]:
            input_values["url_col"] = 2
        
        if not input_values["start_num"]:
            input_values["start_num"] = 1
        
        if not input_values["num_msg"]:
            input_values["num_msg"] = 100

        if not input_values["filepath"] or not path.exists(input_values["filepath"]):
            self.status_label.config(text="Status: Invalid filepath")
            print("Invalid FilePath")
            return

        if not input_values["msg_template"]:
            self.status_label.config(text="Status: Empty Message")
            print("Empty Message")
            return

        msg_sender = MsgSender(**input_values)
        error = msg_sender.start_sending()
        if error == 0:
            self.status_label.config(text="Status: Sent Messages")
            print("DONE")
        elif error == 1:
            self.status_label.config(text="Status: FAILED BROWSER CLOSED")
            print("Browser Closed")
        else:
            self.status_label.config(text="Status: FAILED UNKNOWN ERROR")
            print("FAILED")

