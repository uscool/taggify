import re
import win32com.client as win32
import tkinter as tk
from tkinter import ttk,messagebox,simpledialog
import os
import pandas as pd

tags={
    "work":[r"\bproject\b",r"\bmeeting\b"],
    "important":[r"\bimportant\b",r"\battention\b",r"follow[-\s]?up"],
    "urgent":[r"\burgent\b",r"\basap\b",r"\bimmediately\b"],
    "time-sensitive":[r"due\s?date",r"deadline",r"\btoday\b",r"\btomorrow\b",r"schedule"],
    "business event":[r"\bevent\b",r"\bconference\b",r"\bwebinar\b",r"\bMUN\b"],
    "classes":[r"\bclass\b",r"\bcourse\b",r"\blecture\b",r"\bPOA\b",r"\bPOE\b",r"\bFOE\b",r"\bDSA\b",r"\bIPP\b",r"\bIBM\b",r"\bBSAM\b"],
}

class Detagger:
    def __init__(self,filename="detagged_emails.csv"):
        self.filename=filename
        self.detagged_emails=self.load_detagged()

    def load_detagged(self):
        if os.path.exists(self.filename):
            df=pd.read_csv(self.filename)
            return set(zip(df["Subject"], df["Received"]))
        return set()

    def save_detagged_email(self,subject,received):
        new_entry=pd.DataFrame({"Subject":[subject],"Received":[received]})
        if os.path.exists(self.filename):
            new_entry.to_csv(self.filename,mode="a",header=False,index=False)
        else:
            new_entry.to_csv(self.filename,index=False)
        self.detagged_emails.add((subject,received))

detagged_manager=Detagger()

def fetch_and_tag_emails():
    outlook=win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox=outlook.Folders.Item("ujjwal.sharma04@nmims.in").Folders("Inbox")
    messages=inbox.Items
    messages.Sort("[ReceivedTime]",True)
    email_data=[]

    for message in messages:
        subject=message.Subject
        received_time=message.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S")
        body=message.Body
        tag_list=[]
        for tag,patterns in tags.items():
            if any(re.search(pattern,subject,re.IGNORECASE) or re.search(pattern,body,re.IGNORECASE) for pattern in patterns):
                tag_list.append(tag)
        
        if (subject,received_time) in detagged_manager.detagged_emails:
            tag_list=["No Tags"]
        elif not tag_list:
            tag_list=["No Tags"]

        priority="urgent" if "urgent" in tag_list else ("important" if "important" in tag_list else "normal")
        email_data.append((subject,received_time,", ".join(tag_list),priority,message.EntryID))

    return sorted(email_data,key=lambda x:(x[3]=="urgent",x[3]=="important",x[1]),reverse=True)

def display_emails(filter_text=""):
    email_tree.delete(*email_tree.get_children())
    for email in fetch_and_tag_emails():
        if filter_text.lower() in email[0].lower():
            tag_color="urgent" if email[3]=="urgent" else ("important" if email[3]=="important" else "normal")
            email_tree.insert("",tk.END,values=(email[0],email[1],email[2]),tags=(tag_color,email[4]))

def open_selected_email():
    selected=email_tree.focus()
    if selected:
        entry_id=email_tree.item(selected,"tags")[1]
        outlook=win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
        outlook.GetItemFromID(entry_id).Display()

def detag_mail():
    selected=email_tree.focus()
    if selected:
        subject,received=email_tree.item(selected,"values")[:2]
        if messagebox.askyesno("Confirm De-Tag","Remove tags from this email?"):
            detagged_manager.save_detagged_email(subject,received)
            email_tree.item(selected,values=(subject,received,"No Tags"))

def filter_emails():
    filter_text=simpledialog.askstring("Filter Emails","Enter keyword or tag to filter:")
    if filter_text:
        display_emails(filter_text)

root=tk.Tk()
root.title("Taggify")

email_tree=ttk.Treeview(root,columns=("Subject","Received","Tags"),show="headings")
email_tree.heading("Subject",text="Subject")
email_tree.heading("Received",text="Received Time")
email_tree.heading("Tags",text="Tags")
email_tree.tag_configure("urgent",background="indianred",foreground="white")
email_tree.tag_configure("important",background="#F0FFFF")
email_tree.tag_configure("normal",background="white")
email_tree.grid(row=1,column=0,columnspan=3,sticky="nsew")
ttk.Scrollbar(root,orient="vertical",command=email_tree.yview).grid(row=1,column=3,sticky="ns")

toolbar=tk.Frame(root)
toolbar.grid(row=0,column=0,columnspan=3,sticky="ew")
ttk.Button(toolbar,text="Fetch",command=display_emails).pack(side="left",padx=5,pady=5)
ttk.Button(toolbar,text="Open Email",command=open_selected_email).pack(side="left",padx=5,pady=5)
ttk.Button(toolbar,text="De-Tag",command=detag_mail).pack(side="left",padx=5,pady=5)
ttk.Button(toolbar,text="Filter",command=filter_emails).pack(side="left",padx=5,pady=5)

root.grid_rowconfigure(1,weight=1)
root.grid_columnconfigure(0,weight=1)

display_emails()
root.mainloop()
