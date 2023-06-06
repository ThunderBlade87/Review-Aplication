import tkinter as tk
import os
from tkinter import messagebox
from docx import Document
from docx.enum.section import WD_SECTION
from datetime import datetime

if os.path.isfile("output.docx"):
    print("Fișierul există.")
else:
    document = Document()
    document.save("output.docx")

def export_to_word_title(ticket_number, assignee, reviewer, dt_string):
    input_ticket_number= ticket_number
    input_assignee = assignee
    input_reviewer = reviewer
    input_dt_string = dt_string
    document = Document("output.docx")
    total_text_length = sum(len(p.text) for p in document.paragraphs)
    if total_text_length > 0:
        document.add_section(WD_SECTION.NEW_PAGE)
    document.add_paragraph(f"Data: {input_dt_string}")
    document.add_paragraph(f"Assignee: {input_assignee}")
    document.add_paragraph(f"Reviewer: {input_reviewer}")
    document.add_paragraph(f"Ticket number: {input_ticket_number}")
    document.save("output.docx")

def export_to_word():
    input_text1 = text1.get("1.0", "end-1c")
    input_text2 = text2.get("1.0", "end-1c")

    document = Document("output.docx")
    #document.add_section(WD_SECTION.NEW_PAGE)
    document.add_paragraph(f"Text1: {input_text1}")
    document.add_paragraph(f"Text2: {input_text2}")
    document.save("output.docx")

    messagebox.showinfo("Pop-up", "Text a fost exportat în fișierul output.docx")

def open_popup():
    def on_popup_close():
        ticket_number = entry1.get()
        assignee = entry2.get()
        reviewer = entry3.get()
        now = datetime.now()
        dt_string = now.strftime("%d %B %Y %H:%M:%S")
        export_to_word_title(ticket_number, assignee, reviewer, dt_string)
        popup.destroy()

    popup = tk.Toplevel()
    popup.geometry("400x200")
    popup.title("Pop-up")

    popup_label = tk.Label(popup, text="Number ticket:")
    popup_label.pack()

    entry1 = tk.Entry(popup)
    entry1.pack()

    popup_label = tk.Label(popup, text="Assignee:")
    popup_label.pack()

    entry2 = tk.Entry(popup)
    entry2.pack()

    popup_label = tk.Label(popup, text="Reviewer:")
    popup_label.pack()

    entry3 = tk.Entry(popup)
    entry3.pack()

    popup_button = tk.Button(popup, text="Export to Word", command=on_popup_close)
    popup_button.pack()

window = tk.Tk()
window.geometry("500x500")  # Setarea dimensiunilor ferestrei
#window.resizable(False, False)

label1 = tk.Label(window, text="Problem")
label1.pack(anchor="w")

text1 = tk.Text(window, width=0,height=0,borderwidth =7)
text1.pack(anchor="w",fill="both", expand=True)

label2 = tk.Label(window, text="Solution")
label2.pack(anchor="w")

text2 = tk.Text(window, width=0,height=0,borderwidth =7)
text2.pack(anchor="w",fill="both", expand=True)

button = tk.Button(window, text="Export to Word", command=export_to_word)
button.pack(anchor="c")

popup_button = tk.Button(window, text="Open Pop-up", command=open_popup)
popup_button.pack(anchor="c")

window.mainloop()