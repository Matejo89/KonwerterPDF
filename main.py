import tkinter as tk
from tkinter import filedialog, messagebox, Toplevel
from tkinter.ttk import Progressbar
from pdf2docx import Converter
from PIL import Image, ImageTk
import os

# Funkcja do aktualizacji paska stanu
def zmien_status(text):
    pasek_stanu.config(text=text)

# Funkcja do wyboru pliku PDF
def select_pdf():
    file_path = filedialog.askopenfilename(
        filetypes=[("PDF files", "*.pdf")])
    if file_path:
        pdf_label.config(text=f"Wybrano plik: {file_path}")
        export_button.config(state=tk.NORMAL)
        zmien_status("Wybrano plik PDF")
        return file_path
    zmien_status("Nie wybrano pliku PDF")
    return None

# Funkcja do tworzenia okna postępu
def create_progress_window(task_description):
    progress_window = Toplevel(root)
    progress_window.title("Progres")
    progress_label = tk.Label(progress_window, text=task_description)
    progress_label.pack(pady=10)
    
    progress_bar = Progressbar(progress_window, length=300, mode='indeterminate')
    progress_bar.pack(pady=10)
    progress_bar.start()

    zmien_status(task_description)
    
    return progress_window, progress_bar

# Funkcja do zamknięcia okna postępu
def close_progress_window(progress_window, progress_bar):
    progress_bar.stop()
    progress_window.destroy()
    zmien_status("Zakończono konwersję PDF do Word")

# Funkcja do konwersji PDF do Word za pomocą pdf2docx
def export_word(pdf_path):
    try:
        output_docx_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word files", "*.docx")])
        if output_docx_path:
            # Tworzymy okno postępu
            progress_window, progress_bar = create_progress_window("Konwertowanie PDF do Worda...")

            # Konwersja PDF do Worda
            cv = Converter(pdf_path)
            cv.convert(output_docx_path, start=0, end=None)
            cv.close()

            # Zamykamy okno postępu po zakończeniu
            close_progress_window(progress_window, progress_bar)
            messagebox.showinfo("Sukces", "Zapisano plik Word z zachowanym formatowaniem!")
        zmien_status("Zapisano plik Word")
    except Exception as e:
        messagebox.showerror("Błąd", f"Wystąpił błąd: {e}")
        zmien_status("Wystąpił błąd podczas konwersji")

# Funkcja do wyboru formatu eksportu
def export_pdf():
    pdf_path = pdf_label.cget("text").replace("Wybrano plik: ", "")
    if not pdf_path:
        messagebox.showerror("Błąd", "Najpierw wybierz plik PDF!")
        zmien_status("Błąd: brak wybranego pliku PDF")
        return
    
    export_word(pdf_path)

# Tworzenie GUI
root = tk.Tk()
root.title("PDF Konwerter")
script_dir = os.path.dirname(os.path.abspath(__file__))
icon_path = os.path.join(script_dir, "favicon.ico")
root.iconbitmap(icon_path)
root.geometry("400x300")
root.resizable ( 0 , 0 )

# Dodawanie loga do okna aplikacji
def add_logo():
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(script_dir, "Naglak.png")
        logo_img = Image.open(icon_path)  # Podaj ścieżkę do pliku z logiem
        logo_img = logo_img.resize((250, 100))  # Zmiana rozmiaru loga
        logo_photo = ImageTk.PhotoImage(logo_img)
        logo_label = tk.Label(root, image=logo_photo)
        logo_label.image = logo_photo  # Referencja do obrazu (aby go nie usunęło)
        logo_label.pack(pady=10)
    except Exception as e:
        messagebox.showerror("Błąd", f"Nie udało się wczytać loga: {e}")

# Wywołanie funkcji dodającej logo
add_logo()

# Etykieta dla wybranego pliku PDF
pdf_label = tk.Label(root, text="Nie wybrano pliku PDF")
pdf_label.pack(pady=10)

# Przycisk do wyboru pliku PDF
select_button = tk.Button(root, text="Wybierz PDF", command=select_pdf)
select_button.pack(pady=5)

# Przycisk do eksportu
export_button = tk.Button(root, text="Eksportuj do Worda", command=export_pdf, state=tk.DISABLED)
export_button.pack(pady=10)

# Dodanie paska stanu na dole okna
pasek_stanu = tk.Label(root, text="Gotowy", bd=1, relief=tk.SUNKEN, anchor=tk.W)
pasek_stanu.pack(side=tk.BOTTOM, fill=tk.X)

root.mainloop()
