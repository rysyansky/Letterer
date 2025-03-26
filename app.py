import python_docs
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkcalendar import DateEntry
from docxtpl import DocxTemplate
import os

root = tk.Tk()
root.title("Letterer")
root.geometry("600x500")

main_frame = tk.Frame(root)
main_frame.pack(fill="both", expand=True, padx=5, pady=5)

canvas = tk.Canvas(main_frame)
scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
scrollable_frame = tk.Frame(canvas)

scrollable_frame.bind(
    "<Configure>",
    lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
)
canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
canvas.configure(yscrollcommand=scrollbar.set)

canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")



def add_label_entry(row, label_text, widget_type="entry", height=1):
    label = tk.Label(scrollable_frame, text=label_text)
    label.grid(row=row, column=0, sticky="w", padx=5, pady=2)

    if widget_type == "entry":
        entry = tk.Entry(scrollable_frame)
        entry.grid(row=row, column=1, sticky="ew", padx=5, pady=2)
        return entry
    elif widget_type == "text":
        text = tk.Text(scrollable_frame, height=height)
        text.grid(row=row, column=1, sticky="nsew", padx=5, pady=2)
        return text
    elif widget_type == "date":
        date = DateEntry(scrollable_frame, date_pattern='dd.mm.yyyy')
        date.grid(row=row, column=1, sticky="ew", padx=5, pady=2)
        return date


scrollable_frame.grid_columnconfigure(1, weight=1)

org_e = add_label_entry(0, "Название организации:")
from_t = add_label_entry(1, "От кого:", "text", 3)
to_t = add_label_entry(2, "Кому:", "text", 3)
date_d = add_label_entry(3, "Дата:", "date")
purp_e = add_label_entry(4, "Назначение письма:")
l_text_t = add_label_entry(5, "Текст письма:", "text", 6)
pos_e = add_label_entry(6, "Должность отправителя:")
name_e = add_label_entry(7, "Имя отправителя:")


def save_to_docx():
    context = get_form_data()

    template_path = os.path.join(os.path.dirname(__file__), "Template/letter_template.docx")
    doc = DocxTemplate(template_path)

    # Запрашиваем место сохранения
    file_path = filedialog.asksaveasfilename(
        defaultextension=".docx",
        filetypes=[("Word Documents", "*.docx")],
        title="Сохранить документ"
    )

    if file_path:
        # Вставляем данные в шаблон
        doc.render(context)
        doc.save(file_path)
        messagebox.showinfo("Успех", "Документ успешно сохранён!")

def get_form_data():
    return {
        "org_name": org_e.get(),
        "from_text": from_t.get("1.0", tk.END).strip(),
        "to_text": to_t.get("1.0", tk.END).strip(),
        "date": date_d.get(),
        "purpose": purp_e.get(),
        "letter_text": l_text_t.get("1.0", tk.END).strip(),
        "sender_position": pos_e.get(),
        "sender_name": name_e.get()
    }

btn = tk.Button(scrollable_frame, text="Сохранить документ", command=save_to_docx)
btn.grid(row=8, column=0, columnspan=2, pady=10)


root.mainloop()