import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import threading
import pandas as pd
import treepoem
import os
from PIL import Image, ImageDraw, ImageFont
from fpdf import FPDF


class BarcodeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("SSCC (GS1-128) PDF Generator")
        self.root.geometry("500x200")

        self.file_path = None

        menubar = tk.Menu(root)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Открыть (Ctrl+O)", command=self.select_file)
        menubar.add_cascade(label="Файл", menu=file_menu)
        root.config(menu=menubar)
        root.bind('<Control-o>', lambda event: self.select_file())

        self.select_button = tk.Button(root, text="Выбрать Excel файл", command=self.select_file)
        self.select_button.pack(pady=10)

        self.process_button = tk.Button(root, text="Старт обработки", command=self.start_thread, state=tk.DISABLED)
        self.process_button.pack(pady=10)

        self.status_var = tk.StringVar()
        self.status_var.set("Ожидание файла...")
        self.status_bar = ttk.Label(root, textvariable=self.status_var, relief=tk.SUNKEN, anchor='w')
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def select_file(self):
        filetypes = [("Excel файлы", "*.xlsx")]
        path = filedialog.askopenfilename(filetypes=filetypes)
        if path:
            self.file_path = path
            self.status_var.set(f"Выбран файл: {os.path.basename(path)}")
            self.process_button.config(state=tk.NORMAL)

    def start_thread(self):
        self.process_button.config(state=tk.DISABLED)
        thread = threading.Thread(target=self.process_file)
        thread.start()

    def process_file(self):
        try:
            self.update_status("Чтение Excel...")
            df = pd.read_excel(self.file_path)
            codes = df.iloc[:, 0].dropna().astype(str).tolist()

            output_folder = os.path.dirname(self.file_path)
            output_name = os.path.splitext(os.path.basename(self.file_path))[0] + ".pdf"
            output_path = os.path.join(output_folder, output_name)

            images = []
            font = ImageFont.load_default()

            for idx, code in enumerate(codes):
                self.update_status(f"Генерация штрих-кодов ({idx + 1} из {len(codes)})...")

                code = code.strip()

                # Удалим лишние символы (если код содержит (00))
                if code.startswith("(00)"):
                    numeric_code = code[4:]
                else:
                    numeric_code = code

                if not numeric_code.isdigit() or len(numeric_code) != 18:
                    raise ValueError(f"Некорректный SSCC-код в строке {idx + 1}: {code} (должно быть 18 цифр после (00))")

                full_code = f"(00){numeric_code}"
                barcode = treepoem.generate_barcode(barcode_type="gs1-128", data=full_code)
                img = barcode.convert("RGB")

                # Подпись под штрихкодом
                width, height = img.size
                new_img = Image.new("RGB", (width, height + 30), "white")
                new_img.paste(img, (0, 0))

                draw = ImageDraw.Draw(new_img)
                draw.text((10, height + 5), full_code, fill="black", font=font)

                images.append(new_img)

            self.update_status("Создание PDF...")

            pdf = FPDF(unit="mm", format="A4")
            x_margin = 30
            y_margin = 20
            spacing = 70
            items_per_page = 10

            for i, img in enumerate(images):
                if i % items_per_page == 0:
                    pdf.add_page()
                    y_offset = y_margin

                img_path = os.path.join(output_folder, f"temp_{i}.png")
                img.save(img_path)
                pdf.image(img_path, x=x_margin, y=y_offset, w=150)
                os.remove(img_path)
                y_offset += spacing

            pdf.output(output_path)
            self.update_status(f"Готово! Сохранено: {output_path}")
            messagebox.showinfo("Успех", f"PDF сохранён:\n{output_path}")

        except Exception as e:
            self.update_status("Ошибка")
            messagebox.showerror("Ошибка", str(e))
        finally:
            self.process_button.config(state=tk.NORMAL)

    def update_status(self, text):
        self.status_var.set(text)
        self.status_bar.update_idletasks()


if __name__ == "__main__":
    root = tk.Tk()
    app = BarcodeApp(root)
    root.mainloop()
