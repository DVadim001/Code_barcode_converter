import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import threading
import pandas as pd
import treepoem
import os
from PIL import Image, ImageDraw, ImageFont
from fpdf import FPDF
from concurrent.futures import ThreadPoolExecutor, as_completed


class BarcodeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("SSCC (GS1-128) PDF Generator — Оптимизированная версия")
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
            df = pd.read_excel(self.file_path, dtype=str, keep_default_na=False, header=None, engine='openpyxl')

            codes = []
            error_lines = []
            for idx, row in df.iterrows():
                if len(row) == 0:
                    continue

                code = str(row[0]).strip()
                code = ''.join(code.split())

                if not code:
                    error_lines.append(idx + 1)
                    continue

                if code.startswith("(00)"):
                    numeric_code = code[4:]
                else:
                    numeric_code = code

                if not numeric_code.isdigit() or len(numeric_code) != 18:
                    error_lines.append(idx + 1)
                    continue

                codes.append((idx, f"(00){numeric_code}"))

            if error_lines:
                raise ValueError(f"Ошибки в строках: {', '.join(map(str, error_lines))}")

            output_folder = os.path.dirname(self.file_path)
            output_name = os.path.splitext(os.path.basename(self.file_path))[0] + ".pdf"
            output_path = os.path.join(output_folder, output_name)

            # Увеличили размер шрифта и изменили на Arial
            try:
                font = ImageFont.truetype("arial.ttf", size=14)
            except:
                try:
                    font = ImageFont.truetype("LiberationSans-Regular.ttf", size=14)
                except:
                    font = ImageFont.load_default(size=14)

            self.update_status("Генерация штрих-кодов...")
            images = [None] * len(codes)

            def generate(index, full_code):
                barcode = treepoem.generate_barcode(barcode_type="gs1-128", data=full_code)
                img = barcode.convert("RGB")
                width, height = img.size

                # Увеличенные отступы для текста
                text_bbox = font.getbbox(full_code)
                text_height = text_bbox[3] - text_bbox[1]

                # Новые параметры отступов (увеличены в 2 раза)
                padding_top = 10  # Отступ сверху для штрих-кода
                padding_bottom = 15  # Отступ снизу для текста

                new_height = height + text_height + padding_top + padding_bottom
                new_img = Image.new("RGB", (width, new_height), "white")
                new_img.paste(img, (0, padding_top))  # Сдвигаем штрих-код вниз

                draw = ImageDraw.Draw(new_img)
                text_width = font.getlength(full_code)

                # Позиция текста (увеличена Y-координата)
                text_y = height + padding_top + 5
                draw.text(
                    ((width - text_width) // 2, text_y),
                    full_code,
                    fill="black",
                    font=font
                )
                return index, new_img

            with ThreadPoolExecutor(max_workers=6) as executor:
                futures = [executor.submit(generate, idx, full_code) for idx, full_code in codes]
                for count, future in enumerate(as_completed(futures), 1):
                    idx, image = future.result()
                    images[idx] = image
                    self.update_status(f"Генерировано: {count} из {len(codes)}")

            self.update_status("Создание PDF...")

            # Увеличенный размер страницы PDF
            pdf = FPDF(unit="mm", format=(100, 60))  # Было 40, увеличено до 60 мм
            pdf.set_auto_page_break(False)

            for i, img in enumerate(images):
                pdf.add_page()
                img_path = os.path.join(output_folder, f"temp_{i}.png")
                img.save(img_path, "PNG", quality=100)

                # Оптимальные размеры для размещения на увеличенной странице
                pdf_width = 100
                img_w_mm = 80
                img_h_mm = 40  # Фиксированная высота области для штрих-кода

                # Центрирование на странице с увеличенными отступами
                x = (pdf_width - img_w_mm) / 2
                y = 10  # Отступ сверху увеличен

                pdf.image(img_path, x=x, y=y, w=img_w_mm, h=img_h_mm)
                os.remove(img_path)

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