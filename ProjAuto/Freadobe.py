import os
import fitz 
import re
from pdf2image import convert_from_path
from PIL import Image 
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
from tkinter import (
     Toplevel, Label, Button, Listbox, messagebox,
    simpledialog, filedialog, SINGLE, END
)
from tkinterdnd2 import TkinterDnD, DND_FILES


class StartMenu:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Utility")
        self.root.geometry("300x300")
        self.root.resizable(False, False)

        Label(root, text="Choose a function:", font=("Arial", 14)).pack(pady=20)

        Button(root, text="🖼 Image to PDF", width=20, command=self.open_image_to_pdf).pack(pady=10)
        Button(root, text="📄 Merge PDFs", width=20, command=self.open_merge_pdf).pack(pady=10)
        Button(root, text="✂️ Split PDF", width=20, command=self.open_split_pdf).pack(pady=10) 
        Button(root, text="🔄 Rearrange PDF Pages", width=20, command=self.open_rearrange_pdf).pack(pady=10)

    def open_image_to_pdf(self):
        ImageToPDFApp(Toplevel(self.root))

    def open_merge_pdf(self):
        MergePDFApp(Toplevel(self.root))

    def open_split_pdf(self):
        SplitPDFApp(Toplevel(self.root)) 

    def open_rearrange_pdf(self):
        RearrangePDFApp(Toplevel(self.root)) 


class ImageToPDFApp:
    def __init__(self, window):
        self.window = window
        self.window.title("Multi-Image to PDF Converter")
        self.window.geometry("500x400")
        self.window.resizable(False, False)

        self.image_files = []

        self.drop_label = Label(self.window, text="📂 Drag and drop images here", relief="groove",
                                borderwidth=2, width=50, height=4, font=("Arial", 12))
        self.drop_label.pack(pady=10)
        self.drop_label.drop_target_register(DND_FILES)
        self.drop_label.dnd_bind('<<Drop>>', self.on_drop)

        self.image_listbox = Listbox(self.window, width=60, height=10, selectmode=SINGLE)
        self.image_listbox.pack(pady=10)

        Button(self.window, text="Next ➜", font=("Arial", 12), command=self.convert_images_to_pdf,
               bg="#4CAF50", fg="white", padx=10, pady=5).pack(pady=10)

        Button(self.window, text="Exit", command=self.window.destroy).pack()

        control_frame = Label(self.window)
        control_frame.pack(pady=5)

        Button(control_frame, text="🗑 Remove Selected", command=self.remove_selected_image).grid(row=0, column=0, padx=5)
        Button(control_frame, text="🧹 Clear All", command=self.clear_all_images).grid(row=0, column=1, padx=5)
        Button(control_frame, text="🔼 Move Up", command=self.move_selected_up).grid(row=0, column=2, padx=5)
        Button(control_frame, text="🔽 Move Down", command=self.move_selected_down).grid(row=0, column=3, padx=5)

    def add_images(self, files):
        for path in files:
            path = path.strip('{}')
            if os.path.isfile(path):
                ext = os.path.splitext(path)[1].lower()
                if ext in [".png", ".jpg", ".jpeg", ".bmp", ".tiff"]:
                    if path not in self.image_files:
                        self.image_files.append(path)
                        self.image_listbox.insert(END, os.path.basename(path))

    def on_drop(self, event):
        raw = event.data
        files = re.findall(r'{.*?}|\S+', raw)
        files = [f.strip('{}') for f in files]
        self.add_images(files)

    def convert_images_to_pdf(self):
        if not self.image_files:
            messagebox.showwarning("No Images", "Please add at least one image.")
            return

        file_name = simpledialog.askstring("PDF Name", "Enter a name for the PDF file (without extension):")
        if not file_name:
            return

        folder = filedialog.askdirectory(title="Select folder to save the PDF")
        if not folder:
            return

        pdf_path = os.path.join(folder, f"{file_name}.pdf")

        try:
            images = []
            for img_path in self.image_files:
                img = Image.open(img_path)
                if img.mode != "RGB":
                    img = img.convert("RGB")
                images.append(img)

            first, *rest = images
            first.save(pdf_path, save_all=True, append_images=rest)

            self.image_files.clear()
            self.image_listbox.delete(0, END)

            messagebox.showinfo("Success", f"PDF saved to:\n{pdf_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create PDF:\n{str(e)}")

    def remove_selected_image(self):
        selected_index = self.image_listbox.curselection()
        if not selected_index:
            return
        index = selected_index[0]
        del self.image_files[index]
        self.image_listbox.delete(index)

    def clear_all_images(self):
        self.image_files.clear()
        self.image_listbox.delete(0, END)

    def move_selected_up(self):
        selected = self.image_listbox.curselection()
        if not selected or selected[0] == 0:
            return
        i = selected[0]
        self.image_files[i-1], self.image_files[i] = self.image_files[i], self.image_files[i-1]
        self.update_listbox_selection(i - 1)

    def move_selected_down(self):
        selected = self.image_listbox.curselection()
        if not selected or selected[0] == len(self.image_files) - 1:
            return
        i = selected[0]
        self.image_files[i], self.image_files[i+1] = self.image_files[i+1], self.image_files[i]
        self.update_listbox_selection(i + 1)

    def update_listbox_selection(self, new_index):
        self.image_listbox.delete(0, END)
        for img_path in self.image_files:
            self.image_listbox.insert(END, os.path.basename(img_path))
        self.image_listbox.selection_set(new_index)


class MergePDFApp:
    def __init__(self, master):
        self.master = master
        self.pdf_files = []  # List of PDF file paths

        self.window = Toplevel(master)
        self.window.title("Merge PDFs")
        self.window.geometry("500x400")
        self.window.resizable(False, False)

        self.drop_label = Label(self.window, text="📂 Drag and drop PDF files here", relief="groove",
                                borderwidth=2, width=50, height=4, font=("Arial", 12))
        self.drop_label.pack(pady=10)
        self.drop_label.drop_target_register(DND_FILES)
        self.drop_label.dnd_bind('<<Drop>>', self.on_drop)

        self.file_listbox = Listbox(self.window, width=60, height=10, selectmode=SINGLE)
        self.file_listbox.pack(pady=10)

        control_frame = Label(self.window)
        control_frame.pack(pady=5)

        Button(control_frame, text="🗑 Remove", command=self.remove_selected).grid(row=0, column=0, padx=5)
        Button(control_frame, text="🔼 Move Up", command=self.move_up).grid(row=0, column=1, padx=5)
        Button(control_frame, text="🔽 Move Down", command=self.move_down).grid(row=0, column=2, padx=5)
        Button(control_frame, text="🧹 Clear All", command=self.clear_all).grid(row=0, column=3, padx=5)

        Button(self.window, text="Merge ➜", font=("Arial", 12), command=self.merge_files,
               bg="#4CAF50", fg="white", padx=10, pady=5).pack(pady=10)

        Button(self.window, text="Back", command=self.window.destroy).pack()

    def on_drop(self, event):
        raw = event.data
        files = re.findall(r'{.*?}|[^\s]+', raw)
        for f in files:
            f = f.strip('{}')
            ext = os.path.splitext(f)[1].lower()
            if ext == '.pdf' and f not in self.pdf_files:
                self.pdf_files.append(f)
                self.file_listbox.insert(END, os.path.basename(f))

    def remove_selected(self):
        i = self.file_listbox.curselection()
        if not i:
            return
        index = i[0]
        del self.pdf_files[index]
        self.file_listbox.delete(index)

    def move_up(self):
        i = self.file_listbox.curselection()
        if not i or i[0] == 0:
            return
        idx = i[0]
        self.pdf_files[idx - 1], self.pdf_files[idx] = self.pdf_files[idx], self.pdf_files[idx - 1]
        self.refresh_listbox(idx - 1)

    def move_down(self):
        i = self.file_listbox.curselection()
        if not i or i[0] == len(self.pdf_files) - 1:
            return
        idx = i[0]
        self.pdf_files[idx], self.pdf_files[idx + 1] = self.pdf_files[idx + 1], self.pdf_files[idx]
        self.refresh_listbox(idx + 1)

    def clear_all(self):
        self.pdf_files.clear()
        self.file_listbox.delete(0, END)

    def refresh_listbox(self, selected_index):
        self.file_listbox.delete(0, END)
        for path in self.pdf_files:
            self.file_listbox.insert(END, os.path.basename(path))
        self.file_listbox.selection_set(selected_index)

    def merge_files(self):
        if not self.pdf_files:
            messagebox.showwarning("Empty List", "Please add some PDF files first.")
            return

        name = simpledialog.askstring("PDF Name", "Enter name for merged PDF (without .pdf):")
        if not name:
            return

        folder = filedialog.askdirectory(title="Select where to save the PDF")
        if not folder:
            return

        out_path = os.path.join(folder, f"{name}.pdf")
        merger = PdfMerger()

        try:
            for pdf_path in self.pdf_files:
                merger.append(pdf_path)

            merger.write(out_path)
            merger.close()

            self.clear_all()
            messagebox.showinfo("Success", f"Merged PDF saved to:\n{out_path}")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to merge files:\n{str(e)}")


class SplitPDFApp:
    def __init__(self, master):
        self.master = master
        self.file_path = None

        self.window = Toplevel(master)
        self.window.title("Split PDF")
        self.window.geometry("500x300")
        self.window.resizable(False, False)

        self.drop_label = Label(self.window, text="📂 Drag and drop a PDF file here", relief="groove",
                                borderwidth=2, width=50, height=4, font=("Arial", 12))
        self.drop_label.pack(pady=10)
        self.drop_label.drop_target_register(DND_FILES)
        self.drop_label.dnd_bind('<<Drop>>', self.on_drop)

        self.file_listbox = Listbox(self.window, width=60, height=5, selectmode=SINGLE)
        self.file_listbox.pack(pady=10)

        Button(self.window, text="Split ➜", font=("Arial", 12),
               command=self.split_pdf, bg="#4CAF50", fg="white", padx=10, pady=5).pack(pady=10)

        Button(self.window, text="Back", command=self.window.destroy).pack()

    def on_drop(self, event):
        raw = event.data
        files = re.findall(r'{.*?}|[^\s]+', raw)
        for f in files:
            f = f.strip('{}')
            ext = os.path.splitext(f)[1].lower()
            if ext == '.pdf':
                self.file_path = f
                self.file_listbox.delete(0, END)
                self.file_listbox.insert(END, os.path.basename(f))
                break  # Accept only one PDF

    def split_pdf(self):
        if not self.file_path:
            messagebox.showwarning("No File", "Please drop a PDF file first.")
            return

        choice = messagebox.askquestion(
            "Split Format",
            "Do you want to split the PDF into image files?\n\n"
            "Click 'Yes' for images, 'No' for individual PDF pages."
        )
        save_as_image = choice == "yes"
        

        folder = filedialog.askdirectory(title="Select folder to save output")
        if not folder:
            return

        try:
            if save_as_image:
                self.split_to_images(folder)
            else:
                self.split_to_pdfs(folder)

            messagebox.showinfo("Success", f"Pages successfully split and saved to:\n{folder}")
            self.file_path = None
            self.file_listbox.delete(0, END)

        except Exception as e:
            messagebox.showerror("Error", f"Failed to split PDF:\n{str(e)}")

    def split_to_pdfs(self, folder):
        reader = PdfReader(self.file_path)
        for i, page in enumerate(reader.pages, start=1):
            writer = PdfWriter()
            writer.add_page(page)
            out_path = os.path.join(folder, f"{i}.pdf")
            with open(out_path, "wb") as f:
                writer.write(f)

    def split_to_images(self, folder):
        doc = fitz.open(self.file_path)
        for i, page in enumerate(doc, start=1):
            pix = page.get_pixmap(dpi=200)
            img_path = os.path.join(folder, f"{i}.jpg")
            pix.save(img_path)




class RearrangePDFApp:
    def __init__(self, master):
        self.master = master
        self.pages = []
        self.original_file_path = None

        self.window = Toplevel(master)
        self.window.title("Rearrange PDF Pages")
        self.window.geometry("500x450")
        self.window.resizable(False, False)

        self.drop_label = Label(self.window, text="📂 Drag and drop a PDF file here", relief="groove",
                                borderwidth=2, width=50, height=4, font=("Arial", 12))
        self.drop_label.pack(pady=10)
        self.drop_label.drop_target_register(DND_FILES)
        self.drop_label.dnd_bind('<<Drop>>', self.on_drop)

        self.page_listbox = Listbox(self.window, width=60, height=12, selectmode=SINGLE)
        self.page_listbox.pack(pady=10)
        

        control_frame = Label(self.window)
        control_frame.pack(pady=5)

        Button(control_frame, text="🔼 Move Up", command=self.move_up).grid(row=0, column=0, padx=5)
        Button(control_frame, text="🔽 Move Down", command=self.move_down).grid(row=0, column=1, padx=5)
        Button(control_frame, text="🗑 Delete", command=self.delete_selected_page).grid(row=0, column=2, padx=5)
        Button(control_frame, text="📤 Extract Page", command=self.extract_selected_page).grid(row=0, column=3, padx=5)
        self.insert_label = Button(control_frame, text="📥 Insert Page", command=self.insert_page)
        self.insert_label.grid(row=0, column=4, padx=5)
        self.insert_label.drop_target_register(DND_FILES)
        self.insert_label.dnd_bind('<<Drop>>', self.on_insert_drop)



        Button(self.window, text="💾 Save", font=("Arial", 12),
               command=self.save_pdf, bg="#4CAF50", fg="white", padx=10, pady=5).pack(pady=10)

        Button(self.window, text="Back", command=self.window.destroy).pack()

    def on_drop(self, event):
        raw = event.data
        files = re.findall(r'{.*?}|[^\s]+', raw)
        for f in files:
            f = f.strip('{}')
            ext = os.path.splitext(f)[1].lower()
            if ext == '.pdf':
                self.original_file_path = f
                self.load_pdf_pages(f)
                break

    def load_pdf_pages(self, file_path):
        self.reader = PdfReader(file_path)
        self.pages = list(self.reader.pages)
        self.refresh_listbox()

    def refresh_listbox(self, selected_index=0):
        self.page_listbox.delete(0, END)
        for i in range(len(self.pages)):
            self.page_listbox.insert(END, f"Page {i + 1}")
        if 0 <= selected_index < len(self.pages):
            self.page_listbox.selection_set(selected_index)

    def move_up(self):
        i = self.page_listbox.curselection()
        if not i or i[0] == 0:
            return
        idx = i[0]
        self.pages[idx - 1], self.pages[idx] = self.pages[idx], self.pages[idx - 1]
        self.refresh_listbox(idx - 1)

    def move_down(self):
        i = self.page_listbox.curselection()
        if not i or i[0] == len(self.pages) - 1:
            return
        idx = i[0]
        self.pages[idx], self.pages[idx + 1] = self.pages[idx + 1], self.pages[idx]
        self.refresh_listbox(idx + 1)

    def delete_selected_page(self):
        i = self.page_listbox.curselection()
        if not i:
            return
        idx = i[0]
        del self.pages[idx]
        self.refresh_listbox(max(0, idx - 1))

    def extract_selected_page(self):
        i = self.page_listbox.curselection()
        if not i:
            messagebox.showwarning("No Selection", "Please select a page to extract.")
            return
        idx = i[0]
        if idx >= len(self.pages):
            messagebox.showerror("Invalid Index", "Selected page index is out of range.")
            return

        page = self.pages[idx]

        folder = filedialog.askdirectory(title="Select folder to save extracted page")
        if not folder:
            return

        filename = simpledialog.askstring("Save As", "Enter name for extracted PDF (without .pdf):")
        if not filename:
            return

        out_path = os.path.join(folder, f"{filename}.pdf")

        try:
            writer = PdfWriter()
            writer.add_page(page)
            with open(out_path, "wb") as f:
                writer.write(f)
            messagebox.showinfo("Success", f"Page extracted and saved to:\n{out_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to extract page:\n{str(e)}")



    def on_insert_drop(self, event):
        if not self.original_file_path:
            messagebox.showwarning("No PDF Loaded", "Please load a PDF file first.")
            return

        selected = self.page_listbox.curselection()
        insert_index = selected[0] + 1 if selected else len(self.pages)

        raw = event.data
        files = re.findall(r'{.*?}|[^\s]+', raw)
        inserted_pages = []

        for file in files:
            file = file.strip('{}')
            ext = os.path.splitext(file)[1].lower()

            if ext == '.pdf':
                try:
                    reader = PdfReader(file)
                    for page in reader.pages:
                        inserted_pages.append(page)
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to read PDF:\n{str(e)}")

            elif ext in ['.png', '.jpg', '.jpeg', '.bmp', '.tiff']:
                try:
                    img = Image.open(file)
                    if img.mode != 'RGB':
                        img = img.convert('RGB')

                    temp_path = os.path.join(os.getcwd(), "_temp_insert.pdf")
                    img.save(temp_path, "PDF")
                    reader = PdfReader(temp_path)
                    inserted_pages.append(reader.pages[0])
                    os.remove(temp_path)

                except Exception as e:
                    messagebox.showerror("Error", f"Failed to insert image:\n{str(e)}")
            else:
                messagebox.showwarning("Unsupported Format", f"Unsupported file: {file}")

        if inserted_pages:
            for i, page in enumerate(inserted_pages):
                self.pages.insert(insert_index + i, page)

            self.refresh_listbox(selected_index=insert_index + len(inserted_pages) - 1)








    def save_pdf(self):
        if not self.pages:
            messagebox.showwarning("Empty PDF", "There are no pages to save.")
            return

        folder = filedialog.askdirectory(title="Select folder to save rearranged PDF")
        if not folder:
            return

        name = simpledialog.askstring("PDF Name", "Enter name for rearranged PDF (without .pdf):")
        if not name:
            return

        out_path = os.path.join(folder, f"{name}.pdf")
        try:
            writer = PdfWriter()
            for page in self.pages:
                writer.add_page(page)
            with open(out_path, "wb") as f:
                writer.write(f)
            messagebox.showinfo("Success", f"Rearranged PDF saved to:\n{out_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save PDF:\n{str(e)}")







# --- Entry Point ---
if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = StartMenu(root)
    root.mainloop()
