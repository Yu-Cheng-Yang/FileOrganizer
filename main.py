import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import openai
import docx
import fitz
from openai import OpenAI
import customtkinter as ctk
from customtkinter import filedialog
from time import sleep

# Initialize OpenAI API client
api_key = os.environ.get("OPENAI_API_KEY")
if not api_key:
    raise ValueError("Please set the OPENAI_API_KEY environment variable.")
openai.api_key = api_key

categories = {
            'Images': ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.webp'],
            'Documents': ['.pdf', '.docx', '.xlsx', '.pptx', '.txt', '.md', '.csv'],
            'Archives': ['.zip', '.rar', '.tar', '.gz', '.7z'],
            'Audio': ['.mp3', '.wav', '.aac', '.flac'],
            'Videos': ['.mp4', '.mkv', '.mov', '.avi', '.wmv'],
            'Executables': ['.exe', '.msi'],
            "Other":['Anything else. Any edits towards this folder will be pointless']

        }

class FileOrganizerApp:
    def __init__(self, root):
        self.start_button = None
        self.root = root
        self.root.title('File Organizer')
        self.create_widgets()
        self.client = OpenAI(api_key=api_key)

    def create_widgets(self):
        ctk.set_appearance_mode("Light")
        self.folder_frame = ctk.CTkFrame(self.root)  # Changed to CTkFrame
        self.folder_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        # Configure the column and row weights to make them responsive
        self.root.columnconfigure(0, weight=1)
        self.folder_frame.columnconfigure(0, weight=1)

        self.downloads_folder_var = tk.StringVar()
        self.folder_entry = ctk.CTkEntry(self.folder_frame, textvariable=self.downloads_folder_var, width=50)
        self.folder_entry.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

        self.browse_button = ctk.CTkButton(self.folder_frame, text="Browse", command=self.browse_folder)
        self.browse_button.grid(row=0, column=1, padx=10, pady=10)

        # Start Button
        self.start_button = ctk.CTkButton(self.root, text="Start Organizing", command=self.start_organizing)
        self.start_button.grid(row=1, column=0, padx=10, pady=10, sticky="ew")

        # Metadata Display Frame
        self.metadata_frame = ttk.LabelFrame(self.root, text="Metadata Display")
        self.metadata_frame.grid(row=2, column=0, padx=10, pady=10, sticky="nsew")

        # Configure the column and row weights for the metadata frame
        self.metadata_frame.columnconfigure(0, weight=1)
        self.metadata_frame.rowconfigure(0, weight=1)

        # Replace Treeview with Text widget for metadata display
        self.metadata_text = ctk.CTkTextbox(self.metadata_frame, wrap="word")
        self.metadata_text.grid(row=0, column=0, sticky='nsew')

        # Add scrollbars for the Text widget
        self.scrollbar_vertical_text = ctk.CTkScrollbar(self.metadata_frame, command=self.metadata_text.yview)
        self.scrollbar_vertical_text.grid(row=0, column=1, sticky='ns')
        self.metadata_text.configure(yscrollcommand=self.scrollbar_vertical_text.set)
        # Add components for category management
        self.create_category_management_widgets()

    def create_category_management_widgets(self):
        # Category Management Frame
        self.category_frame = ttk.LabelFrame(self.root, text="Category Management")
        self.category_frame.grid(row=3, column=0, padx=10, pady=10, sticky="nsew")
        # Create a Treeview to display categories and file types
        self.category_tree = ttk.Treeview(self.category_frame)
        self.category_tree["columns"] = "File_Types"
        self.category_tree.column("#0", width=120, minwidth=25)
        self.category_tree.column("File_Types", width=400, minwidth=25)
        self.category_tree.heading("#0", text="Category", anchor=tk.W)
        self.category_tree.heading("File_Types", text="File Types", anchor=tk.W)
        self.category_tree.grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")

        # Scrollbar for Treeview
        self.tree_scroll = ctk.CTkScrollbar(self.category_frame, command=self.category_tree.yview,
                                            orientation="vertical")  # Changed to CTkScrollbar
        self.tree_scroll.grid(row=0, column=2, sticky='ns')
        self.category_tree.configure(yscrollcommand=self.tree_scroll.set)

        # Populate the Treeview with default categories
        self.populate_category_tree()

        self.category_name = tk.StringVar()
        self.file_types = tk.StringVar()

        ctk.CTkLabel(self.category_frame, text="Category Name:").grid(row=1,column=0,padx=10, pady=10,sticky="ew")
        ctk.CTkEntry(self.category_frame, textvariable=self.category_name).grid(row=1, column=1,padx=10, pady=10, sticky="ew")

        ctk.CTkLabel(self.category_frame, text="File Types (comma-separated):").grid(row=2, column=0,padx=10, pady=10,sticky="ew")
        ctk.CTkEntry(self.category_frame, textvariable=self.file_types).grid(row=2, column=1,padx=10, pady=10, sticky="ew")
        # Buttons for adding and removing file types
        self.add_button = ctk.CTkButton(self.category_frame, text="Add File Type", command=self.add_folder)
        self.add_button.grid(row=2, column=2, padx=10, pady=10, sticky="ew")
        self.remove_button =ctk.CTkButton(self.category_frame, text="Remove destination folder", command=self.remove_folder)
        self.remove_button.grid(row=1, column=2, padx=10, pady=10, sticky="ew")



    def populate_category_tree(self):
        # Clear the tree
        for i in self.category_tree.get_children():
            self.category_tree.delete(i)

        # Add categories and file types to the tree
        for category, file_types in categories.items():
            self.category_tree.insert("", "end", iid=category, text=category, values=([", ".join(file_types)]))

    def add_folder(self):
        category = self.category_name.get()
        file_types_list = self.file_types.get().split(',')
        if category and file_types_list:
            categories[category] = file_types_list
            self.category_name.set('')
            self.file_types.set('')
        self.populate_category_tree()



    def remove_folder(self):
        # Get the selected item from the Treeview
        selected_item = self.category_tree.selection()
        if not selected_item:
            messagebox.showerror("Error", "No category selected")
            return

        # Get the category name from the selected item
        category_name = self.category_tree.item(selected_item[0])['text']

        # Check if the category exists in the categories dictionary
        if category_name in categories:
            # Remove the category from the dictionary
            del categories[category_name]
            # Remove the category from the Treeview
            self.category_tree.delete(selected_item[0])
            messagebox.showinfo("Success", f"{category_name} category removed successfully")
        else:
            messagebox.showerror("Error", "Category not found")

    def browse_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.downloads_folder_var.set(folder_selected)

    def start_organizing(self):

        downloads_folder = self.downloads_folder_var.get()
        if not os.path.exists(downloads_folder):
            messagebox.showerror("Error", "Please select a valid downloads folder.")
            return


        self.organize_files(downloads_folder)

    def extract_file_content(self, file_path, file_ext):
        """
        Extract text content from a file based on its extension.
        """
        try:
            if file_ext == '.txt':
                with open(file_path, 'r', encoding='utf-8') as file:
                    return file.read()
            elif file_ext == '.pdf':
                with fitz.open(file_path) as doc:
                    text = ""
                    for page in doc:
                        text += page.get_text()
                    return text
            elif file_ext == '.docx':
                doc = docx.Document(file_path)
                return "\n".join(paragraph.text for paragraph in doc.paragraphs)
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while reading file {file_path}: {e}")
            return ""

    def generate_metadata(self, text_content):
        """
        Generate metadata using the OpenAI API for the given text content.
        """
        if not text_content:
            return "No content to generate metadata."

        try:
            response = self.client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[{"role": "user", "content":"Create a title, subject, tags, categories, and a summary for the following text:\n\n{"+text_content}]
            )
            return response.choices[0].message.content
        except Exception as e:
            return f"An error occurred while generating metadata: {e}"

    def organize_files(self, downloads_folder):

        """
        Organizes the files and executes metadata related methods.
        :param downloads_folder:
        :return:
        """
        organize_dict = categories
        for filename in os.listdir(downloads_folder):
            if os.path.isdir(os.path.join(downloads_folder, filename)):
                continue

            file_ext = os.path.splitext(filename)[-1].lower()
            source_file = os.path.join(downloads_folder, filename)
            """
            If it is a text file, extract text and generate metadata from gpt api.
            """
            if file_ext in ['.txt', '.pdf', '.docx']:
                file_content = self.extract_file_content(source_file, file_ext)
                metadata = self.generate_metadata(file_content)
                dest_folder = os.path.join(downloads_folder, "Documents")
                if not os.path.exists(dest_folder):
                    shutil.move(source_file, os.path.join(downloads_folder, "Documents"))
                else:
                    print(
                        f"{filename} already exists in the destination folder {dest_folder}. If you wish to move the file, please either rename the file or do so manually.")
                self.metadata_text.insert("end", f"File name: {filename}\n"
                                                 f"Metadata: {metadata})\n")
                continue  # Continue to the next file
            for folder, extensions in organize_dict.items():
                if file_ext in extensions:
                    dest_folder = os.path.join(downloads_folder, folder)
                    dest_file = os.path.join(dest_folder,filename)
                    if not os.path.exists(dest_folder):
                        os.makedirs(dest_folder)
                    if not os.path.exists(dest_file):
                        shutil.move(source_file, dest_folder)
                    else:
                        print(f"{filename} already exists in the destination folder {folder}. If you wish to move the file, please either rename the file or do so manually.")
                    break
                else:
                    dest_folder = os.path.join(downloads_folder,"Other")
                    if not os.path.exists(dest_folder):
                        os.makedirs(dest_folder)
                    shutil.move(source_file, dest_folder)

        messagebox.showinfo("Success", "Files have been organized.")


def main():
    root = tk.Tk()
    app = FileOrganizerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
