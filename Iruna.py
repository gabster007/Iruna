import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from datetime import datetime
import os

# Default file path for saving/loading data
DATA_FILE = "data.csv"

# Function to calculate volume and add data to the table
def add_entry():
    try:
        etiqueta = entry_etiqueta.get()
        ubicacion = combo_ubicacion.get()
        especies = combo_especies.get()
        tipo = combo_tipo.get()
        calidad = combo_calidad.get()
        cantidad = float(entry_cantidad.get())
        largo = float(entry_largo.get())
        ancho = float(entry_ancho.get())
        espesor = float(entry_espesor.get())
        date = datetime.now().strftime("%Y-%m-%d")  # Get today's date

        # Calculate volume
        if cantidad < 1:
            volumen = round(cantidad * largo * ancho, 3)
        else:
            volumen = round(cantidad * largo * ancho * espesor, 3)
        
        # Add data to the treeview
        tree.insert("", "end", values=(etiqueta, ubicacion, especies, tipo, calidad, cantidad, largo, ancho, espesor, volumen, date))
        
        # Save data to CSV
        save_data_to_csv()
        
        # Update total volume
        update_total_volume()
        
        # Clear input fields
        clear_fields()
    except ValueError:
        messagebox.showerror("Input Error", "Please enter valid numerical values for Cantidad, Largo, Ancho, and Espesor.")

# Function to delete a selected entry
def delete_entry():
    try:
        selected_item = tree.selection()[0]  # Get selected item
        tree.delete(selected_item)  # Remove from the treeview
        
        # Save data to CSV after deletion
        save_data_to_csv()
        
        # Update total volume
        update_total_volume()
        messagebox.showinfo("Delete Successful", "Selected entry has been deleted.")
    except IndexError:
        messagebox.showerror("Delete Error", "No entry selected. Please select an entry to delete.")

# Function to clear input fields
def clear_fields():
    entry_etiqueta.delete(0, tk.END)
    combo_ubicacion.set("")
    combo_especies.set("")
    combo_tipo.set("")
    combo_calidad.set("")
    entry_cantidad.delete(0, tk.END)
    entry_largo.delete(0, tk.END)
    entry_ancho.delete(0, tk.END)
    entry_espesor.delete(0, tk.END)

# Function to calculate and display total volume
def update_total_volume(filtered_data=None):
    if filtered_data is None:
        data = [tree.item(child)["values"] for child in tree.get_children()]
    else:
        data = filtered_data
    total_volume = sum(float(row[9]) for row in data)  # Volume is at index 9
    formatted_volume = "{:,.2f}".format(total_volume)  # Format with thousand separators and 2 decimals
    lbl_total_volume.config(text=f"Total Volume: {formatted_volume}")

# Function to save data to default CSV
def save_data_to_csv():
    data = [tree.item(child)["values"] for child in tree.get_children()]
    df = pd.DataFrame(data, columns=["N. Etiqueta", "Ubicación", "Espécies", "Tipo", "Calidad", 
                                     "Cantidad", "Largo", "Ancho", "Espesor", "Volumen", "Date"])
    df.to_csv(DATA_FILE, index=False)

# Function to save data to a user-specified Excel file
def save_to_excel():
    try:
        # Get file path from user
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                                 filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if not file_path:
            return  # User canceled the save dialog

        # Collect data from the table
        data = [tree.item(child)["values"] for child in tree.get_children()]
        df = pd.DataFrame(data, columns=["N. Etiqueta", "Ubicación", "Espécies", "Tipo", "Calidad", 
                                         "Cantidad", "Largo", "Ancho", "Espesor", "Volumen", "Date"])
        
        # Save the DataFrame to the specified Excel file
        df.to_excel(file_path, index=False, engine='openpyxl')
        messagebox.showinfo("Export Successful", f"Data exported successfully to {file_path}")
    except Exception as e:
        messagebox.showerror("Export Error", f"An error occurred while saving the file:\n{e}")

# Function to load data from CSV
def load_data_from_csv():
    if os.path.exists(DATA_FILE):
        df = pd.read_csv(DATA_FILE)
        for _, row in df.iterrows():
            tree.insert("", "end", values=row.tolist())
        update_total_volume()

# Function to filter data by date and update total volume
def filter_by_date():
    start_date = entry_start_date.get()
    end_date = entry_end_date.get()

    if not start_date or not end_date:
        messagebox.showerror("Input Error", "Please enter both start and end dates.")
        return

    try:
        start_date = datetime.strptime(start_date, "%Y-%m-%d")
        end_date = datetime.strptime(end_date, "%Y-%m-%d")
    except ValueError:
        messagebox.showerror("Input Error", "Dates must be in the format YYYY-MM-DD.")
        return

    filtered_data = []
    for child in tree.get_children():
        row = tree.item(child)["values"]
        row_date = datetime.strptime(row[10], "%Y-%m-%d")  # Date is at index 10
        if start_date <= row_date <= end_date:
            filtered_data.append(row)

    # Clear and display filtered data
    for child in tree.get_children():
        tree.delete(child)
    for row in filtered_data:
        tree.insert("", "end", values=row)

    # Update total volume for filtered data
    update_total_volume(filtered_data)

# Function to reset the table and total volume
def reset_filter():
    tree.delete(*tree.get_children())  # Clear the tree
    load_data_from_csv()  # Reload the data from CSV

# Create the main window
root = tk.Tk()
root.title("Iruña Sistema de Láminas")
root.geometry("1200x700")
root.configure(bg="#f7f7f7")  # Light background color

# Global font styles
font_title = ("Arial", 16, "bold")
font_label = ("Arial", 12)
font_button = ("Arial", 12, "bold")
font_total = ("Arial", 14, "bold")

# Header
header_frame = tk.Frame(root, bg="#003366", height=50)
header_frame.pack(fill=tk.X)
header_label = tk.Label(header_frame, text="Iruña Sistema de Láminas", bg="#003366", fg="white", font=font_title, pady=10)
header_label.pack()

# Input frame
frame_input = tk.Frame(root, padx=10, pady=10, bg="#f7f7f7")
frame_input.pack(fill=tk.X)

tk.Label(frame_input, text="N. Etiqueta", font=font_label, bg="#f7f7f7").grid(row=0, column=0, padx=5, pady=5)
entry_etiqueta = tk.Entry(frame_input)
entry_etiqueta.grid(row=0, column=1, padx=5, pady=5)

tk.Label(frame_input, text="Ubicación", font=font_label, bg="#f7f7f7").grid(row=0, column=2, padx=5, pady=5)
combo_ubicacion = ttk.Combobox(frame_input, values=["B20", "B22", "B23", "B24", "B27", "C2", "C4", "C23", "D5", "D06", "D07", "D09", "D10", "D11", "D12", "D13", "D14", "D15", "D16", "D17", "D18", "D19", "D20", "D21", "D25", "D26", "D27", "F12", "F19", "F20", "F21", "F22", "F23"], state="readonly")
combo_ubicacion.grid(row=0, column=3, padx=5, pady=5)

tk.Label(frame_input, text="Espécies", font=font_label, bg="#f7f7f7").grid(row=1, column=0, padx=5, pady=5)
combo_especies = ttk.Combobox(frame_input, values=["PARICÁ", "EUCALIPTO", "PINO"], state="readonly")
combo_especies.grid(row=1, column=1, padx=5, pady=5)

tk.Label(frame_input, text="Tipo", font=font_label, bg="#f7f7f7").grid(row=1, column=2, padx=5, pady=5)
combo_tipo = ttk.Combobox(frame_input, values=["Externo", "Interno"], state="readonly")
combo_tipo.grid(row=1, column=3, padx=5, pady=5)

tk.Label(frame_input, text="Calidad", font=font_label, bg="#f7f7f7").grid(row=2, column=0, padx=5, pady=5)
combo_calidad = ttk.Combobox(frame_input, values=["C1", "CAPA", "CAPA CC", "CAPA CCF", "CAPA CR", "CAPA RELLENO", "CC", "CCF", "CEPINO", "CF", "CG", "CR", "RECORTE", "RECORTE RETOZO", "RELLENO", "RETOSO", "-"], state="readonly")
combo_calidad.grid(row=2, column=1, padx=5, pady=5)

tk.Label(frame_input, text="Cantidad", font=font_label, bg="#f7f7f7").grid(row=2, column=2, padx=5, pady=5)
entry_cantidad = tk.Entry(frame_input)
entry_cantidad.grid(row=2, column=3, padx=5, pady=5)

tk.Label(frame_input, text="Largo (m)", font=font_label, bg="#f7f7f7").grid(row=3, column=0, padx=5, pady=5)
entry_largo = tk.Entry(frame_input)
entry_largo.grid(row=3, column=1, padx=5, pady=5)

tk.Label(frame_input, text="Ancho (m)", font=font_label, bg="#f7f7f7").grid(row=3, column=2, padx=5, pady=5)
entry_ancho = tk.Entry(frame_input)
entry_ancho.grid(row=3, column=3, padx=5, pady=5)

tk.Label(frame_input, text="Espesor (m)", font=font_label, bg="#f7f7f7").grid(row=4, column=0, padx=5, pady=5)
entry_espesor = tk.Entry(frame_input)
entry_espesor.grid(row=4, column=1, padx=5, pady=5)

tk.Button(frame_input, text="Add Entry", font=font_button, bg="#0066cc", fg="white", command=add_entry).grid(row=5, column=0, pady=10, sticky="ew")
tk.Button(frame_input, text="Delete Entry", font=font_button, bg="#dc3545", fg="white", command=delete_entry).grid(row=5, column=1, pady=10, sticky="ew")
tk.Button(frame_input, text="Save as Excel", font=font_button, bg="#28a745", fg="white", command=save_to_excel).grid(row=5, column=2, pady=10, sticky="ew")

# Data table
frame_table = tk.Frame(root, padx=10, pady=10, bg="#f7f7f7")
frame_table.pack(fill=tk.BOTH, expand=True)

columns = ["N. Etiqueta", "Ubicación", "Espécies", "Tipo", "Calidad", "Cantidad", "Largo", "Ancho", "Espesor", "Volumen", "Date"]
tree = ttk.Treeview(frame_table, columns=columns, show="headings", height=15)
tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

for col in columns:
    tree.heading(col, text=col)
    tree.column(col, width=100, anchor=tk.CENTER)

scrollbar = ttk.Scrollbar(frame_table, orient=tk.VERTICAL, command=tree.yview)
tree.configure(yscrollcommand=scrollbar.set)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

# Filter controls
frame_filter = tk.Frame(root, padx=10, pady=10, bg="#f7f7f7")
frame_filter.pack(fill=tk.X)

tk.Label(frame_filter, text="Start Date (YYYY-MM-DD):", font=font_label, bg="#f7f7f7").grid(row=0, column=0, padx=5, pady=5)
entry_start_date = tk.Entry(frame_filter)
entry_start_date.grid(row=0, column=1, padx=5, pady=5)

tk.Label(frame_filter, text="End Date (YYYY-MM-DD):", font=font_label, bg="#f7f7f7").grid(row=0, column=2, padx=5, pady=5)
entry_end_date = tk.Entry(frame_filter)
entry_end_date.grid(row=0, column=3, padx=5, pady=5)

tk.Button(frame_filter, text="Filter by Date", font=font_button, bg="#ffc107", command=filter_by_date).grid(row=0, column=4, padx=5, pady=5)
tk.Button(frame_filter, text="Reset Filter", font=font_button, bg="#dc3545", fg="white", command=reset_filter).grid(row=0, column=5, padx=5, pady=5)

# Total Volume Label Frame
frame_total = tk.Frame(root, padx=10, pady=5, bg="#f7f7f7")
frame_total.pack(fill=tk.X)


# Add Total Volume Label to the Right
lbl_total_volume = tk.Label(
    frame_input,
    text="Total Volume: 0.00",
    font=("Arial", 20, "bold"),  # Match font style to input labels
    bg="#f7f7f7",
    fg="#003366",
    anchor="e"
)
lbl_total_volume.grid(row=5, column=15, padx=100, sticky="e")  # Align to the right

# Load data on startup
load_data_from_csv()

# Run the application
root.mainloop() 