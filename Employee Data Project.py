import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import pyodbc
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import random

# ---------------------- SQL CONNECTION ----------------------
def get_sql_connection():
    server = 'LAPTOP-TRPT95L8\\MSSQLSERVER01'
    database = 'Class1'
    connection_string = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};Trusted_Connection=yes;'
    return pyodbc.connect(connection_string)

# ---------------------- VIEW DATA ----------------------
def show_data():
    conn = get_sql_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM Employes")
    rows = cursor.fetchall()
    columns = [column[0] for column in cursor.description]
    conn.close()

    window = tk.Toplevel()
    window.title("Employee Data")

    tree = ttk.Treeview(window, columns=columns, show='headings')
    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, width=100)

    for row in rows:
        tree.insert("", tk.END, values=list(row))

    tree.pack(expand=True, fill=tk.BOTH)

# ---------------------- GRAPH GENERATION ----------------------
def show_graphs():
    def generate_graphs():
        selected_columns = [col for col, var in zip(columns, col_vars) if var.get()]

        if not selected_columns:
            messagebox.showwarning("Warning", "Select at least one column.")
            return

        conn = get_sql_connection()
        df = pd.read_sql("SELECT * FROM Employes", conn)
        conn.close()

        fig, ax = plt.subplots(figsize=(9, 6))
        graph_type = selected_graph_type.get()

        if graph_type == "Pie Chart":
            if len(selected_columns) != 1:
                messagebox.showwarning("Warning", "Please select one column for Pie Chart.")
                return
            col = selected_columns[0]
            if not pd.api.types.is_numeric_dtype(df[col]):
                messagebox.showerror("Error", f"{col} must be numeric.")
                return
            pie_data = df.groupby('Department')[col].sum()
            ax.pie(pie_data, labels=pie_data.index, autopct='%1.1f%%')
            ax.set_title(f"{col} Distribution by Department")

        elif graph_type == "Line Chart":
            if len(selected_columns) < 2:
                messagebox.showwarning("Warning", "Select at least 2 columns.")
                return
            x_col = selected_columns[0]
            y_cols = selected_columns[1:]
            for y_col in y_cols:
                ax.plot(df[x_col], df[y_col], marker='o', label=f"{y_col} vs {x_col}")
            ax.legend()
            ax.set_title(" vs. ".join(selected_columns))
            ax.set_ylabel('Values')
            ax.set_xlabel(x_col)

        elif graph_type == "Bar Chart":
            bar_data = df.groupby('Department')[selected_columns].sum()
            bar_data.plot(kind='bar', ax=ax)
            ax.set_title(f"Bar Chart: {', '.join(selected_columns)} by Department")
            ax.set_ylabel('Values')
            ax.set_xlabel('Department')
            ax.legend()

        ax.grid(True)
        canvas = FigureCanvasTkAgg(fig, master=window)
        canvas.get_tk_widget().pack()
        canvas.draw()

    window = tk.Toplevel()
    window.title("Graph Options")
    window.configure(bg="#e8f4f8")

    tk.Label(window, text="Select columns to graph:", bg="#e8f4f8", font=("Arial", 12)).pack(anchor='w')

    conn = get_sql_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM Employes")
    columns = [column[0] for column in cursor.description if column[0] != 'Department']
    conn.close()

    col_vars = [tk.BooleanVar() for _ in columns]
    for i, col in enumerate(columns):
        tk.Checkbutton(window, text=col, variable=col_vars[i], bg="#e8f4f8").pack(anchor='w')

    tk.Label(window, text="Select Graph Type:", bg="#e8f4f8", font=("Arial", 12)).pack(anchor='w')
    selected_graph_type = tk.StringVar(value="Line Chart")
    for graph in ["Line Chart", "Bar Chart", "Pie Chart"]:
        tk.Radiobutton(window, text=graph, variable=selected_graph_type, value=graph, bg="#e8f4f8").pack(anchor='w')

    tk.Button(window, text="Generate Graph", command=generate_graphs, bg="#87cefa", font=("Arial", 10, "bold")).pack(pady=10)

# ---------------------- SQL NOTEPAD ----------------------
def sql_notepad():
    def execute_query():
        query = text_area.get("1.0", tk.END).strip()
        if not query:
            messagebox.showwarning("Warning", "Query is empty!")
            return
        try:
            conn = get_sql_connection()
            cursor = conn.cursor()
            cursor.execute(query)
            conn.commit()
            conn.close()
            messagebox.showinfo("Success", "Executed successfully.")
        except Exception:
            messagebox.showerror("Error", "Execution failed.")

    window = tk.Toplevel()
    window.title("SQL Notepad")
    window.configure(bg="#f0f8ff")
    text_area = scrolledtext.ScrolledText(window, wrap=tk.WORD, width=60, height=15, bg="#fdfdfe")
    text_area.pack(padx=10, pady=10)
    tk.Button(window, text="Execute Query", command=execute_query, bg="#87cefa", font=("Arial", 10, "bold")).pack()

# ---------------------- MAIN GUI ----------------------
def main():
    root = tk.Tk()
    root.title("Employee Data Manager")
    root.geometry("800x600")
    root.configure(bg="#001f2d")

    canvas = tk.Canvas(root, bg="#001f2d", highlightthickness=0)
    canvas.pack(fill=tk.BOTH, expand=True)

    def draw_background():
        canvas.delete("all")
        width = root.winfo_width()
        height = root.winfo_height()

        for _ in range(200):
            x = random.randint(0, width)
            y = random.randint(0, height)
            canvas.create_oval(x, y, x+2, y+2, fill="#00ffcc", outline="")

        for _ in range(50):
            x1 = random.randint(0, width)
            y1 = random.randint(0, height)
            x2 = x1 + random.randint(-100, 100)
            y2 = y1 + random.randint(-100, 100)
            canvas.create_line(x1, y1, x2, y2, fill="#00ffcc", width=1)

        canvas.lower()

    root.after(100, draw_background)
    root.bind("<Configure>", lambda e: draw_background())

    content = tk.Frame(root, bg="#001f2d")
    content.place(relx=0.5, rely=0.5, anchor="center")

    tk.Label(content, text="Employee Data Management", font=("Arial", 18, "bold"),
             bg="#001f2d", fg="#00ffff").pack(pady=30)

    style = {'font': ("Arial", 12), 'width': 25, 'bg': "#004f66", 'fg': 'white'}

    buttons = [
        ("View Employee Data", show_data),
        ("Generate Graphs", show_graphs),
        ("SQL Notepad", sql_notepad),
        ("Exit Application", root.destroy)
    ]

    for text, cmd in buttons:
        tk.Button(content, text=text, command=cmd, **style).pack(pady=10)

    tk.Label(content, text="Powered by Mridul", font=("Arial", 9, "italic"),
             bg="#001f2d", fg="#00ffcc").pack(pady=20)

    root.mainloop()

if __name__ == "__main__":
    main()
