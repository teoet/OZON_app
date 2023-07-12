import tkinter as tk
from tkinter import filedialog
import subprocess
from tkinter import ttk
from queue import Queue, Empty
import threading

def select_file1():
    file1_path = filedialog.askopenfilename()
    file1_entry.delete(0, tk.END)
    file1_entry.insert(0, file1_path)

def select_file2():
    file2_path = filedialog.askopenfilename()
    file2_entry.delete(0, tk.END)
    file2_entry.insert(0, file2_path)

def select_file3():
    file3_path = filedialog.askopenfilename()
    file3_entry.delete(0, tk.END)
    file3_entry.insert(0, file3_path)

def select_file4():
    file4_path = filedialog.askopenfilename()
    file4_entry.delete(0, tk.END)
    file4_entry.insert(0, file4_path)

def read_output(output, queue):
    for line in iter(output.readline, ''):
        queue.put(line)
    output.close()

def update_output_text():
    while True:
        try:
            line = output_queue.get_nowait()
            output_text.config(state=tk.NORMAL)
            output_text.delete("1.0", "2.0")  # Keeps only the last two lines in the logbook
            output_text.insert(tk.END, line.rstrip('\n'))
            output_text.config(state=tk.DISABLED)
            output_text.update()
        except Empty:
            break
    if process.poll() is None or not output_queue.empty():
        window.after(100, update_output_text)

def toggle_run():
    global process
    if process is None or process.poll() is not None:
        start_subprocess()
    else:
        stop_subprocess()

def start_subprocess():
    file1_path = file1_entry.get()
    file2_path = file2_entry.get()
    file3_path = file3_entry.get()
    file4_path = file4_entry.get()

    global process, output_queue, run_button
    process = subprocess.Popen(['python3', 'program_opt.py', file1_path, file2_path, file3_path, file4_path],
                               stdout=subprocess.PIPE, universal_newlines=True)

    output_queue = Queue()
    output_thread = threading.Thread(target=read_output, args=(process.stdout, output_queue))
    output_thread.daemon = True
    output_thread.start()

    window.after(100, update_output_text)
    run_button.config(text="Остановить", bg="red")

def stop_subprocess():
    global process, run_button
    if process is not None and process.poll() is None:
        process.terminate()
        process = None
        run_button.config(text="Начать обновление данных", bg="SystemButtonFace")

window = tk.Tk()
window.title("Остатки и Цены")

file1_button = tk.Button(window, text="Выберите файл Ассортимента Цен", command=select_file1)
file1_button.pack()

file1_entry = tk.Entry(window)
file1_entry.pack()

file2_button = tk.Button(window, text="Выберите файл Прайс Лист", command=select_file2)
file2_button.pack()

file2_entry = tk.Entry(window)
file2_entry.pack()

file3_button = tk.Button(window, text="Выберите файл Каталог Остатков", command=select_file3)
file3_button.pack()

file3_entry = tk.Entry(window)
file3_entry.pack()

file4_button = tk.Button(window, text="Выберите файл Шаблон Остатков", command=select_file4)
file4_button.pack()

file4_entry = tk.Entry(window)
file4_entry.pack()

run_button = tk.Button(window, text="Начать обновление данных", command=toggle_run)
run_button.pack()

output_text = tk.Text(window, height=4, wrap="word", font=("Helvetica", 14))
output_text.pack()
output_text.config(state=tk.DISABLED)

process = None  # Initialize the process variable
output_queue = None  # Initialize the output_queue variable

window.mainloop()

