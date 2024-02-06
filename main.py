import tkinter as tk
import tkinter.ttk as ttk
from booking_report_function import on_file_selected
from queries import new_function
import os


# Completed


def open_powerpoint():
    powerpoint_file = r'"P:\Partnership_Python_Projects\Booking Report Program\Booking Report Program Instructions.pptx"'

    try:
        os.startfile(powerpoint_file)
    except Exception as e:
        print(f"Error: {e}")


def setup_gui(main_window):
    style = ttk.Style()
    main_window.configure(bg="white")
    style.configure("TButton", font=("Roboto", 16, "bold"), width=40, height=40)
    style.map("TButton", foreground=[('active', 'white')], background=[('active', '#007BFF')])

    title_label = ttk.Label(main_window, text="Welcome Partnership Team!",
                            font=("Segoe UI", 36, "underline"), background="white", foreground="#103d81")
    title_label.pack(pady=20)

    description_label = ttk.Label(main_window,
                                  text="This tool allows you to run your Booking Report Automatically\n"
                                       "Also adds several important columns with formulas!\n\n"
                                       "1. Select The Run Query Button to auto run the query.\n "
                                       "2. Once you Ran and Saved your Query, close the file\n"
                                       "and then use the 'Select Booking Report Button' to finish the process.",
                                  font=("Roboto", 18), background="white", anchor="center",
                                  justify="center")
    description_label.pack(pady=20)

    open_powerpoint_button = ttk.Button(main_window, text='Open PowerPoint Instructions',
                                        command=open_powerpoint)
    open_powerpoint_button.pack(pady=10)

    run_queries_button = ttk.Button(main_window, text="Run Booking Report Query",
                                    command=new_function, style="TButton")
    run_queries_button.pack(pady=20)

    open_report_button = ttk.Button(main_window, text="Select Booking Report",
                                    command=on_file_selected, style="TButton")
    open_report_button.pack(pady=20)


if __name__ == "__main__":
    root = tk.Tk()
    root.title('Creation File Setup')
    root.geometry('1000x580')
    setup_gui(root)
    root.mainloop()
