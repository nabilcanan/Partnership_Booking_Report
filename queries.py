import time
import datetime
from tkinter import messagebox, Tk, simpledialog
from pywinauto.application import Application
import pyautogui
import os



def log_user_activity(username):
    with open("user_activity_log.txt", "a") as log_file:  # 'a' mode ensures we're appending and not overwriting
        log_file.write(f"{username} logged in at {datetime.datetime.now()}\n")


def click_button_image(image_path, confidence=0.8, offset=0, double_click_required=False):
    try:
        print(f"Looking for image '{image_path}' on screen...")
        location = pyautogui.locateOnScreen(image_path, confidence=confidence)

        if 'WHERETOCLICKIMG4' in image_path:
            custom_offset_x = 618  # Replace with actual x offset
            custom_offset_y = 522  # Replace with actual y offset
            pyautogui.doubleClick(location[0] + custom_offset_x, location[1] + custom_offset_y)

        else:
            # For other images, find the center and then apply any offset
            center = pyautogui.center(location)
            new_click_location = (center[0] + offset, center[1])

            if double_click_required:
                pyautogui.doubleClick(new_click_location)
            else:
                pyautogui.click(new_click_location)

        print(f"Successfully clicked on '{image_path}'.")
        time.sleep(3)  # wait a bit after clicking

    except TypeError:
        print(f"Image '{image_path}' not found on screen!")


def get_user_credentials():  # This will gaiter the user credentials we need to log in to peoplesoft
    root = Tk()
    root.withdraw()

    username = simpledialog.askstring("Username", "Please enter your PeopleSoft username:", parent=root)

    if username is None:  # Check if user pressed Cancel for username
        messagebox.showinfo("Cancelled", "Operation was cancelled.")
        root.destroy()
        return None, None

    password = simpledialog.askstring("Password", "Please enter your PeopleSoft password:", parent=root, show='*')

    if password is None:  # Check if user pressed Cancel for password
        messagebox.showinfo("Cancelled", "You have cancelled the Run Queries Process.")
        root.destroy()
        return None, None

    root.destroy()  # Close the tkinter window after getting input

    return username, password


def new_function():
    # Modify these paths to account for the new subdirectory
    base_directory = os.path.dirname(os.path.abspath(__file__))
    image_folder_path = os.path.join(base_directory, 'images-videos')

    image_path_run_to_excel = os.path.join(image_folder_path, 'run_to_excel.png')

    query1 = "PUBLIC.QUERY.STRATEGIC_BOOKINGS"

    username, password = get_user_credentials()

    # Check if credentials were provided
    if username is None or password is None:
        print("Operation was cancelled.")
        return

    def login_and_run_query(query):
        print(f"Starting '{query}'...")

        # Launch and Connect to PeopleSoft
        app = Application().start(r'C:\FS760\bin\CLIENT\WINX86\pstools.exe')
        signon_window = app['PeopleSoft Signon']
        signon_window.wait('ready', timeout=20)

        # Use the previously gathered credentials
        username_field = signon_window.child_window(class_name="Edit", found_index=1)
        username_field.set_focus().type_keys(username, with_spaces=True)

        password_field = signon_window.child_window(class_name="Edit", found_index=2)
        password_field.set_focus().type_keys(password, with_spaces=True)

        signon_window.child_window(title="OK", class_name="Button").click()
        time.sleep(5)  # Switch back to 8

        # Check if login failed
        if app.window(title="Network API").exists():
            messagebox.showerror("Error", "PeopleSoft login failed. Please check your credentials NOW!.")
            raise Exception("LoginFailed")  # Raise an exception when login fails
        else:
            log_user_activity(username)  # Log the user's activity when the login is successful

        # Go to Query menu
        app = Application().connect(title_re="Application Designer - .*")
        app.top_window().menu_select("Go->PeopleTools->Query")
        app.top_window().close()
        time.sleep(8)  # ORIGINALLY 5

        # Open Query
        query_app = Application().connect(title="Untitled - Query")  # Modified line
        toolbar = query_app.top_window().child_window(class_name="ToolbarWindow32")
        toolbar.button(1).click_input()
        time.sleep(8)  # ORIGINALLY 5

        open_query_app = Application().connect(title="Open Query")
        open_query_window = open_query_app['Open Query']

        open_query_window.Edit.set_text(query)
        print(f"Query text for '{query}' set. Waiting a bit before clicking OK...")
        time.sleep(8)
        open_query_window.OK.click_input()
        time.sleep(3)

        # Marker to check if run_to_excel has been clicked
        run_to_excel_clicked = False

        def click_run_to_excel():
            nonlocal run_to_excel_clicked
            if not run_to_excel_clicked:
                # Look for the image and click it
                click_button_image(image_path_run_to_excel)
                run_to_excel_clicked = True
                time.sleep(17)  # Wait for 20 seconds after clicking

        click_run_to_excel()

        try:
            if query_app.window(title_re=".*Query.*").exists():
                print("Closing query window...")
                query_app.window(title_re=".*Query.*").close()
                print("Query window closed.")
        except Exception as es:
            print(f"Error while closing the query window: {str(es)}")

    try:
        # Use a counter to keep track of completed queries
        queries_completed = 0

        def on_query_completed():
            nonlocal queries_completed
            queries_completed += 1
            print(f"Completed {queries_completed} queries.")
            if queries_completed == 5:
                messagebox.showinfo("Queries Completed",
                                    "Both queries have been executed. Please allow some time for Excel sheets to load.")

        # Run both queries
        print("Running first query...")
        login_and_run_query(query1)

        on_query_completed()

    except Exception as e:
        if str(e) == "LoginFailed":  # Check for the custom exception
            print("Login failed. Stopping further queries.")
        else:
            print(f"Error: {str(e)}")
