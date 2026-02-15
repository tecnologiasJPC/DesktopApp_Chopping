'''
To create the corresponding exe file it is required to execute the next command
pyinstaller --onefile -w -n Chopping_v1.2 --add-data "folder.ico;." --add-data "button.png;." --add-data "icono.ico;." --icon=icono.ico main.py
'''

import ctypes
import datetime
import glob
import os
import subprocess
import sys
import time
import tkinter as tk
import webbrowser
import cv2
import pyautogui
import pyperclip
import pytesseract
from PIL import Image, ImageTk, ImageGrab

# it is required to download tesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'

# this adapts dimensions of screen
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)  # Windows 8.1+
except Exception as e:
    try:
        ctypes.windll.user32.SetProcessDPIAware()  # Windows 7
    except:
        pass


# here the transparent window is displayed over screen
class RectOverlay:
    def __init__(self, master, on_close_callback):
        self.master = master
        self.on_close_callback = on_close_callback
        self.process = True

        # variables to get the current monitor size
        user32 = ctypes.windll.user32
        self.virtual_width = user32.GetSystemMetrics(78)
        self.virtual_height = user32.GetSystemMetrics(79)
        self.virtual_left = user32.GetSystemMetrics(76)
        self.virtual_top = user32.GetSystemMetrics(77)

        # this creates an transparent window over the screen
        self.overlay = tk.Toplevel(master)
        self.overlay.geometry(f"{self.virtual_width}x{self.virtual_height}+{self.virtual_left}+{self.virtual_top}")
        self.overlay.attributes('-alpha', 0.3)
        self.overlay.attributes('-topmost', True)
        self.overlay.configure(bg='black')
        self.overlay.overrideredirect(True)

        # changes the cursor aspect
        self.canvas = tk.Canvas(self.overlay, cursor="cross", bg='black', highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)

        # defines the initial values for the required attributes
        if getattr(sys, 'frozen', False):   # For executable it gets the current route
            self.main_route = os.path.dirname(sys.executable)
        else:
            self.main_route = os.path.dirname(__file__)   # for development this is the route
        self.start_x, self.start_y = None, None
        self.rect = None
        self.text_found = None
        self.qr_found = None
        self.section = None
        self.current_name = None

        # defines the events based on keys pressed
        self.canvas.bind("<ButtonPress-1>", self.on_click)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)
        self.canvas.bind_all("<Escape>", lambda e: self.close)

    def on_click(self, event):
        self.start_x = self.canvas.canvasx(event.x)
        self.start_y = self.canvas.canvasy(event.y)
        self.rect = self.canvas.create_rectangle(self.start_x, self.start_y, self.start_x, self.start_y, outline='red', width=2)
        time.sleep(0.05)

    def on_drag(self, event):
        cur_x = self.canvas.canvasx(event.x)
        cur_y = self.canvas.canvasy(event.y)
        self.canvas.coords(self.rect, self.start_x, self.start_y, cur_x, cur_y)

    def on_release(self, event):
        end_x = self.canvas.canvasx(event.x)
        end_y = self.canvas.canvasy(event.y)

        # in case the start point is different to top upper
        abs_start_x = int(self.start_x + self.virtual_left)
        abs_start_y = int(self.start_y + self.virtual_top)
        abs_end_x = int(end_x + self.virtual_left)
        abs_end_y = int(end_y + self.virtual_top)

        left = min(abs_start_x, abs_end_x)
        top = min(abs_start_y, abs_end_y)
        right = max(abs_start_x, abs_end_x)
        bottom = max(abs_start_y, abs_end_y)

        if right > left and bottom > top:
            part = [left, top, right, bottom]
            self.section = tuple(part)
            self.process = True
        else:
            print("Invalid rectangle, capture is not saved.")
            self.process = False
        self.close()

    def qr_analyze(self):
        qr_detector = cv2.QRCodeDetector()
        try:
            img = cv2.imread(os.path.join(self.main_route, self.current_name))
            q_data, bbox, _ = qr_detector.detectAndDecode(img)
            if q_data:
                print(f"QR detected: {q_data}")
                self.qr_found = str(q_data)
                return True
            else:
                self.qr_found = None
                print("No QR code is detected")
                return False
        except Exception as e:
            self.qr_found = None
            print(f"Error reading QR code: {e}")
            return False

    def text_analyze(self):     # analyze the text found in the image
        img_punt = Image.open(os.path.join(self.main_route, self.current_name))
        text = pytesseract.image_to_string(img_punt, config='--psm 11 --oem 3')
        text = text.replace('\n', ' ').replace('  ', ' ')
        pyperclip.copy(text)
        text = text[:-1]
        self.text_found = text
        print('Text found:', self.text_found)

    def close(self, data=None):
        self.overlay.destroy()
        time.sleep(0.1)
        if self.process:
            now = datetime.datetime.now()
            stamp = now.strftime("%Y-%m-%d %H_%M_%S")
            screenshot = ImageGrab.grab(self.section, all_screens=True)

            self.current_name = "captures/capture" + stamp + ".png"
            screenshot.save(os.path.join(self.main_route, self.current_name))

            if self.qr_analyze():
                self.text_found = self.qr_found
            else:
                self.text_analyze()

        if self.on_close_callback:
            self.on_close_callback({"State": self.process, "Text": self.text_found, "QR": self.qr_found,
                                    "Image": self.current_name, "Section": self.section})


# this is the main GUI displayed since the beginning
class MainGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Chopping v1.2")
        if getattr(sys, 'frozen', False):   # For executable it gets the current route
            self.route = os.path.join(os.path.dirname(sys.executable), 'captures')
        else:
            self.route = os.path.dirname(__file__) + '\\captures\\'   # for development this is the route
        if not os.path.exists(self.route):
            os.makedirs(self.route, exist_ok=True)
        self.root.iconbitmap(os.path.join(os.path.dirname(__file__), 'icono.ico'))

        user_ = ctypes.windll.user32
        v_size = user_.GetSystemMetrics(79)  # vertical size
        v = int(v_size * 0.13)
        h = int(v * 3)
        self.initial_window = str(h) + 'x' + str(v)
        self.root.geometry(self.initial_window)    # initial size of window
        self.root.minsize(400, 125)         # fixed size of window
        self.root.resizable(False, False)   # window is not resizable
        self.frame = tk.Frame(self.root)
        self.frame.pack(pady=10)

        # button to start selection
        img_chop = Image.open(os.path.join(os.path.dirname(__file__), "button.png"))
        img_chop = img_chop.resize((147, 50))
        chop_button = ImageTk.PhotoImage(img_chop)
        self.button = tk.Button(self.frame, image=chop_button, text="Draw", command=self.launch_overlay)
        self.button.image = chop_button
        self.button.pack(side=tk.LEFT)

        # button to access images taken
        img = Image.open(os.path.join(os.path.dirname(__file__), "folder.ico"))
        img = img.resize((50, 50))
        folder_button = ImageTk.PhotoImage(img)
        self.buttonRoute = tk.Button(self.frame, image=folder_button, text="Folder", command=self.open_location)
        self.buttonRoute.image = folder_button
        self.buttonRoute.pack(side=tk.LEFT)

        # shows the position of cursor on screen
        self.labelCoordinates = tk.Label(self.root, text="")
        self.labelCoordinates.pack()

        # section where captured image is displayed
        self.labelImg = tk.Label(self.root)

        # section that shows the x1, y1, x2, y2 section taken
        self.textLocation = tk.Entry(self.root, state='readonly', bg="#f0f0f0",
                                     fg='black', justify="center", relief='flat')

        # show the instructions text
        self.labelInfo = tk.Label(self.root, text="Select the area to analyze")
        self.labelInfo.pack()

        # author info
        self.author = tk.Label(self.root, text="A software from tecnologiasJPC company")

        # create the text analyzed from the image
        self.frame_text = tk.Frame(self.root)
        # Create the vertical scrollbar
        self.scrollbar = tk.Scrollbar(self.frame_text)
        # Create the text widget
        self.text_box = tk.Text(
            self.frame_text,
            wrap=tk.WORD,  # Wrap at word boundaries
            yscrollcommand=self.scrollbar.set,  # Connect to scrollbar
            width=50,
            height=20
        )
        # configure the scrollbar to work with the text widget
        self.scrollbar.config(command=self.text_box.yview)

    # get the current position of cursor and updates every 100 ms
    def update_coordinates(self):
        x, y = pyautogui.position()
        self.labelCoordinates.config(text=f"X: {x} Y: {y}")
        self.root.after(100, self.update_coordinates)  # update every 100 milliseconds

    # creates a transparent window over the screen
    def launch_overlay(self):
        self.root.withdraw()
        RectOverlay(self.root, on_close_callback=self.show_main)

    # opens the folder where the captures are saved
    def open_location(self):
        subprocess.run(['start', self.route], shell=True, check=True)

    # this is called once the first area selection is done
    def show_main(self, data):

        def open_link(event):
            url = self.text_box.get("1.0", "end-1c").strip()
            chrome_path = "C:/Program Files/Google/Chrome/Application/chrome.exe"
            if os.path.exists(chrome_path):
                webbrowser.register('chrome', None, webbrowser.BackgroundBrowser(chrome_path))
                webbrowser.get('chrome').open(url)
            else:
                webbrowser.open(url)

        self.root.deiconify()   # shows again the original window

        if data["State"]:

            self.frame_text.pack_forget()  # decoded text hidden
            self.scrollbar.pack_forget()   # scrollbar hidden
            self.text_box.pack_forget()  # textbox hidden
            self.labelInfo.pack_forget()  # info text hidden

            files = glob.glob(os.path.join(self.route, '*.png'))
            most_recent = max(files, key=os.path.getmtime)
            image = Image.open(most_recent)
            w, h = image.size

            window_width = 400 if w < 400 else w
            window_height = h + 300

            self.root.geometry(str(window_width) + "x" + str(window_height))
            self.root.minsize(window_width, window_height)
            self.root.resizable(True, True)

            capture = ImageTk.PhotoImage(image)
            self.labelImg.config(image=capture)
            self.labelImg.image = capture
            self.labelImg.pack(padx=5)

            lugar = "Location " + str(data["Section"])
            self.textLocation.config(state='normal')
            self.textLocation.delete(0, tk.END)
            self.textLocation.insert(0, lugar)
            self.textLocation.config(state='readonly')
            self.textLocation.pack(pady=5, padx=10, fill=tk.X)

            if data["QR"] is None:
                self.labelInfo.config(text="Text found")
                print('Text was found')
            else:
                self.labelInfo.config(text="QR/bar code detected")
            self.labelInfo.pack(anchor='w')

            # create a label with creator info
            self.author.pack(side=tk.BOTTOM, anchor='e')

            # create a container for Text and Scrollbar
            self.frame_text.pack(fill=tk.BOTH, expand=True)  # it expands to top
            # move text and scrollbar to container
            self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            self.text_box.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            self.text_box.config(state=tk.NORMAL)
            self.text_box.delete("1.0", tk.END)

            message = data["Text"]

            if message.startswith(("http://", "https://", "www.")):
                self.text_box.insert(tk.END, message, ("link", message))
                self.text_box.tag_config("link", foreground="blue", underline=1)
                self.text_box.tag_bind("link", "<Button-1>", open_link)
                self.text_box.tag_bind("link", "<Enter>", lambda e: self.text_box.config(cursor="hand2"))
                self.text_box.tag_bind("link", "<Leave>", lambda e: self.text_box.config(cursor=""))
            else:
                self.text_box.insert(tk.END, message)
            self.text_box.config(state=tk.DISABLED)     # this avoid to edit text
        else:
            self.root.geometry(self.initial_window)
            self.root.minsize(400, 125)  # fixed size of window
            self.root.resizable(False, False)  # window is not resizable
            self.labelInfo.config(text="Capture not taken, try again")
            self.labelInfo.pack(anchor='center')
            self.labelImg.pack_forget()  # previous image is hidden
            self.textLocation.pack_forget()  # previous location is hidden
            self.author.pack_forget()   # creator label hidden
            self.frame_text.pack_forget()  # decoded text hidden
            self.scrollbar.pack_forget()   # scrollbar hidden
            self.text_box.pack_forget()  # textbox hidden

    # initialize the first GUI
    def run(self):
        self.update_coordinates()
        self.root.mainloop()


if __name__ == "__main__":
    MainGUI().run()
