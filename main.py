import tkinter as tk
import ctypes
import pyperclip
from PIL import ImageGrab
from PIL import Image
from pyzbar.pyzbar import decode
from textblob import TextBlob
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'

# this adapts dimensions of screen
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)  # Windows 8.1+
except:
    try:
        ctypes.windll.user32.SetProcessDPIAware()  # Windows 7
    except:
        pass


# here the transparent window is displayed over screen
class RectOverlay:
    def __init__(self, master, on_close_callback):
        self.master = master
        self.on_close_callback = on_close_callback

        user32 = ctypes.windll.user32
        self.virtual_width = user32.GetSystemMetrics(78)
        self.virtual_height = user32.GetSystemMetrics(79)
        self.virtual_left = user32.GetSystemMetrics(76)
        self.virtual_top = user32.GetSystemMetrics(77)

        self.overlay = tk.Toplevel(master)
        self.overlay.geometry(f"{self.virtual_width}x{self.virtual_height}+{self.virtual_left}+{self.virtual_top}")
        self.overlay.attributes('-alpha', 0.3)
        self.overlay.attributes('-topmost', True)
        self.overlay.configure(bg='black')
        self.overlay.overrideredirect(True)

        self.canvas = tk.Canvas(self.overlay, cursor="cross", bg='black', highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)

        self.start_x = None
        self.start_y = None
        self.rect = None
        self.text_found = None
        self.qr_found = None

        self.canvas.bind("<ButtonPress-1>", self.on_click)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)
        self.overlay.bind("<Escape>", lambda e: self.close())

    def on_click(self, event):
        self.start_x = self.canvas.canvasx(event.x)
        self.start_y = self.canvas.canvasy(event.y)
        self.rect = self.canvas.create_rectangle(self.start_x, self.start_y, self.start_x, self.start_y, outline='red', width=2)

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

        part = (left, top, right, bottom)
        screenshot = ImageGrab.grab(part, all_screens=True)
        screenshot.save("capture.png")
        if self.qr_analyze():
            self.text_found = self.qr_found
        else:
            self.text_analyze()
        self.close()

    def qr_analyze(self):
        try:
            img = Image.open('capture.png')
            decoded_objects = decode(img)
            if decoded_objects:
                self.qr_found = decoded_objects[0].data.decode('utf-8')
                print('QR detected:', decoded_objects[0].data.decode('utf-8'))
                return True
            else:
                self.qr_found = None
                print("No QR code is detected")
                return False
        except Exception as e:
            self.qr_found = None
            print(f"Error reading QR code: {e}")
            return False

    def text_analyze(self):
        img_punt = Image.open('capture.png')
        text = pytesseract.image_to_string(img_punt, config='--psm 11 --oem 3')
        text = text.replace('\n', ' ').replace('  ', ' ')
        pyperclip.copy(text)
        text = text[:-1]

        blob = TextBlob(text)
        text = str(blob.correct())

        self.text_found = text
        print('Text found:', self.text_found)

    def close(self, data=None):
        self.overlay.destroy()
        if self.on_close_callback:
            self.on_close_callback({"Text": self.text_found, "QR": self.qr_found})


# this is the main GUI displayed since the beginning
class MainGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Chopping v1.0")
        self.root.iconbitmap('icon.ico')
        self.root.geometry("400x75")
        self.root.minsize(400, 75)

        self.button = tk.Button(self.root, text="Draw", command=self.launch_overlay, width=25, height=2)
        self.button.pack()

        self.label1 = tk.Label(self.root, text="Select the area to analyze")
        self.label1.pack()

        # Create the vertical scrollbar
        self.scrollbar = tk.Scrollbar(self.root)

        # Create the text widget
        self.text_box = tk.Text(
            self.root,
            wrap=tk.WORD,  # Wrap at word boundaries
            yscrollcommand=self.scrollbar.set,  # Connect to scrollbar
            width=50,
            height=20
        )

        # Configure the scrollbar to work with the text widget
        self.scrollbar.config(command=self.text_box.yview)

    def launch_overlay(self):
        self.root.withdraw()
        RectOverlay(self.root, on_close_callback=self.show_main)

    def show_main(self, data):
        self.root.deiconify()
        if data:
            enunciado = data["Text"]
            codigo = data["QR"]
            if codigo is None:
                self.label1.config(text="Text found")
            else:
                self.label1.config(text="QR code detected")
            self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            self.text_box.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            self.root.geometry("400x200")
            self.root.minsize(400, 200)
            self.text_box.delete("1.0", tk.END)
            self.text_box.insert(tk.END, enunciado)

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    MainGUI().run()
