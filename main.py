import tkinter as tk
import ctypes
import pyperclip
import pytesseract
import os
import datetime
import glob
import webbrowser
from PIL import Image, ImageTk, ImageGrab
from pyzbar.pyzbar import decode
from textblob import TextBlob

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
        self.section = None

        self.current_name = None

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

        part = [left, top, right, bottom]
        self.section = tuple(part)
        self.close()

    def qr_analyze(self):
        try:
            img = Image.open(os.path.join(os.path.dirname(__file__), self.current_name))
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
        img_punt = Image.open(os.path.join(os.path.dirname(__file__), self.current_name))
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
        now = datetime.datetime.now()
        stamp = now.strftime("%Y-%m-%d %H_%M_%S")
        screenshot = ImageGrab.grab(self.section, all_screens=True)
        self.current_name = "captures/capture" + stamp + ".png"
        screenshot.save(os.path.join(os.path.dirname(__file__), self.current_name))

        if self.qr_analyze():
            self.text_found = self.qr_found
        else:
            self.text_analyze()

        if self.on_close_callback:
            self.on_close_callback({"Text": self.text_found, "QR": self.qr_found})


# this is the main GUI displayed since the beginning
class MainGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Chopping v1.0")
        self.route = os.path.dirname(__file__) + '/captures/'
        icon_path = os.path.join(os.path.dirname(__file__), 'icon.ico')
        self.root.iconbitmap(icon_path)
        self.root.geometry("400x95")
        self.root.minsize(400, 95)

        self.frame = tk.Frame(self.root)
        self.frame.pack(pady=10)

        self.button = tk.Button(self.frame, text="Draw", command=self.launch_overlay, width=25, height=2)
        self.button.pack(side=tk.LEFT)

        img = Image.open(os.path.join(os.path.dirname(__file__), "folder.ico"))
        img = img.resize((49, 49), Image.LANCZOS)
        photo = ImageTk.PhotoImage(img)
        self.buttonRoute = tk.Button(self.frame, image=photo, text="Folder", command=self.open_location)
        self.buttonRoute.image = photo
        self.buttonRoute.pack(side=tk.LEFT)

        self.labelImg = tk.Label(self.root)

        '''try:
            image = Image.open(os.path.join(os.path.dirname(__file__), 'capture.png'))  # Cambia por tu ruta
            #image = image.resize((400, 300))  # (Ancho, Alto)
            capture = ImageTk.PhotoImage(image)
            self.labelImg = tk.Label(self.root, image=capture)
            self.labelImg.image = capture
            #self.labelImg.pack(pady=10)
        except FileNotFoundError:
            print("No se encuentra imagen")'''

        self.label1 = tk.Label(self.root, text="Select the area to analyze")
        self.label1.pack()

        self.author = tk.Label(self.root, text="A software from tecnologiasJPC company")

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
        # Configure the scrollbar to work with the text widget
        self.scrollbar.config(command=self.text_box.yview)

    def launch_overlay(self):
        self.root.withdraw()
        RectOverlay(self.root, on_close_callback=self.show_main)

    def open_location(self):
        os.startfile(self.route)

    def show_main(self, data):  # this is called once the selection is done

        def open_link(event):
            url = self.text_box.get("1.0", "end-1c").strip()

            chrome_path = "C:/Program Files/Google/Chrome/Application/chrome.exe"
            if os.path.exists(chrome_path):
                webbrowser.register('chrome', None, webbrowser.BackgroundBrowser(chrome_path))
                webbrowser.get('chrome').open(url)
            else:
                webbrowser.open(url)

        self.root.deiconify()

        if self.label1.cget("text") == "Select the area to analyze":
            self.label1.pack_forget()

        files = glob.glob(os.path.join(self.route, '*.png'))
        most_recent = max(files, key=os.path.getmtime)
        image = Image.open(most_recent)
        w, h = image.size
        if w < 400:
            window_width = 400
        else:
            window_width = w
        window_height = h + 200

        capture = ImageTk.PhotoImage(image)
        self.labelImg.config(image=capture)
        self.labelImg.image = capture
        self.labelImg.pack(padx=5)

        if self.label1.cget("text") == "Select the area to analyze":
            self.label1.pack()

        if data:
            enunciado = data["Text"]
            codigo = data["QR"]
            if codigo is None:
                self.label1.config(text="Text found")
            else:
                self.label1.config(text="QR code detected")

            # --- Crear un frame contenedor para Text + Scrollbar ---
            # self.frame_text = tk.Frame(self.root)
            self.author.pack(side=tk.BOTTOM, anchor='e')
            self.frame_text.pack(fill=tk.BOTH, expand=True)  # Se expandirá en la parte superior

            # Mover Text y Scrollbar al frame (no a root)
            self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            self.text_box.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

            self.text_box.config(state=tk.NORMAL)
            self.text_box.delete("1.0", tk.END)
            if enunciado.startswith(("http://", "https://", "www.")):
                self.text_box.insert(tk.END, enunciado, ("link", enunciado))
                self.text_box.tag_config("link", foreground="blue", underline=1)
                self.text_box.tag_bind("link", "<Button-1>", open_link)
                self.text_box.tag_bind("link", "<Enter>", lambda e: self.text_box.config(cursor="hand2"))
                self.text_box.tag_bind("link", "<Leave>", lambda e: self.text_box.config(cursor=""))
            else:
                self.text_box.insert(tk.END, enunciado)
            self.text_box.config(state=tk.DISABLED)

            self.root.geometry(str(window_width)+"x"+str(window_height))
            self.root.minsize(window_width, window_height)

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    MainGUI().run()
