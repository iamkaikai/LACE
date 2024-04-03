import tkinter as tk
from PIL import Image, ImageTk
import os
import win32clipboard
import win32com.client
from io import BytesIO
import win32gui
import win32con
import re

class ImageSlideshow:

    def __init__(self, master, folder_path, rows=3, columns=3):
        self.master = master
        self.folder_path = folder_path
        self.rows = rows
        self.columns = columns
        self.img_size = 200
        self.num_img = 9
        self.PS_window_title = None
        self.master.title("Image Slideshow")
        self.master.attributes("-topmost", True)
        self.canvas = tk.Canvas(master, width=self.img_size * columns, height=self.img_size * rows)
        self.canvas.pack()
        # self.enumerate_windows()
        self.image_objects = []
        self.last_batch = []

        self.canvas.bind("<Button-1>", self.on_canvas_click)
        self.images = self.load_images(self.folder_path)
        self.update_image()

    def load_images(self, folder_path):
        image_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith(('png', 'jpg', 'jpeg', 'gif'))]
        image_files.sort(key=lambda x: os.path.getmtime(x), reverse=True)       # Sort the files by modification time in descending order
        latest_images = image_files[:self.num_img]                                         # Keep only the latest nine images
        return latest_images, [Image.open(img) for img in latest_images]
    
    def bring_to_front(self, window_title, doc):
        hwnd = win32gui.FindWindow(None, window_title)
        if hwnd:
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)  # Unminimize
            win32gui.SetForegroundWindow(hwnd)  # Bring to front
            doc.Paste()



    def enumerate_windows(self):
        def callback(hwnd, extra):
            if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd) != "":
                print(f"Window Handle: {hwnd}, Window Title: {win32gui.GetWindowText(hwnd)}")
        win32gui.EnumWindows(callback, None)


    def find_photoshop_window(self):
        titles = []

        def callback(hwnd, titles):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if re.match(r".*@.*\*.*", title):
                    titles.append((hwnd, title))

        win32gui.EnumWindows(callback, titles)
        self.PS_window_title = titles[0][1]
        
    

    def update_image(self):
        
        try:
            latest_images, image_objects = self.load_images(self.folder_path)  # Reload the latest images

            if not latest_images:
                self.canvas.create_text(
                    self.canvas.winfo_width() / 2, self.canvas.winfo_height() / 2,
                    text="Loading...", font=('Helvetica', 12), fill="gray")

            if latest_images != self.last_batch:
                self.last_batch = latest_images
                self.images = image_objects
                self.canvas.delete("all")  # Clear the canvas

                for i in range(self.rows):
                    for j in range(self.columns):
                        img_index = (i * self.columns + j) % len(self.images)
                        pil_image = self.images[img_index].resize((self.img_size, self.img_size), Image.Resampling.LANCZOS)
                        tk_image = ImageTk.PhotoImage(pil_image)
                        image_id = self.canvas.create_image(j * self.img_size + self.img_size/2, i * self.img_size + self.img_size/2, image=tk_image, tags=f"image{img_index}")
                        self.image_objects.append((image_id, tk_image, latest_images[img_index]))
                    
            # Schedule the next update; adjust the interval as needed
            self.master.after(1000, self.update_image)

        except Exception as e:
            print(f"Error: {e}")
            self.master.after(500, self.update_image)



    def copy_image_to_clipboard(self, img_path):
        print(f"Copying to clipboard: {img_path}")  # Debug: Print the image path being processed
        image = Image.open(img_path)
        if image.mode != 'RGB':
            image = image.convert('RGB')
        
        output = BytesIO()
        image.save(output, 'BMP')
        data = output.getvalue()[14:]
        output.close()

        win32clipboard.OpenClipboard()  # Open the clipboard to enable us to change its contents
        win32clipboard.EmptyClipboard()  # Clear the current contents of the clipboard
        win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)  # Set the clipboard data to our image
        win32clipboard.CloseClipboard()  # Close the clipboard to release it for other applications to use

        # Create a new Photoshop instance
        ps = win32com.client.Dispatch("Photoshop.Application")
        doc = ps.Documents[0]
        self.find_photoshop_window()
        self.bring_to_front(self.PS_window_title, doc)

    def on_canvas_click(self, event):
        clicked_img = self.canvas.find_closest(event.x, event.y)[0]
        for img_id, _, img_path in self.image_objects:
            if img_id == clicked_img:
                print(f"Clicked on image with ID: {img_id}")
                print(f"Image path: {img_path}")
                self.copy_image_to_clipboard(img_path)
                break

# Usage

root = tk.Tk()
app = ImageSlideshow(root, './temp')
root.mainloop()



