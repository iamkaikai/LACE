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

    def __init__(self, master, folder_path, rows=2, columns=2):
        self.master = master
        self.folder_path = folder_path
        self.rows = rows
        self.columns = columns
        self.img_size = 256
        self.num_img = rows * columns
        self.PS_window_title = None
        self.window_titles = []
        self.master.title("LACE - by Kyle")
        self.master.attributes("-topmost", True)
        self.canvas = tk.Canvas(master, width=self.img_size * columns, height=self.img_size * rows)
        self.canvas.pack()
        
        self.is_borderless = False
        self.image_objects = []
        self.last_batch = []
        self.toggle_button = tk.Button(self.master, text="📌", command=self.borderless)
        self.toggle_button.pack()
        self.toggle_button.place( x = columns*self.img_size-25, y=4)
       
        self.start_x = self.master.winfo_x()
        self.start_y = self.master.winfo_y()

        self.canvas.bind("<Button-1>", self.on_canvas_click)
        self.canvas.bind("<B1-Motion>", self.on_motion)
        
        self.images = self.load_images(self.folder_path)
        self.update_image()


    def on_motion(self, event):
        # Get the absolute screen position of the mouse
        current_mouse_x = self.master.winfo_pointerx()
        current_mouse_y = self.master.winfo_pointery()

        # Calculate how much the mouse has moved
        deltax = current_mouse_x - self.start_x
        deltay = current_mouse_y - self.start_y

        # Update the start positions for the next motion event
        self.start_x = current_mouse_x
        self.start_y = current_mouse_y

        # Calculate the new window position
        new_x = self.master.winfo_x() + deltax
        new_y = self.master.winfo_y() + deltay

        # Move the window
        self.master.geometry(f"+{new_x}+{new_y}")

    def load_images(self, folder_path):
        image_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith(('png', 'jpg', 'jpeg', 'gif'))]
        image_files.sort(key=lambda x: os.path.getmtime(x), reverse=True)       # Sort the files by modification time in descending order
        latest_images = image_files[:self.num_img]                                         # Keep only the latest nine images
        return latest_images, [Image.open(img) for img in latest_images]

    def borderless(self):
        if self.is_borderless:
            self.master.overrideredirect(False)
            self.is_borderless = False
            self.master.attributes("-topmost", False)
            self.Notif("Borderless Disabled")
        else:
            self.master.attributes("-topmost", True)
            self.master.overrideredirect(True)
            self.is_borderless = True
            self.Notif("Borderless Enabled")
        
    def enumerate_windows(self):
        self.window_titles=[]
        def callback(hwnd, extra):
            if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd) != "":
                self.window_titles.append({win32gui.GetWindowText(hwnd)})
        win32gui.EnumWindows(callback, None)

    def find_photoshop_window(self):
        for title_set in self.window_titles:
            if not title_set:  # Skip empty sets
                continue
            title = next(iter(title_set))  # Get the first item from the set
            if re.match(r".*@.*\*.*", title):
                self.PS_window_title = title
                break

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
                
                for _, tk_image, _ in self.image_objects:
                    self.canvas.delete(tk_image)  # Remove from canvas
                self.image_objects.clear() 

                for i in range(self.rows):
                    for j in range(self.columns):
                        img_index = (i * self.columns + j) % len(self.images)
                        with Image.open(latest_images[img_index]) as pil_image:
                            pil_image = pil_image.resize((self.img_size, self.img_size), Image.Resampling.LANCZOS)
                            tk_image = ImageTk.PhotoImage(pil_image)
                            x = j * self.img_size + self.img_size / 2
                            y = i * self.img_size + self.img_size / 2
                            self.canvas.create_rectangle(x - self.img_size / 2, y - self.img_size / 2, x + self.img_size / 2, y + self.img_size / 2, outline="white", width=5)
                            image_id = self.canvas.create_image(x, y, image=tk_image, tags=f"image{img_index}")
                            self.image_objects.append((image_id, tk_image, latest_images[img_index]))

                    
            # Schedule the next update; adjust the interval as needed
            self.master.after(1000, self.update_image)

        except Exception as e:
            print(f"Error: {e}")
            self.master.after(500, self.update_image)

    def bring_to_front(self, window_title, doc):
        hwnd = win32gui.FindWindow(None, window_title)
        if hwnd:
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)  # Unminimize
            win32gui.SetForegroundWindow(hwnd)  # Bring to front
            doc.Paste()

    def activate_document(self, ps, target_title):
        for i in range(ps.Documents.Count):
            doc = ps.Documents.Item(i + 1)
            if doc.Name == target_title:
                doc.Activate()
                break

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
        self.enumerate_windows()
        self.find_photoshop_window()
        self.bring_to_front(self.PS_window_title, doc)

    def on_canvas_click(self, event):
        self.start_x = self.master.winfo_pointerx()
        self.start_y = self.master.winfo_pointery()
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



