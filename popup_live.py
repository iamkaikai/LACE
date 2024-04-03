import tkinter as tk
from PIL import Image, ImageTk
import os

class ImageSlideshow:
    def __init__(self, master, folder_path, rows=3, columns=3):
        self.master = master
        self.folder_path = folder_path
        self.rows = rows
        self.columns = columns

        self.master.title("Image Slideshow")
        self.master.attributes("-topmost", True)
        self.canvas = tk.Canvas(master, width=256 * columns, height=256 * rows)
        self.canvas.pack()

        self.images = self.load_images(self.folder_path)
        self.update_image()

    def load_images(self, folder_path):
        """Load the latest nine images from the folder and return a list of PIL images."""
        image_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith(('png', 'jpg', 'jpeg', 'gif'))]
        image_files.sort(key=lambda x: os.path.getmtime(x), reverse=True)       # Sort the files by modification time in descending order
        latest_images = image_files[:9]                                         # Keep only the latest nine images
        return [Image.open(img) for img in latest_images]

    def update_image(self):
        self.images = self.load_images(self.folder_path)  # Reload the latest images
        self.image_objects = []  # Reset the list to keep references to PhotoImage objects
        if not self.images:  # No images to display
            return

        self.canvas.delete("all")  # Clear the canvas

        for i in range(self.rows):
            for j in range(self.columns):
                img_index = (i * self.columns + j) % len(self.images)
                pil_image = self.images[img_index].resize((256, 256), Image.Resampling.LANCZOS)
                tk_image = ImageTk.PhotoImage(pil_image)
                self.canvas.create_image(j * 256 + 128, i * 256 + 128, image=tk_image)
                self.image_objects.append(tk_image)  # Keep a reference

        # Schedule the next update; adjust the interval as needed
        self.master.after(1000, self.update_image)  # E.g., refresh every 5 seconds


# Usage
root = tk.Tk()
app = ImageSlideshow(root, './temp')
root.mainloop()
