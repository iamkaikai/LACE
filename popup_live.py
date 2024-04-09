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
        self.img_size = 180
        self.num_img = rows * columns
        self.PS_window_title = None
        self.window_titles = []
        self.master.title("LACE - by Kyle")
        self.master.attributes("-topmost", False)

        # Set up the grid configuration for master
        self.master.rowconfigure(0, weight=1)
        for i in range(self.columns):
            self.master.columnconfigure(i, weight=1)

        # Canvas setup with grid
        self.canvas = tk.Canvas(master, width=self.img_size * columns, height=self.img_size * rows)
        self.canvas.grid(row=0, column=0, columnspan=self.columns, sticky='nsew')
        self.canvas.config(bg='gray75')

        # Sliders setup with grid
        self.sampling_step_slider = tk.Scale(self.master, from_=1, to=20, orient='horizontal', label='Sampling Steps')
        self.sampling_step_slider.set(16)
        self.sampling_step_slider.grid(row=1, column=0, sticky='ew', padx=5, pady=5)

        self.LACE_step_slider = tk.Scale(self.master, from_=1, to=20, orient='horizontal', label='Visualized Steps')
        self.LACE_step_slider.set(12)
        self.LACE_step_slider.grid(row=1, column=1, sticky='ew', padx=5, pady=5)

        # num_images Scale
        self.num_images_slider = tk.Scale(self.master, from_=1, to=9, tickinterval=1, orient='horizontal', label='Number of Images')
        self.num_images_slider.grid(row=2, column=0, sticky='ew', padx=5, pady=5)

        # noise_scale Scale
        self.noise_scale_slider = tk.Scale(self.master, from_=0, to=0.5, resolution=0.01, orient='horizontal', label='Noise Scale')
        self.noise_scale_slider.set(0.2)
        self.noise_scale_slider.grid(row=2, column=1, sticky='new', padx=5, pady=5)

        # prompt_positive Entry
        self.prompt_positive_entry = tk.Entry(self.master)
        self.prompt_positive_label = tk.Label(self.master, text='Prompt Positive')
        self.prompt_positive_entry = tk.Text(self.master, height=3, width=48, wrap=tk.WORD)
        self.prompt_positive_label.grid(row=3, column=0, sticky='w', padx=5, pady=0, columnspan=2)
        self.prompt_positive_entry.grid(row=4, column=0, sticky='new', padx=5, pady=5, columnspan=2, ipady=3)

        # prompt_negative Entry
        self.prompt_negative_entry = tk.Entry(self.master)
        self.prompt_negative_label = tk.Label(self.master, text='Prompt Negative')
        self.prompt_negative_entry = tk.Text(self.master, height=3, width=48, wrap=tk.WORD)
        self.prompt_negative_label.grid(row=5, column=0, sticky='w', padx=5, pady=0, columnspan=2)
        self.prompt_negative_entry.grid(row=6, column=0, sticky='new', padx=5, pady=5, columnspan=2, ipady=3)

        # noise_type OptionMenu
        self.noise_type_var = tk.StringVar(self.master)
        self.diversity_label = tk.Label(self.master, text='Output Diversity')
        self.diversity_label.grid(row=7, column=0, sticky='w', padx=5, pady=0)
        self.noise_type_var.set('Gaussian')  # default value
        self.noise_type_menu = tk.OptionMenu(self.master, self.noise_type_var, 'Gaussian', 'Uniform', 'Exponential')
        self.noise_type_menu.grid(row=8, column=0, sticky='ew', padx=5, pady=0)

        # creative_mode OptionMenu
        self.creative_mode_var = tk.StringVar(self.master)
        self.creative_mode_label = tk.Label(self.master, text='Creative Mode')
        self.creative_mode_label.grid(row=7, column=1, sticky='w', padx=5, pady=0)
        self.creative_mode_var.set('Normal')  # default value
        self.creative_mode_menu = tk.OptionMenu(self.master, self.creative_mode_var, 'Normal', 'Radical')
        self.creative_mode_menu.grid(row=8, column=1, sticky='ew', padx=5, pady=0)
        
        # seed OptionMenu
        self.seed_var = tk.StringVar(self.master)
        self.seed_label = tk.Label(self.master, text='Sampling Seed')
        self.seed_label.grid(row=9, column=0, sticky='w', padx=5, pady=0)
        self.seed_var.set('Incremental')  # default value
        self.seed_menu = tk.OptionMenu(self.master, self.seed_var, 'Fixed', 'Incremental', 'Randomized')
        self.seed_menu.grid(row=10, column=0, sticky='ew', padx=5, pady=0)

        # LoRA OptionMenu
        self.lora_var = tk.StringVar(self.master)  # Changed variable name to lora_var
        self.lora_label = tk.Label(self.master, text='LoRA Model')  # Changed variable name to lora_label
        self.lora_label.grid(row=9, column=1, sticky='w', padx=5, pady=0)
        self.lora_var.set('Cubism')  # default value
        self.lora_menu = tk.OptionMenu(self.master, self.lora_var, 'Cubism', 'Impressionism', 'Surrealism')  # Changed options for demonstration
        self.lora_menu.grid(row=10, column=1, sticky='ew', padx=5, pady=0)

        # Submit Button
        self.submit_button = tk.Button(self.master, text="Generate", command=self.submit)
        self.submit_button.grid(row=11, column=0, columnspan=2, sticky='ew', padx=5, pady=15, ipady=6)


        # Button setup with grid
        self.toggle_button = tk.Button(self.master, text="📌", command=self.borderless)
        self.toggle_button.grid(row=11, column=self.columns - 1, sticky='ne', padx=5, pady=15, ipady=6, ipadx=6)

        # Initialize additional attributes and bindings
        self.is_borderless = False
        self.image_objects = []
        self.last_batch = []

        self.canvas.bind("<Button-1>", self.on_canvas_click)
        self.canvas.bind("<B1-Motion>", self.on_motion)
        
        self.images = self.load_images(self.folder_path)
        self.update_image()

    def submit(self):
        # Here you will collect all the values from the widgets and process them
        print("Seed:", self.seed_var.get())
        print("Noise Scale:", self.noise_scale_slider.get())
        print("Noise Type:", self.noise_type_var.get())
        print("Creative Mode:", self.creative_mode_var.get())
        print("Number of Images:", self.num_images_slider.get())
        print("Prompt Positive:", self.prompt_positive_entry.get("1.0", tk.END).strip())  # For Text widget
        print("Prompt Negative:", self.prompt_negative_entry.get("1.0", tk.END).strip())  # For Text widget
        print("LoRA Model:", self.lora_var.get())
        print("Visualized Steps:", self.LACE_step_slider.get())
        print("Sampling Steps:", self.sampling_step_slider.get())

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
        if not os.path.exists(folder_path):
            return [],[]
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
                    text="Loading...", font=('Helvetica', 8), fill="gray")

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



