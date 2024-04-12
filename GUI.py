from PIL import Image, ImageTk, ImageOps
import os
import win32clipboard
import win32com.client
from io import BytesIO
import win32gui
import win32con
import tkinter as tk
import re
import sys, json, requests, time
from tkinter import font

class popup_GUI:
    def __init__(self, master, folder_path):
        self.saved_parameters_path = 'user_input_data.json'
        self.master = master
        self.rows = 3
        self.columns = 3
        self.num_img = self.rows * self.columns
        self.canvas_height = 360
        self.folder_path = folder_path
        self.img_size = self.canvas_height // self.rows
        self.PS_window_title = None
        self.window_titles = []
        self.master.title("LACE - by Kyle")
        self.master.attributes("-topmost", False)
        self.user_input_data = None  # store the parameters of GUI as a dictionary
        self.lora_models = self.get_lora_model('./models/loras')    # Get the list of LoRA models

        self.custom_font = font.Font(family="Helvetica", size=9)
        self.transparent_images = []
        self.counter = 0
        # Set up the grid configuration for master
        self.master.rowconfigure(0, weight=1)
        for i in range(self.columns):
            self.master.columnconfigure(i, weight=1)

        # Canvas setup with grid
        self.canvas = tk.Canvas(master, width=self.img_size * self.columns, height=self.canvas_height)
        self.canvas.grid(row=0, column=0, columnspan=2, sticky='nsew')
        self.canvas.config(bg='gray75')

        #1 Sliders setup with grid
        self.sampling_step_slider = tk.Scale(self.master, from_=1, to=20, orient='horizontal', label='Sampling Steps')
        self.sampling_step_slider.set(20)
        self.sampling_step_slider.grid(row=1, column=0, sticky='ew', padx=5, pady=5)

        #2 LACE_step_slider Scale
        self.LACE_step_slider = tk.Scale(self.master, from_=10, to=20, orient='horizontal', label='Preview Step')
        self.LACE_step_slider.set(12)
        self.LACE_step_slider.grid(row=1, column=1, sticky='ew', padx=5, pady=5)

        #3 denoise Scale
        self.denoise_slider = tk.Scale(self.master, from_=0.5, to=1, resolution=0.01, orient='horizontal', label='Denoise Level')
        self.denoise_slider.set(1)
        self.denoise_slider.grid(row=2, column=0, sticky='ew', padx=5, pady=0)

        #4 noise_scale Scale
        self.noise_scale_slider = tk.Scale(self.master, from_=0.01, to=0.5, resolution=0.01, orient='horizontal', label='Diversity Level')
        self.noise_scale_slider.set(0.25)
        self.noise_scale_slider.grid(row=2, column=1, sticky='ew', padx=5, pady=0)

        #5 Control Net strength Slider    
        self.controlnet_slider = tk.Scale(self.master, from_=0, to=1, resolution=0.01, orient='horizontal', label='Output Influence')
        self.controlnet_slider.set(0.1)
        self.controlnet_slider.grid(row=3, column=0, sticky='nesw', padx=5, pady=10, columnspan=2)

        #6 prompt_positive Entry
        self.prompt_positive_entry = tk.Entry(self.master)
        self.prompt_positive_label = tk.Label(self.master, text='Prompt Positive')
        self.prompt_positive_entry = tk.Text(self.master, height=3, width=48, wrap=tk.WORD, font=self.custom_font)
        self.prompt_positive_label.grid(row=4, column=0, sticky='w', padx=5, pady=0, columnspan=2)
        self.prompt_positive_entry.grid(row=5, column=0, sticky='new', padx=5, pady=5, columnspan=2, ipady=3)
        self.prompt_positive_entry.insert(tk.END, "MATISSEE-ART")

        #7 prompt_negative Entry
        self.prompt_negative_entry = tk.Entry(self.master)
        self.prompt_negative_label = tk.Label(self.master, text='Prompt Negative')
        self.prompt_negative_entry = tk.Text(self.master, height=3, width=48, wrap=tk.WORD, font=self.custom_font)
        self.prompt_negative_label.grid(row=6, column=0, sticky='w', padx=5, pady=0, columnspan=2)
        self.prompt_negative_entry.grid(row=7, column=0, sticky='new', padx=5, pady=5, columnspan=2, ipady=3)
        self.prompt_negative_entry.insert(tk.END, "bad composition, blurry image, low resolution")

        #8 noise_type OptionMenu
        self.noise_type_var = tk.StringVar(self.master)
        self.diversity_label = tk.Label(self.master, text='Output Diversity')
        self.diversity_label.grid(row=8, column=0, sticky='w', padx=5, pady=0)
        self.noise_type_var.set('Gaussian')  # default value
        self.noise_type_menu = tk.OptionMenu(self.master, self.noise_type_var, 'Gaussian', 'Uniform', 'Exponential')
        self.noise_type_menu.grid(row=9, column=0, sticky='ew', padx=5, pady=0)

        #9 creative_mode OptionMenu
        self.creative_mode_var = tk.StringVar(self.master)
        self.creative_mode_label = tk.Label(self.master, text='Creative Mode (Reverse CADS)')
        self.creative_mode_label.grid(row=8, column=1, sticky='w', padx=5, pady=0)
        self.creative_mode_var.set('Normal')  # default value
        self.creative_mode_menu = tk.OptionMenu(self.master, self.creative_mode_var, 'Normal', 'Radical')
        self.creative_mode_menu.grid(row=9, column=1, sticky='ew', padx=5, pady=0)
        
        #10 num of output OptionMenu
        self.num_output_var = tk.StringVar(self.master)
        self.num_output_label = tk.Label(self.master, text='Number of Output')
        self.num_output_label.grid(row=10, column=0, sticky='w', padx=5, pady=0)
        self.num_output_var.set('4')
        self.num_output_menu = tk.OptionMenu(self.master, self.num_output_var, '1', '4', '9')
        self.num_output_menu.grid(row=11, column=0, sticky='ew', padx=5, pady=0)

        #11 LoRA OptionMenu     
        self.lora_var = tk.StringVar(self.master)  # Changed variable name to lora_var
        self.lora_label = tk.Label(self.master, text='LoRA Model')  # Changed variable name to lora_label
        self.lora_label.grid(row=10, column=1, sticky='w', padx=5, pady=0)
        if self.lora_models:
            default_lora_model = self.lora_models[-2]
        else:
            default_lora_model = 'No models found'
        self.lora_var.set(default_lora_model)  # default value
        self.lora_menu = tk.OptionMenu(self.master, self.lora_var, *self.lora_models)
        self.lora_menu.grid(row=11, column=1, sticky='ew', padx=5, pady=0)

        #12 Divider
        self.divider = tk.Frame(self.master, height=10, bd=0, relief=tk.SUNKEN)
        self.divider.grid(row=12, column=0, sticky='ew', padx=5, pady=0, columnspan=2)

        # Submit Button
        # self.submit_button = tk.Button(self.master, text="Generate", command=self.submit)
        # self.submit_button.grid(row=11, column=0, columnspan=2, sticky='ew', padx=5, pady=15, ipady=6)


        # Button setup with grid
        self.toggle_button = tk.Button(self.master, text="📌", command=self.borderless)
        self.toggle_button.grid(row=0, column=self.columns - 2, sticky='ne', padx=5, pady=5, ipady=3, ipadx=3)

        # Initialize additional attributes and bindings
        self.is_borderless = False
        self.image_objects = []
        self.last_batch = []

        self.canvas.bind("<Button-1>", self.on_canvas_click)
        self.canvas.bind("<B1-Motion>", self.on_motion)
        
        self.display_queue = []
        self.update_image()

    def restart_program(self):
        python = sys.executable
        os.execl(python, python, * sys.argv)
    def send_prompt(self):
        url_history = "http://localhost:8188/history"
        url_prompt = "http://localhost:8188/prompt"
        
        history_data_dict = requests.get(url_history).json()
        first_key = next(iter(history_data_dict))
        first_res = history_data_dict[first_key]
        first_promt = first_res['prompt']
        
        first_json_prompt = first_promt[2]
        first_json_extra = first_promt[3]

        final_json = {'prompt': first_json_prompt, 'extra_data':first_json_extra}
        # s = final_json['extra_data']['extra_pnginfo']['workflow']['180']['inputs']['seed']

        # print(f"first_json_extra keys\n{first_json_extra['extra_pnginfo']['workflow'].keys()}")
        # print(f'seed: {s}')

        response = requests.post(url_prompt, json=final_json)
        
        if response.status_code == 200:
            print("Prompt queued successfully")
        else:
            print("Failed to queue prompt. Status code:", response.status_code)
                    

    def submit(self):
        # Collect all the values from the widgets and store them
        self.user_input_data = {
            'num_output': int(self.num_output_var.get()),
            'noise_scale': self.noise_scale_slider.get(),
            'noise_type': self.noise_type_var.get(),
            'reverse_CADS': self.creative_mode_var.get(),
            'denoise': self.denoise_slider.get(),
            'prompt_positive': self.prompt_positive_entry.get("1.0", tk.END).strip(),
            'prompt_negative': self.prompt_negative_entry.get("1.0", tk.END).strip(),
            'lora_model': self.lora_var.get(),
            'visualized_steps': self.LACE_step_slider.get(),
            'sampling_steps': self.sampling_step_slider.get(),
            'strength': self.controlnet_slider.get(),
        }
        with open(self.saved_parameters_path, 'w') as f:
            json.dump(self.user_input_data, f, indent=4)
        
        # self.send_prompt()
        
    def get_lora_model(self, path):
        try:
            return ["None"] + [model for model in os.listdir(path)]
        except Exception as e:
            print(f"Error: {e}")
            return ["None"]

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

    def clear_folder(self, folder_path):
        for filename in os.listdir(folder_path):
            file_path = os.path.join(folder_path, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
            except Exception as e:
                print(f'Failed to delete {file_path}. Reason: {e}')

    def load_images(self, folder_path):
        if not os.path.exists(folder_path):
            return []
        image_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith(('png', 'jpg', 'jpeg', 'gif'))]
        image_files.sort(key=lambda x: os.path.getctime(x), reverse=True)
        num_output = int(self.num_output_var.get())  # Convert to integer
        latest_images = image_files[:num_output]
        return latest_images

    def borderless(self):
        if self.is_borderless:
            self.master.overrideredirect(False)
            self.is_borderless = False
            self.master.attributes("-topmost", False)
            # self.Notif("Borderless Disabled")
        else:
            self.master.attributes("-topmost", True)
            self.master.overrideredirect(True)
            self.is_borderless = True
            # self.Notif("Borderless Enabled")
        
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
        latest_images = self.load_images(self.folder_path)

        if latest_images != self.last_batch:
            self.last_batch = latest_images
            new_images_count = int(self.num_output_var.get())

            # Initialize or update transparent images if necessary
            if not self.transparent_images or len(self.transparent_images) != len(self.display_queue):
                self.transparent_images = [Image.new("RGBA", (self.img_size, self.img_size), (0, 0, 0, 0)) for _ in range(max(len(self.display_queue), new_images_count))]

            # Overlay new images and initiate fade-in animations
            for img_index, img_path in enumerate(latest_images[:new_images_count]):
                render_idx = (self.counter + img_index) % (self.rows * self.columns)
                pil_image = Image.open(img_path).resize((self.img_size, self.img_size), Image.Resampling.LANCZOS)
                border_size = 3  # Size of the border
                pil_image = ImageOps.expand(pil_image, border=border_size, fill='white')
                x = (render_idx % self.columns) * self.img_size + self.img_size / 2
                y = (render_idx // self.columns) * self.img_size + self.img_size / 2
                delay = img_index * 500  # for example, 1000ms per image
                self.master.after(delay, lambda x=x, y=y, pil_image=pil_image, img_path=img_path: self.fade_in_image(x, y, 0, 10, pil_image, img_path))
            self.counter += new_images_count

        self.submit()
        self.master.after(1000, self.update_image)

    def fade_in_image(self, x, y, step, num_steps, pil_image, img_path):
        if step <= num_steps:
            alpha = step / num_steps
            faded_image = pil_image.copy()
            faded_image.putalpha(int(alpha * 255))
            tk_faded_image = ImageTk.PhotoImage(faded_image)
            if step == 0:
                image_id = self.canvas.create_image(x, y, image=tk_faded_image)
                # Store the image_id in the object's property for future reference
                self.image_objects.append((image_id, tk_faded_image, pil_image, img_path))
            else:
                # Update the existing image instead of creating a new one
                image_id = self.image_objects[-1][0]  # Get the last image id stored
                self.canvas.itemconfig(image_id, image=tk_faded_image)
                # Update the stored image to prevent garbage collection
                self.image_objects[-1] = (image_id, tk_faded_image, pil_image, img_path)

            if step < num_steps:
                # Increase the delay to slow down the animation
                self.master.after(20, lambda: self.fade_in_image(x, y, step + 1, num_steps, pil_image, img_path))


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
        print(self.image_objects)
        for img_id, _, _, img_path in self.image_objects:
            if img_id == clicked_img:
                print(f"Clicked on image with ID: {img_id}")
                print(f"Image path: {img_path}")
                self.copy_image_to_clipboard(img_path)
                break

# Usage
if __name__ == "__main__":
    root = tk.Tk()
    width = 360
    height = 855
    x_position = 1200
    y_position = 50
    root.geometry(f'{width}x{height}+{x_position}+{y_position}')
    app = popup_GUI(root, './temp')
    root.mainloop()