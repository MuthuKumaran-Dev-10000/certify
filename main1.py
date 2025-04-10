import os
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import tkinter as tk
from tkinter import filedialog, simpledialog
from tkinter import messagebox

positions = []
clicked = 0

def on_click(event, canvas, image_on_canvas, header_labels):
    global clicked, positions
    if clicked < len(header_labels):
        x, y = event.x, event.y
        positions.append((x, y))
        canvas.create_oval(x-5, y-5, x+5, y+5, outline="red", width=2)
        canvas.create_text(x+40, y, text=header_labels[clicked], fill="blue", anchor="w")
        clicked += 1
        if clicked == len(header_labels):
            canvas.quit()

def get_positions_gui(image_path, headers):
    global positions, clicked
    clicked = 0
    positions = []

    window = tk.Tk()
    window.title("Click positions for each column")

    # Load and display image
    img = Image.open(image_path)
    tk_img = tk.PhotoImage(file=image_path)
    canvas = tk.Canvas(window, width=img.width, height=img.height)
    canvas.pack()
    canvas_image = canvas.create_image(0, 0, anchor=tk.NW, image=tk_img)

    # Instruction
    label = tk.Label(window, text=f"Click on the image for each of the following headers in order:\n{headers}")
    label.pack()

    # Bind click
    canvas.bind("<Button-1>", lambda e: on_click(e, canvas, canvas_image, headers))

    window.mainloop()
    window.destroy()

    return positions

def get_font_sizes(headers):
    root = tk.Tk()
    root.withdraw()

    font_sizes = []
    for col in headers:
        size = simpledialog.askinteger("Font Size", f"Enter font size for '{col}':", initialvalue=30)
        font_sizes.append(size)
    return font_sizes

def certify(cert_image_path, excel_path, n, name_format_index=None):
    df = pd.read_excel(excel_path)
    cert_template = Image.open(cert_image_path)

    headers = df.columns[:n].tolist()
    positions_selected = get_positions_gui(cert_image_path, headers)
    font_sizes = get_font_sizes(headers)

    output_folder = "certificates"
    os.makedirs(output_folder, exist_ok=True)

    for index, row in df.iterrows():
        cert = cert_template.copy()
        draw = ImageDraw.Draw(cert)

        for i in range(n):
            value = str(row[headers[i]])
            position = positions_selected[i]
            try:
                font = ImageFont.truetype("arial.ttf", font_sizes[i])
            except IOError:
                font = ImageFont.load_default()
            draw.text(position, value, font=font, fill="black")

        file_name = (
            str(row[df.columns[name_format_index]]) if name_format_index is not None else str(row[df.columns[0]])
        ) + ".png"
        cert.save(os.path.join(output_folder, file_name))

    messagebox.showinfo("Done", f"{len(df)} certificates created in '{output_folder}' folder.")

# Example usage:
certify("certificate.png", "data.xlsx", 4, name_format_index=0)
