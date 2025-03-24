import tkinter as tk
from tkinter import filedialog, messagebox
import os
import shutil
import base64
import requests
from pdf2image import convert_from_path
import pandas as pd
from PIL import Image, ImageEnhance, ImageFilter

# OpenAI API Key
api_key = "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"


def start_loading(text, item, total):
    """Starts the loading count."""
    load = (item / total)*100
    loading_label.config(text=f"{text}: {round(load, 1)}%")
    root.update()  # Force update of the GUI

def preprocess_image_advanced(image_path, contrast=1.5, sharpness=1.5):
    """Preprocesses low-quality images using OpenCV."""
    img = Image.open(image_path)
    img = img.convert('L')  # Convert to grayscale
    img = ImageEnhance.Contrast(img).enhance(1.5)
    img = ImageEnhance.Sharpness(img).enhance(0.6)
    img = img.filter(ImageFilter.MedianFilter())
    img.save(image_path)

def pdf_to_jpg(pdf_path, output_folder=None):

    if output_folder is None:
        output_folder = os.path.dirname(pdf_path)
    data_type = os.path.splitext(pdf_path)[1]
    if data_type=='.pdf':
        images = convert_from_path(pdf_path)
        jpg_paths = []
        total  = len(images)
        iter  = 1
        for i, image in enumerate(images):
            jpg_path = os.path.join(output_folder, f'page_{i + 1}.jpg')
            image.save(jpg_path, 'JPEG')
            jpg_paths.append(jpg_path)
            preprocess_image_advanced(jpg_path)
            start_loading("Analizando PDF", iter, total)
            iter +=1
        return jpg_paths
    else:
        return pdf_path


def encode_image(image_path):
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode('utf-8')


def ocr_gpt(images):
    compania = []
    rnc = []
    ncf = []
    fecha_emision = []
    sub_itbis = []
    sub_valor = []
    jpeg = []
    iter = 1
    total = len(images)
    if isinstance(images, list)==False:
        jpeg.append(images)
        images = jpeg
    for i in images:

        # Getting the base64 string
        base64_image = encode_image(i)

        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}"
        }

        payload = {
            "model": "gpt-4o-mini",
            "messages": [
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "text",
                            "text": """De esta factura extraer la 
                            empresa que realiza la factura, 
                            extraer el RNC del emisor o proveedor (el RNC con un minimo de 9 caracteres y esta debajo del nombre de la empresa que emite la factura), 
                            extraer el NCF (el NCF inicia con las letras B o E, con un total de 11 caracteres.), 
                            la fecha de emision(sin hora, en formato DD/MM/YYYY), 
                            el subtotal del itbis y 
                            el subtotal del valo. Solo devolver los datos solicitados, ignorar simbolos de moneda y no usar asteriscos"""
                        },
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/jpeg;base64,{base64_image}"
                            }
                        }
                    ]
                }
            ],
            "max_tokens": 1024,
            "n": 1,
            "temperature": 0.25
        }

        response = requests.post("https://api.openai.com/v1/chat/completions", headers=headers, json=payload)
        text_json = response.json()['choices'][0]['message']['content']
        output = [text_json.split(': ')]
        compania.append(output[0][1].replace("RNC", "").replace("DEL EMISOR O PROVEEDOR", "").replace("DEL EMISOR", "").strip())
        rnc.append(output[0][2].replace("NCF", "").strip())
        ncf.append(output[0][3].replace("Fecha de emisión", "").replace("Fecha de Emisión", "").strip())
        fecha_emision.append(output[0][4].replace("Subtotal ITBIS", "").replace("Subtotal del ITBIS", "").strip())
        sub_itbis.append(output[0][5].replace("Subtotal Valor", "").replace("Subtotal valor", "").replace("Subtotal del valor", "").replace("Subtotal", "").strip())
        sub_valor.append(output[0][6].strip())
        start_loading("Analizando Facturas", iter, total)
        iter += 1
        os.remove(i)


    data = {
        'compania': compania,
        'rnc': rnc,
        'ncf': ncf,
        'fecha_emision': fecha_emision,
        'sub_itbis': sub_itbis,
        'sub_valor': sub_valor
    }

    df = pd.DataFrame(data)
    df.compania=df.compania.str.replace('-','').str.strip().str.upper()
    df.rnc=df.rnc.str.replace('-', '').str.strip()
    df.ncf=df.ncf.str.replace('-', '').str.strip()
    df.fecha_emision = df.fecha_emision.str.replace('-', '').str.strip()
    df.sub_itbis=df.sub_itbis.str.replace('-', '' ).str.strip()
    df.sub_valor=df.sub_valor.str.replace('-', '' ).str.strip()
    return df


def get_and_print_filename():
    """Opens a file dialog, gets the selected filename, and displays it."""
    filepath = filedialog.askopenfilename()  # Open file dialog

    if filepath:
        filename = os.path.basename(filepath)  # Get filename
        directory = os.path.dirname(filepath)  # Get the directory the file is in.
        folder_name = os.path.splitext(filename)[0]  # removes extension
        new_folder_path = os.path.join(directory, folder_name)
        new_file_path = os.path.join(new_folder_path, filename)

        try:
            loading_label.config(text="Iniciando Proceso")
            root.update()
            os.makedirs(new_folder_path, exist_ok=True)  # creates folder, no error if it exists.
            shutil.copy(filepath, new_file_path)
            imgs = pdf_to_jpg(new_file_path)
            df_ocr = ocr_gpt(imgs)
            df_ocr.to_excel(new_folder_path+'/Resultados_'+folder_name+'.xlsx')
            messagebox.showinfo("Proceso Completado", f"Proceso completado con éxito para el archivo '{folder_name}'!")
        except OSError as e:
            messagebox.showerror("Error", f"Error en el proceso, cancele e intente de nuevo: {e}")
    else:
        messagebox.showinfo("Selección de Archivo", "Ningún acrchivo seleccionado.")


def close_window():
    """Closes the main window."""
    root.destroy()  # Destroys the main window and all its widgets

# Create the main window
root = tk.Tk()
root.geometry("400x200")  # Width x Height (pixels)
root.title("Análisis de Facturas")

# Create a button to trigger the file selection
select_button = tk.Button(root, text="Seleccione PDF o Imágen", command= get_and_print_filename)
select_button.pack(pady=10, expand=True)  # Add some padding
loading_label = tk.Label(root, text="")
loading_label.pack(pady=20)
close_button = tk.Button(root, text="Cerrar", command=close_window)
close_button.pack(pady=10, expand=True)  # Add some padding

# Run the application
root.mainloop()