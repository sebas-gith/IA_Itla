import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import customtkinter
import pandas as pd
from azure.ai.textanalytics import TextAnalyticsClient
from azure.core.credentials import AzureKeyCredential
from PIL import Image, ImageTk
from dotenv import load_dotenv
import os

load_dotenv()

ENDPOINT = os.getenv("ENDPOINT")
KEY = os.getenv("API_KEY")

def authenticate_client():
    return TextAnalyticsClient(endpoint=ENDPOINT, credential=AzureKeyCredential(KEY))

try:
    icono_cargar = customtkinter.CTkImage(Image.open("cargar_icon.png"), size=(20, 20))
    icono_analizar = customtkinter.CTkImage(Image.open("analizar_icon.png"), size=(20, 20))
    icono_guardar = customtkinter.CTkImage(Image.open("guardar_icon.png"), size=(20, 20))
except FileNotFoundError:
    print("Advertencia: No se encontraron los archivos de íconos. Los botones se mostrarán sin íconos.")
    icono_cargar = None
    icono_analizar = None
    icono_guardar = None

df = None
client = authenticate_client()

def cargar_excel():
    global df
    archivo = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if archivo:
        try:
            df = pd.read_excel(archivo)
            messagebox.showinfo("Éxito", f"Archivo cargado correctamente.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer el archivo:\n{e}")
        
        if df is not None:
            label_estado.configure(text=f"Archivo cargado: {archivo.split('/')[-1]}")
            label_estado.configure(text_color="green")
        else:
            label_estado.configure(text="No se ha cargado ningún archivo.")
            label_estado.configure(text_color="red")

def analizar():
    global df
    if df is None:
        messagebox.showwarning("Advertencia", "Primero carga un archivo Excel.")
        return
    
    if "comentario" not in df.columns:
        messagebox.showerror("Error", "El archivo debe tener una columna llamada 'comentario'.")
        return

    comentarios = df["comentario"].astype(str).tolist()

    try:
        sentiment_result = client.analyze_sentiment(comentarios)

        key_phrases_result = client.extract_key_phrases(comentarios)

        sentimientos = []
        afinidades = []
        etiquetas = []

        for sent, keyp in zip(sentiment_result, key_phrases_result):
            if not sent.is_error:
                sentimientos.append("Positivo" if sent.sentiment == "positive" else "Negativo")
                afinidades.append(round(sent.confidence_scores.positive * 100, 2))
            else:
                sentimientos.append("Error")
                afinidades.append(0)

            if not keyp.is_error:
                etiquetas.append(", ".join(keyp.key_phrases))
            else:
                etiquetas.append("")

        df["Sentimiento"] = sentimientos
        df["Afinidad (%)"] = afinidades
        df["Etiquetas"] = etiquetas

        # Mostrar resultados en interfaz
        texto_resultados.delete(1.0, tk.END)
        for i, row in df.iterrows():
            texto_resultados.insert(tk.END, f"Comentario: {row['comentario']}\n")
            texto_resultados.insert(tk.END, f"Sentimiento: {row['Sentimiento']} | Afinidad: {row['Afinidad (%)']}%\n")
            texto_resultados.insert(tk.END, f"Etiquetas: {row['Etiquetas']}\n\n")

        messagebox.showinfo("Éxito", "Análisis completado.")

    except Exception as e:
        messagebox.showerror("Error", f"No se pudo analizar:\n{e}")

def guardar_excel():
    global df
    if df is None or "Sentimiento" not in df.columns:
        messagebox.showwarning("Advertencia", "No hay resultados para guardar.")
        return
    
    archivo = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            filetypes=[("Excel files", "*.xlsx")])
    if archivo:
        try:
            df.to_excel(archivo, index=False)
            messagebox.showinfo("Éxito", f"Archivo guardado: {archivo}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar el archivo:\n{e}")

customtkinter.set_appearance_mode("System")  
customtkinter.set_default_color_theme("blue")  

app = customtkinter.CTk()
app.title("Analizador de Comentarios de STEAM- Azure AI Language")
app.geometry("900x700")

main_frame = customtkinter.CTkFrame(master=app, corner_radius=10)
main_frame.pack(pady=20, padx=20, fill="both", expand=True)

titulo = customtkinter.CTkLabel(master=main_frame, text="Analizador de comentarios de STEAM", font=("Roboto", 24, "bold"))
titulo.pack(pady=(20, 10))

frame_botones = customtkinter.CTkFrame(master=main_frame, corner_radius=10)
frame_botones.pack(pady=10, padx=10, fill="x")

btn_cargar = customtkinter.CTkButton(master=frame_botones, text="Cargar Excel", command=cargar_excel, image=icono_cargar, compound="left", font=("Roboto", 14), fg_color="#4CAF50")
btn_cargar.pack(side="left", expand=True, padx=10, pady=10)

btn_analizar = customtkinter.CTkButton(master=frame_botones, text="Analizar Comentarios", command=analizar, image=icono_analizar, compound="left", font=("Roboto", 14), fg_color="#2196F3")
btn_analizar.pack(side="left", expand=True, padx=10, pady=10)

btn_guardar = customtkinter.CTkButton(master=frame_botones, text="Guardar Resultados", command=guardar_excel, image=icono_guardar, compound="left", font=("Roboto", 14), fg_color="#FF9800")
btn_guardar.pack(side="left", expand=True, padx=10, pady=10)

label_estado = customtkinter.CTkLabel(master=main_frame, text="No se ha cargado ningún archivo.", font=("Roboto", 12, "italic"), text_color="red")
label_estado.pack(pady=(0, 10))

label_resultados = customtkinter.CTkLabel(master=main_frame, text="Resultados del Análisis:", font=("Roboto", 16, "bold"))
label_resultados.pack(pady=(10, 5))

# --- Área de texto con scroll para los resultados ---
texto_resultados = customtkinter.CTkTextbox(master=main_frame, wrap="word", width=800, height=400, font=("Roboto Mono", 12))
texto_resultados.pack(pady=10, padx=10, fill="both", expand=True)

app.mainloop()