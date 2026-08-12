# -*- coding: utf-8 -*-
#
# Aplicación GUI (Interfaz Gráfica de Usuario) unificada para:
# 1. Gestión de Inventario (productos, stock, facturación - usa data.json)
# 2. Gestión Bovina (registro, reportes, análisis - usa SQLite)

import sys
import json
import time
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import Dict, Any, List, Optional
from datetime import datetime, date

# --- DEPENDENCIAS ADICIONALES PARA FINCA LECHERA ---
# REQUIERE: pip install pandas matplotlib openpyxl
try:
    import sqlite3
    import pandas as pd
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
except ImportError:
    messagebox.showerror("Error de Dependencia",
                         "Faltan bibliotecas requeridas (pandas, matplotlib, sqlite3). Instálelas para usar la pestaña de Gestión Bovina.")

# --- CONFIGURACIÓN Y CONSTANTES DEL INVENTARIO GENERAL ---
DATA_FILE = "data.json"

# --- CONFIGURACIÓN Y CONSTANTES DE FINCA LECHERA ---
DB_PATH = 'finca_lechera.db'
GESTACION_DIAS = 280 # Días promedio de gestación bovina

# --- UTILIDADES GENERALES (Inventario) ---

def generate_id() -> str:
    """Genera un ID único basado en el tiempo actual para productos/facturas."""
    return str(int(time.time() * 1000))

def load_data() -> Dict[str, Dict[str, Any]]:
    """Carga los datos del inventario y las facturas desde el archivo JSON."""
    try:
        with open(DATA_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
            if 'products' not in data:
                data['products'] = {}
            if 'invoices' not in data:
                data['invoices'] = {}
            return data
    except FileNotFoundError:
        return {'products': {}, 'invoices': {}}
    except json.JSONDecodeError:
        messagebox.showerror("Error de Datos", f"El archivo '{DATA_FILE}' está corrupto. Inicializando con datos vacíos.")
        return {'products': {}, 'invoices': {}}

def save_data(data: Dict[str, Dict[str, Any]]):
    """Guarda los datos del inventario y las facturas en el archivo JSON."""
    try:
        with open(DATA_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
    except IOError as e:
        messagebox.showerror("Error Crítico", f"No se pudo guardar el archivo '{DATA_FILE}': {e}")


# --- UTILIDADES Y LÓGICA DE FINCA LECHERA (SQLite/Pandas) ---

def calcular_edad_dias(fecha_nacimiento_str: str) -> Any:
    """Calcula la edad en días desde la fecha de nacimiento."""
    try:
        fn = datetime.strptime(fecha_nacimiento_str, '%Y-%m-%d').date()
        hoy = date.today()
        diferencia = hoy - fn
        return diferencia.days
    except ValueError:
        return "N/A"

def calcular_fecha_parto_est(fecha_prenez_str: str) -> str:
    """Calcula la fecha de parto estimada sumando 280 días."""
    try:
        fp = datetime.strptime(fecha_prenez_str, '%Y-%m-%d').date()
        parto_est = fp + pd.Timedelta(days=GESTACION_DIAS)
        return parto_est.strftime('%Y-%m-%d')
    except ValueError:
        return "N/A"

def obtener_status(data: Dict[str, Any]) -> str:
    """Determina el status de categorización del animal."""
    fecha_nacimiento = data['FechaNacimiento']
    esta_prenada = data['Prenada']
    fecha_prenez = data['FechaPrenez']

    edad_dias = calcular_edad_dias(fecha_nacimiento)
    if edad_dias == "N/A":
        return "Error"

    # Ternera: < 12 meses (aprox 365 días)
    if edad_dias < 365:
        return "Ternera"
   
    # Novilla sin cargar: > 12 meses y no preñada
    if edad_dias >= 365 and not esta_prenada:
        return "Novilla sin cargar (Abierta)"

    if esta_prenada:
        fecha_parto_est = calcular_fecha_parto_est(fecha_prenez)
       
        if fecha_parto_est != "N/A":
            try:
                dias_a_parto = (datetime.strptime(fecha_parto_est, '%Y-%m-%d').date() - date.today()).days
            except ValueError:
                return "Error Fecha Parto" # Manejo de error de conversión

            # Vaca Seca / Para Parir si le quedan <= 60 días para el parto
            if dias_a_parto <= 60:
                return "Vaca Seca / Para Parir"
       
        # Novilla/Vaca Preñada (en algún punto de la gestación)
        return "Vaca de Leche en Producción (Preñada)"

    # Vaca en producción (asumimos que ya parió y no está preñada)
    return "Vaca de Leche en Producción"


# --- CLASE BASE DE DATOS (SQLite) ---
class Database:
    def __init__(self):
        self.conn = sqlite3.connect(DB_PATH)
        self.cursor = self.conn.cursor()
        self.crear_tabla_animales()

    def crear_tabla_animales(self):
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS animales (
                ID_ANIMAL TEXT PRIMARY KEY,
                NOMBRE TEXT NOT NULL,
                FECHA_NACIMIENTO TEXT NOT NULL,
                MADRE_ID TEXT,
                PADRE_ID TEXT,
                PRENADA INTEGER,
                FECHA_PRENEZ TEXT,
                PADRE_PRENEZ_ID TEXT
            )
        ''')
        self.conn.commit()

    def registrar_animal(self, data: tuple) -> bool:
        try:
            self.cursor.execute('''
                INSERT INTO animales (ID_ANIMAL, NOMBRE, FECHA_NACIMIENTO, MADRE_ID, PADRE_ID,
                                      PRENADA, FECHA_PRENEZ, PADRE_PRENEZ_ID)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ''', data)
            self.conn.commit()
            return True
        except sqlite3.IntegrityError:
            return False # ID duplicado
        except Exception as e:
            print(f"Error al registrar: {e}")
            return False

    def obtener_todos_animales(self) -> List[tuple]:
        self.cursor.execute("SELECT * FROM animales")
        return self.cursor.fetchall()
   
    def obtener_columnas(self) -> List[str]:
        # Obtiene los nombres de las columnas para usarlos en Pandas y reportes
        self.cursor.execute("PRAGMA table_info(animales)")
        return [col[1] for col in self.cursor.fetchall()]
       
class InventoryApp(tk.Tk):
    """Clase principal de la aplicación GUI unificada."""
    def __init__(self):
        super().__init__()
        self.title("Gestor Integrado de Negocio (Inventario y Finca)")
        self.geometry("1200x700")
       
        # Estilos mejorados
        style = ttk.Style(self)
        style.theme_use('clam')
        style.configure("TLabel", font=("Helvetica", 10))
        style.configure("TButton", font=("Helvetica", 10, "bold"), padding=6)
        style.configure("Treeview.Heading", font=("Helvetica", 10, "bold"))
        style.configure("Treeview", rowheight=25)

        # --- INICIALIZACIÓN DE DATOS ---
        self.data: Dict[str, Dict[str, Any]] = load_data() # Inventario General (JSON)
        self.current_invoice: List[Dict[str, Any]] = []
        self.farm_db = Database() # Finca Lechera (SQLite)

        # Crear el contenedor de pestañas principal
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(pady=10, padx=10, expand=True, fill="both")
       
        # Pestañas de Inventario General (1-4)
        self.create_inventory_tab()
        self.create_add_product_tab()
        self.create_invoice_tab()
        self.create_history_tab()

        # Pestaña de Gestión Bovina (5)
        self.create_farm_tab()
       
        # Vincular el evento de cambio de pestaña para cargar contenido dinámico
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_change)

        # Cargar datos iniciales
        self.load_inventory_tree()

    # --- MANEJO DE CAMBIO DE PESTAÑA ---
    def on_tab_change(self, event):
        """Recarga datos o genera reportes al cambiar a pestañas dinámicas."""
        selected_tab_index = self.notebook.index(self.notebook.select())
       
        # Pestaña 5 (Gestión Bovina)
        if selected_tab_index == 4:
            # Dentro de la pestaña 5, verificamos qué sub-pestaña está seleccionada
            selected_sub_tab_index = self.farm_notebook.index(self.farm_notebook.select())
           
            # Sub-pestaña 2: Reportes
            if selected_sub_tab_index == 1:
                self.farm_create_reports_interface(self.farm_frame_reportes)
            # Sub-pestaña 3: Tabla de Datos
            elif selected_sub_tab_index == 2:
                self.farm_load_data_table()
       
        # Pestañas de inventario
        elif selected_tab_index == 0:
            self.load_inventory_tree()
        elif selected_tab_index == 3:
            self.load_history_tree()

    # --- UTILIDADES DE GUARDADO/RECARGA DEL INVENTARIO GENERAL ---
    def save_and_reload(self):
        """Guarda los datos en el archivo y recarga las vistas de tablas."""
        save_data(self.data)
        self.data = load_data() # Recargar para asegurar la consistencia
        self.load_inventory_tree()
        self.load_history_tree()
        self.update_invoice_display()

    # --- PESTAÑA 1-4: INVENTARIO GENERAL (Mismo Código Anterior) ---
    # (El código de create_inventory_tab, load_inventory_tree, update_product_gui,
    # update_product_field, create_add_product_tab, add_product_gui, create_invoice_tab,
    # add_item_to_invoice, update_invoice_display, generate_invoice_gui,
    # create_history_tab, load_history_tree, show_history_details va aquí.
    # Por brevedad, se omite el código idéntico del archivo anterior, pero se entiende que
    # estas funciones existen en la clase InventoryApp.)
   
    # ... (MÉTODOS create_inventory_tab, load_inventory_tree, update_product_gui, etc. AQUÍ) ...
    # INICIO DE LOS MÉTODOS DEL INVENTARIO GENERAL (Para asegurar la integridad del archivo)

    # --- PESTAÑA 1: GESTIÓN DE INVENTARIO ---
    def create_inventory_tab(self):
        self.inventory_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.inventory_frame, text="1. Inventario (Productos)")
       
        columns = ("ID", "Nombre", "Precio ($)", "Stock")
        self.inventory_tree = ttk.Treeview(self.inventory_frame, columns=columns, show="headings")
        self.inventory_tree.pack(fill="both", expand=True, pady=10)
       
        for col in columns:
            self.inventory_tree.heading(col, text=col)
            self.inventory_tree.column(col, anchor="center")
        self.inventory_tree.column("ID", width=100)
        self.inventory_tree.column("Nombre", width=150)
        self.inventory_tree.column("Precio ($)", width=100)
        self.inventory_tree.column("Stock", width=80)

        control_frame = ttk.LabelFrame(self.inventory_frame, text="Actualizar Stock o Precio", padding="10")
        control_frame.pack(fill="x", pady=10)

        self.inv_id_var = tk.StringVar()
        self.inv_field_var = tk.StringVar(value="stock")
        self.inv_value_var = tk.StringVar()

        ttk.Label(control_frame, text="ID (prefijo):").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(control_frame, textvariable=self.inv_id_var, width=15).grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        ttk.Label(control_frame, text="Campo:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        ttk.Combobox(control_frame, textvariable=self.inv_field_var, values=["stock", "price"], state="readonly", width=12).grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        ttk.Label(control_frame, text="Nuevo Valor:").grid(row=0, column=2, padx=5, pady=5, sticky="w")
        ttk.Entry(control_frame, textvariable=self.inv_value_var, width=15).grid(row=0, column=3, padx=5, pady=5, sticky="ew")

        ttk.Button(control_frame, text="Actualizar Producto", command=self.update_product_gui).grid(row=1, column=3, padx=5, pady=5, sticky="ew")
       
        control_frame.grid_columnconfigure(3, weight=1)

    def load_inventory_tree(self):
        """Carga los datos del inventario en el Treeview."""
        for item in self.inventory_tree.get_children():
            self.inventory_tree.delete(item)
           
        products = self.data.get('products', {})
        for id, prod in products.items():
            self.inventory_tree.insert("", "end", values=(
                id[:8] + '...',
                prod.get('name', 'N/A'),
                f"{prod.get('price', 0.0):.2f}",
                prod.get('stock', 0)
            ))

    def update_product_gui(self):
        """Maneja la actualización de stock/precio desde la GUI."""
        product_id_prefix = self.inv_id_var.get().strip()
        field = self.inv_field_var.get()
        value = self.inv_value_var.get().strip()
       
        if not product_id_prefix or not value:
            messagebox.showwarning("Advertencia", "Por favor, complete ID y Nuevo Valor.")
            return
       
        inventory = self.data.get('products', {})
        full_id: Optional[str] = next((id for id in inventory if id.startswith(product_id_prefix)), None)
           
        if not full_id:
            messagebox.showerror("Error", f"ID de producto no encontrado: {product_id_prefix[:8]}...")
            return
           
        self.update_product_field(full_id, field, value)

    def update_product_field(self, product_id: str, field: str, value: Any):
        """Lógica de actualización de datos (mismo que CLI, pero integrado)."""
        products = self.data.get('products', {})
       
        try:
            if field == 'price':
                update_value = max(0.0, float(value))
            elif field == 'stock':
                update_value = max(0, int(value))
            else:
                return

            products[product_id][field] = update_value
            self.data['products'] = products
            self.save_and_reload()
            messagebox.showinfo("Éxito", f"'{products[product_id]['name']}' ({field}) actualizado a {update_value}.")
        except ValueError:
            messagebox.showerror("Error de Entrada", "El valor ingresado no es un número válido.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo actualizar el producto: {e}")

    # --- PESTAÑA 2: AÑADIR PRODUCTO ---
    def create_add_product_tab(self):
        self.add_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.add_frame, text="2. Añadir Producto")
       
        form_frame = ttk.LabelFrame(self.add_frame, text="Detalles del Nuevo Producto", padding="20")
        form_frame.pack(padx=50, pady=50)

        self.prod_name_var = tk.StringVar()
        self.prod_price_var = tk.StringVar()
        self.prod_stock_var = tk.StringVar()

        fields = [
            ("Nombre del Producto:", self.prod_name_var),
            ("Precio ($):", self.prod_price_var),
            ("Stock Inicial:", self.prod_stock_var)
        ]

        for i, (label_text, variable) in enumerate(fields):
            ttk.Label(form_frame, text=label_text).grid(row=i, column=0, padx=10, pady=10, sticky="w")
            ttk.Entry(form_frame, textvariable=variable, width=30).grid(row=i, column=1, padx=10, pady=10)

        ttk.Button(form_frame, text="Agregar Producto", command=self.add_product_gui).grid(row=len(fields), column=0, columnspan=2, pady=20, sticky="ew")

    def add_product_gui(self):
        """Maneja la adición de productos desde la GUI."""
        name = self.prod_name_var.get().strip()
        price_str = self.prod_price_var.get().strip()
        stock_str = self.prod_stock_var.get().strip()

        if not name or not price_str or not stock_str:
            messagebox.showwarning("Advertencia", "Todos los campos son obligatorios.")
            return

        try:
            price = float(price_str)
            stock = int(stock_str)
           
            product_id = generate_id()
            product_data = {
                "id": product_id,
                "name": name,
                "price": max(0.0, price),
                "stock": max(0, stock)
            }
            self.data['products'][product_id] = product_data
            self.save_and_reload()
           
            messagebox.showinfo("Éxito", f"Producto '{name}' agregado y guardado.")
           
            self.prod_name_var.set("")
            self.prod_price_var.set("")
            self.prod_stock_var.set("")

        except ValueError:
            messagebox.showerror("Error de Entrada", "Precio y Stock deben ser números válidos.")

    # --- PESTAÑA 3: FACTURACIÓN ---
    def create_invoice_tab(self):
        self.invoice_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.invoice_frame, text="3. Facturación")
       
        add_item_frame = ttk.LabelFrame(self.invoice_frame, text="Añadir Item a la Factura", padding="10")
        add_item_frame.pack(fill="x", pady=10)

        self.inv_prod_id_var = tk.StringVar()
        self.inv_quantity_var = tk.StringVar()

        ttk.Label(add_item_frame, text="ID Producto (prefijo):").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(add_item_frame, textvariable=self.inv_prod_id_var, width=15).grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        ttk.Label(add_item_frame, text="Cantidad:").grid(row=0, column=2, padx=5, pady=5, sticky="w")
        ttk.Entry(add_item_frame, textvariable=self.inv_quantity_var, width=15).grid(row=0, column=3, padx=5, pady=5, sticky="ew")

        ttk.Button(add_item_frame, text="Añadir a Factura", command=self.add_item_to_invoice).grid(row=0, column=4, padx=5, pady=5, sticky="ew")
       
        add_item_frame.grid_columnconfigure(4, weight=1)

        current_invoice_frame = ttk.LabelFrame(self.invoice_frame, text="Factura Actual", padding="10")
        current_invoice_frame.pack(fill="both", expand=True, pady=10)
       
        invoice_columns = ("Cantidad", "Nombre", "Precio Unitario", "Subtotal")
        self.invoice_tree = ttk.Treeview(current_invoice_frame, columns=invoice_columns, show="headings")
        self.invoice_tree.pack(fill="both", expand=True)

        for col in invoice_columns:
            self.invoice_tree.heading(col, text=col)
            self.invoice_tree.column(col, anchor="center")
        self.invoice_tree.column("Cantidad", width=80)
        self.invoice_tree.column("Nombre", width=250, anchor="w")
       
        self.total_label = ttk.Label(current_invoice_frame, text="TOTAL: $0.00", font=("Helvetica", 12, "bold"))
        self.total_label.pack(pady=10, anchor="e")

        ttk.Button(self.invoice_frame, text="GENERAR FACTURA Y DEDUCIR INVENTARIO",
                   command=self.generate_invoice_gui, style="TButton").pack(fill="x", pady=10)
       
    def add_item_to_invoice(self):
        """Añade un producto válido a la lista temporal de la factura."""
        product_id_prefix = self.inv_prod_id_var.get().strip()
        quantity_str = self.inv_quantity_var.get().strip()
       
        if not product_id_prefix or not quantity_str:
            messagebox.showwarning("Advertencia", "Debe ingresar ID y Cantidad.")
            return

        try:
            required_quantity = int(quantity_str)
            if required_quantity <= 0:
                messagebox.showerror("Error", "La cantidad debe ser positiva.")
                return
           
            inventory = self.data.get('products', {})
            full_id = next((id for id in inventory if id.startswith(product_id_prefix)), None)
           
            if not full_id:
                messagebox.showerror("Error", "Producto no encontrado.")
                return
           
            product = inventory[full_id]

            if required_quantity > product['stock']:
                messagebox.showerror("Error de Stock", f"Stock insuficiente. Solo quedan {product['stock']} de {product['name']}.")
                return

            new_item = {
                'id': full_id,
                'name': product['name'],
                'price': product['price'],
                'quantity': required_quantity
            }
           
            existing_item_index = next((i for i, item in enumerate(self.current_invoice) if item['id'] == full_id), -1)

            if existing_item_index != -1:
                self.current_invoice[existing_item_index] = new_item
            else:
                self.current_invoice.append(new_item)
               
            self.update_invoice_display()
            self.inv_prod_id_var.set("")
            self.inv_quantity_var.set("")

        except ValueError:
            messagebox.showerror("Error de Entrada", "Cantidad debe ser un número entero.")

    def update_invoice_display(self):
        """Actualiza el Treeview de la factura actual y el total."""
        for item in self.invoice_tree.get_children():
            self.invoice_tree.delete(item)
           
        total = 0.0
        for item in self.current_invoice:
            subtotal = item['price'] * item['quantity']
            total += subtotal
            self.invoice_tree.insert("", "end", values=(
                item['quantity'],
                item['name'],
                f"${item['price']:.2f}",
                f"${subtotal:.2f}"
            ))
           
        self.total_label.config(text=f"TOTAL: ${total:.2f}")

    def generate_invoice_gui(self):
        """Ejecuta la lógica de facturación y guarda los datos."""
        if not self.current_invoice:
            messagebox.showwarning("Advertencia", "La factura está vacía.")
            return
           
        try:
            products = self.data.get('products', {})
            for item in self.current_invoice:
                product = products.get(item['id'])
                if not product or product['stock'] < item['quantity']:
                    raise Exception(f"Fallo de stock para {item['name']}. Revise el inventario.")

            total_amount = 0.0
            updates = []
           
            for item in self.current_invoice:
                products[item['id']]['stock'] -= item['quantity']
                total_amount += item['price'] * item['quantity']
                updates.append(f"{item['name']} ({item['quantity']})")

            invoice_id = generate_id()
            invoice_data = {
                "id": invoice_id,
                "items": self.current_invoice,
                "total": total_amount,
                "timestamp": time.time()
            }
            self.data['invoices'][invoice_id] = invoice_data
           
            self.save_and_reload()
           
            messagebox.showinfo("Éxito de Facturación",
                                f"Factura de ${total_amount:.2f} generada.\nInventario deducido para: {', '.join(updates)}.")
           
            self.current_invoice = []
            self.update_invoice_display()

        except Exception as e:
            messagebox.showerror("Error de Facturación", f"Fallo al procesar la facturación: {e}\nEl inventario no se modificó.")

    # --- PESTAÑA 4: HISTORIAL DE FACTURAS ---
    def create_history_tab(self):
        self.history_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.history_frame, text="4. Historial (Registros)")
       
        history_columns = ("ID Factura", "Fecha/Hora", "Total ($)", "Cantidad Items")
        self.history_tree = ttk.Treeview(self.history_frame, columns=history_columns, show="headings")
        self.history_tree.pack(fill="both", expand=True, pady=10)
       
        for col in history_columns:
            self.history_tree.heading(col, text=col)
            self.history_tree.column(col, anchor="center")
        self.history_tree.column("ID Factura", width=120)
        self.history_tree.column("Fecha/Hora", width=180)
        self.history_tree.column("Total ($)", width=100)
        self.history_tree.column("Cantidad Items", width=100)
       
        self.history_tree.bind('<<TreeviewSelect>>', self.show_history_details)
       
        self.detail_text = tk.Text(self.history_frame, height=8, state='disabled', wrap='word', font=("Helvetica", 10))
        self.detail_text.pack(fill="x", pady=10)

    def load_history_tree(self):
        """Carga el historial de facturas en el Treeview."""
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)
           
        invoices = self.data.get('invoices', {})
        if not invoices:
            self.detail_text.config(state='normal')
            self.detail_text.delete(1.0, tk.END)
            self.detail_text.insert(tk.END, "No hay registros de facturas en el historial.")
            self.detail_text.config(state='disabled')
            return

        sorted_invoices = sorted(invoices.values(), key=lambda x: x.get('timestamp', 0), reverse=True)
       
        for inv_data in sorted_invoices:
            timestamp = inv_data.get('timestamp')
            timestamp_str = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(timestamp)) if timestamp else "N/A"
           
            self.history_tree.insert("", "end", values=(
                inv_data.get('id', 'N/A')[:10] + '...',
                timestamp_str,
                f"{inv_data.get('total', 0.0):.2f}",
                len(inv_data.get('items', []))
            ), iid=inv_data.get('id'))
           
    def show_history_details(self, event):
        """Muestra los detalles de la factura seleccionada."""
        selected_item_id = self.history_tree.focus()
        if not selected_item_id:
            return
           
        invoice_id = selected_item_id
       
        invoice = self.data.get('invoices', {}).get(invoice_id)
       
        self.detail_text.config(state='normal')
        self.detail_text.delete(1.0, tk.END)

        if invoice:
            timestamp = invoice.get('timestamp')
            timestamp_str = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(timestamp)) if timestamp else "N/A"
           
            details = f"--- DETALLES DE FACTURA ---\n"
            details += f"ID Completo: {invoice.get('id')}\n"
            details += f"Fecha: {timestamp_str}\n"
            details += f"TOTAL: ${invoice.get('total', 0.0):.2f}\n"
            details += f"\nÍTEMS:\n"
           
            for item in invoice.get('items', []):
                subtotal = item.get('price', 0.0) * item.get('quantity', 0)
                details += f"  - {item.get('quantity', 0)} x {item.get('name', 'N/A')} @ ${item.get('price', 0.0):.2f} = ${subtotal:.2f}\n"

            self.detail_text.insert(tk.END, details)
           
        self.detail_text.config(state='disabled')

    # FIN DE LOS MÉTODOS DEL INVENTARIO GENERAL
   
    # --- PESTAÑA 5: GESTIÓN BOVINA (FINCA LECHERA) ---
    def create_farm_tab(self):
        self.farm_main_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.farm_main_frame, text="5. Gestión Bovina (Finca Lechera)")

        # Notebook interno para Registro, Reportes y Tabla
        self.farm_notebook = ttk.Notebook(self.farm_main_frame)
        self.farm_notebook.pack(pady=5, padx=5, expand=True, fill="both")
       
        # Pestaña 5.1: Registro
        self.farm_frame_registro = ttk.Frame(self.farm_notebook, padding="10")
        self.farm_notebook.add(self.farm_frame_registro, text='Registro de Animales')
        self.farm_create_registration_form(self.farm_frame_registro)

        # Pestaña 5.2: Reportes
        self.farm_frame_reportes = ttk.Frame(self.farm_notebook, padding="10")
        self.farm_notebook.add(self.farm_frame_reportes, text='Reportes y Gráficos')
       
        # Pestaña 5.3: Vista de Datos
        self.farm_frame_tabla = ttk.Frame(self.farm_notebook, padding="10")
        self.farm_notebook.add(self.farm_frame_tabla, text='Ver Todos los Registros')
        self.farm_create_data_table_view(self.farm_frame_tabla)
       
    # --- PESTAÑA 5.1: FORMULARIO DE REGISTRO ---
    def farm_create_registration_form(self, parent):
        tk.Label(parent, text="Registro de Animal Bovino", font=('Arial', 16, 'bold')).pack(pady=10)
       
        form_frame = ttk.Frame(parent, padding=10)
        form_frame.pack(pady=10, padx=10)

        self.farm_entries = {}
        campos = [
            ("ID Animal (Único):", 'ID_ANIMAL'),
            ("Nombre:", 'NOMBRE'),
            ("Fecha Nacimiento (YYYY-MM-DD):", 'FECHA_NACIMIENTO'),
            ("ID Madre:", 'MADRE_ID'),
            ("ID Padre:", 'PADRE_ID'),
        ]

        for i, (label_text, key) in enumerate(campos):
            ttk.Label(form_frame, text=label_text).grid(row=i, column=0, sticky='w', padx=5, pady=5)
            entry = ttk.Entry(form_frame, width=30)
            entry.grid(row=i, column=1, padx=5, pady=5)
            self.farm_entries[key] = entry
       
        ttk.Label(form_frame, text="Datos Reproductivos", font=('Arial', 12, 'bold')).grid(row=len(campos), columnspan=2, pady=(15, 5))
       
        self.farm_var_prenada = tk.IntVar(value=0)
        ttk.Checkbutton(form_frame, text="¿Está Preñada?", variable=self.farm_var_prenada,
                        command=self.farm_toggle_prenez_entries).grid(row=len(campos)+1, column=0, sticky='w', padx=5, pady=5)

        self.farm_entries['FECHA_PRENEZ'] = ttk.Entry(form_frame, width=30, state='disabled')
        ttk.Label(form_frame, text="Fecha de Preñez (YYYY-MM-DD):").grid(row=len(campos)+2, column=0, sticky='w', padx=5, pady=5)
        self.farm_entries['FECHA_PRENEZ'].grid(row=len(campos)+2, column=1, padx=5, pady=5)

        self.farm_entries['PADRE_PRENEZ_ID'] = ttk.Entry(form_frame, width=30, state='disabled')
        ttk.Label(form_frame, text="ID Padre de Preñez:").grid(row=len(campos)+3, column=0, sticky='w', padx=5, pady=5)
        self.farm_entries['PADRE_PRENEZ_ID'].grid(row=len(campos)+3, column=1, padx=5, pady=5)

        ttk.Button(parent, text="REGISTRAR ANIMAL", command=self.farm_execute_registration).pack(pady=20, ipadx=20, ipady=5)

    def farm_toggle_prenez_entries(self):
        """Habilita/Deshabilita los campos de preñez."""
        state = 'normal' if self.farm_var_prenada.get() == 1 else 'disabled'
        self.farm_entries['FECHA_PRENEZ'].config(state=state)
        self.farm_entries['PADRE_PRENEZ_ID'].config(state=state)

    def farm_execute_registration(self):
        data = {key: entry.get().strip() for key, entry in self.farm_entries.items()}
       
        if not data['ID_ANIMAL'] or not data['NOMBRE'] or not data['FECHA_NACIMIENTO']:
            messagebox.showerror("Error", "Los campos ID, Nombre y Fecha de Nacimiento son obligatorios.")
            return
       
        try:
            datetime.strptime(data['FECHA_NACIMIENTO'], '%Y-%m-%d')
        except ValueError:
            messagebox.showerror("Error de Fecha", "El formato de Fecha de Nacimiento debe ser YYYY-MM-DD.")
            return

        if self.farm_var_prenada.get() == 1:
            data['PRENADA'] = 1
            if not data['FECHA_PRENEZ']:
                messagebox.showerror("Error", "Debe ingresar la Fecha de Preñez.")
                return
            try:
                datetime.strptime(data['FECHA_PRENEZ'], '%Y-%m-%d')
            except ValueError:
                messagebox.showerror("Error de Fecha", "El formato de Fecha de Preñez debe ser YYYY-MM-DD.")
                return
        else:
            data['PRENADA'] = 0
            data['FECHA_PRENEZ'] = None
            data['PADRE_PRENEZ_ID'] = None

        registro_data = (
            data['ID_ANIMAL'], data['NOMBRE'], data['FECHA_NACIMIENTO'],
            data.get('MADRE_ID', ''), data.get('PADRE_ID', ''),
            data['PRENADA'], data['FECHA_PRENEZ'], data['PADRE_PRENEZ_ID']
        )
       
        if self.farm_db.registrar_animal(registro_data):
            messagebox.showinfo("Éxito", f"Animal {data['ID_ANIMAL']} registrado correctamente.")
            for entry in self.farm_entries.values():
                entry.delete(0, tk.END)
            self.farm_var_prenada.set(0)
            self.farm_toggle_prenez_entries()
        else:
            messagebox.showerror("Error", f"El ID de animal {data['ID_ANIMAL']} ya existe o hubo un error de DB.")

    # --- PESTAÑA 5.3: VISTA DE TABLA ---
    def farm_create_data_table_view(self, parent):
        self.farm_tree_view = ttk.Treeview(parent)
        self.farm_tree_view.pack(fill='both', expand=True, padx=5, pady=5)
       
    def farm_load_data_table(self):
        """Carga el DataFrame calculado en el Treeview de la Pestaña 5.3."""
        df = self.farm_get_complete_dataframe()
       
        if df.empty:
            # Limpiar tabla si no hay datos
            if hasattr(self, 'farm_tree_view'):
                for item in self.farm_tree_view.get_children():
                    self.farm_tree_view.delete(item)
            return

        cols = list(df.columns)
        self.farm_tree_view['columns'] = cols
        self.farm_tree_view['show'] = 'headings'

        for col in cols:
            self.farm_tree_view.heading(col, text=col)
            self.farm_tree_view.column(col, width=100 if col not in ['STATUS_GRUPO', 'FECHA_NACIMIENTO'] else 150)

        for item in self.farm_tree_view.get_children():
            self.farm_tree_view.delete(item)

        for _, row in df.iterrows():
            self.farm_tree_view.insert('', tk.END, values=list(row))
           
    # --- PESTAÑA 5.2: REPORTES Y GRÁFICOS ---
    def farm_create_reports_interface(self, parent):
        """Crea la interfaz de reportes (gráficos y exportación)."""
        for widget in parent.winfo_children():
            widget.destroy()

        df = self.farm_get_complete_dataframe()

        if df.empty:
            tk.Label(parent, text="No hay datos de animales para generar reportes.", font=('Arial', 14)).pack(pady=50)
            return

        tk.Label(parent, text="Reportes Generales de Hato", font=('Arial', 16, 'bold')).pack(pady=10)

        frame_contenido = ttk.Frame(parent)
        frame_contenido.pack(fill='both', expand=True)

        # 1. Gráfico (Izquierda)
        frame_grafico = ttk.LabelFrame(frame_contenido, text="Distribución por Grupo", padding=10)
        frame_grafico.pack(side='left', fill='both', expand=True, padx=10, pady=10)
       
        self.farm_generate_group_chart(df, frame_grafico)

        # 2. Resumen y Exportación (Derecha)
        frame_exportar = ttk.Frame(frame_contenido, padding=10)
        frame_exportar.pack(side='right', fill='y', padx=10, pady=10)
       
        ttk.Button(frame_exportar, text="EXPORTAR DATOS A EXCEL (.xlsx)",
                   command=lambda: self.farm_export_excel(df)).pack(pady=20, ipadx=10, ipady=5)
       
        tk.Label(frame_exportar, text="Resumen Numérico:", font=('Arial', 12, 'bold')).pack(pady=(20, 5))
        resumen_grupos = df['STATUS_GRUPO'].value_counts().to_dict()
       
        resumen_texto = "\n".join([f"- {k}: {v} animales" for k, v in resumen_grupos.items()])
        tk.Label(frame_exportar, text=resumen_texto, justify=tk.LEFT, anchor='w').pack(pady=5, padx=5, fill='x')
       
    def farm_generate_group_chart(self, df: pd.DataFrame, parent: ttk.Frame):
        """Genera y muestra el gráfico de Matplotlib."""
        conteo_grupos = df['STATUS_GRUPO'].value_counts()
       
        fig, ax = plt.subplots(figsize=(6, 5))
        ax.pie(conteo_grupos, labels=conteo_grupos.index, autopct='%1.1f%%', startangle=90)
        ax.set_title("Distribución de Animales por Grupo")
       
        # Integrar Matplotlib en Tkinter
        canvas = FigureCanvasTkAgg(fig, master=parent)
        canvas_widget = canvas.get_tk_widget()
        canvas_widget.pack(fill='both', expand=True)
        canvas.draw()
       
    def farm_export_excel(self, df: pd.DataFrame):
        """Exporta el DataFrame de animales a un archivo Excel."""
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Archivos Excel", "*.xlsx")],
            title="Guardar Registro de Animales"
        )
        if filepath:
            try:
                df.to_excel(filepath, index=False, engine='openpyxl')
                messagebox.showinfo("Éxito", f"Datos exportados a {filepath}")
            except Exception as e:
                messagebox.showerror("Error de Exportación", f"No se pudo exportar a Excel: {e}")

    # --- LÓGICA DE DATOS Y TRANSFORMACIÓN (PANDAS) ---
    def farm_get_complete_dataframe(self) -> pd.DataFrame:
        """Obtiene datos de la DB, los convierte a DataFrame y añade columnas calculadas."""
        columnas = self.farm_db.obtener_columnas()
        datos = self.farm_db.obtener_todos_animales()
       
        if not datos:
            return pd.DataFrame()

        df = pd.DataFrame(datos, columns=columnas)
       
        # 1. Calculo de Edad en Días
        df['EDAD_DIAS'] = df['FECHA_NACIMIENTO'].apply(calcular_edad_dias)

        # 2. Calculo de Fecha de Parto Estimada
        df['FECHA_PARTO_EST'] = df.apply(
            lambda row: calcular_fecha_parto_est(row['FECHA_PRENEZ']) if row['PRENADA'] == 1 else 'N/A',
            axis=1
        )
       
        # 3. Determinación del Status/Grupo
        df['STATUS_GRUPO'] = df.apply(
            lambda row: obtener_status({
                'FechaNacimiento': row['FECHA_NACIMIENTO'],
                'Prenada': row['PRENADA'],
                'FechaPrenez': row['FECHA_PRENEZ']
            }),
            axis=1
        )

        return df


if __name__ == "__main__":
    app = InventoryApp()
    app.mainloop()
