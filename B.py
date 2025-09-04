import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from tkcalendar import DateEntry
import pandas as pd
import os
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from fpdf import FPDF
import subprocess
import platform

# Colores de Amazon
AMAZON_ORANGE = "#FF9900"
AMAZON_DARK = "#232F3E"
AMAZON_LIGHT = "#F0F0F0"
AMAZON_TEXT = "#131921"

# Nombre del archivo Excel
EXCEL_FILE = "logistics_data.xlsx"

class LogisticsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema Integrado de Gestión Logística")
        self.root.geometry("1000x800")
        self.root.configure(bg=AMAZON_LIGHT)

        self.style = ttk.Style()
        self.style.theme_use("clam")
        self.configure_styles()
        self.initialize_excel()

        self.notebook = ttk.Notebook(root)
        self.notebook.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        self.reception_frame = ttk.Frame(self.notebook)
        self.inventory_frame = ttk.Frame(self.notebook)
        self.preparation_frame = ttk.Frame(self.notebook)
        self.dispatch_frame = ttk.Frame(self.notebook)

        self.notebook.add(self.reception_frame, text="Recepción y Almacenamiento")
        self.notebook.add(self.inventory_frame, text="Gestión de Inventario")
        self.notebook.add(self.preparation_frame, text="Preparación de Pedidos")
        self.notebook.add(self.dispatch_frame, text="Despacho y Envío")

        self.setup_reception_tab()
        self.setup_inventory_tab()
        self.setup_preparation_tab()
        self.setup_dispatch_tab()

    def initialize_excel(self):
        if not os.path.exists(EXCEL_FILE):
            inventory_data = []
            suppliers_data = [
                {"id": "1001", "name": "Bernardo Del Olmo", "delivery_time": "2 días", "cost": 60},
                {"id": "2001", "name": "Eduardo Pascal", "delivery_time": "3 días", "cost": 70},
                {"id": "3001", "name": "María Pastor", "delivery_time": "4 días", "cost": 55}
            ]
            supply_requests_data = pd.DataFrame(columns=["product_id", "quantity", "request_date", "status"])
            machines_data = [
                {"id": "SPR44", "model": "SLAM", "type": "Escaneador", "status": "Activo", "next_maintenance": "2025-05-01", "dimensions": "2.5m x 1.5m x 2m"},
                {"id": "SPR29", "model": "SLAM", "type": "Aplicar", "status": "Inactivo", "next_maintenance": "2025-06-01", "dimensions": "2m x 1m x 1.5m"}
            ]
            dispatch_history_data = [
                {"order_id": "192751947", "date": "2025-04-01", "items": "Razer Viper Mini", "status": "Enviado"},
                {"order_id": "192751948", "date": "2025-04-02", "items": "Logitech G Pro", "status": "Enviado"}
            ]
            users_data = [
                {"username": "admin", "password": "password123"},
                {"username": "user1", "password": "abc"}
            ]
            pd.DataFrame(inventory_data).to_excel(EXCEL_FILE, sheet_name="inventory", index=False)
            with pd.ExcelWriter(EXCEL_FILE, mode="a", engine="openpyxl") as writer:
                pd.DataFrame(suppliers_data).to_excel(writer, sheet_name="suppliers", index=False)
                pd.DataFrame(supply_requests_data).to_excel(writer, sheet_name="supply_requests", index=False)
                pd.DataFrame(machines_data).to_excel(writer, sheet_name="machines", index=False)
                pd.DataFrame(dispatch_history_data).to_excel(writer, sheet_name="dispatch_history", index=False)
                pd.DataFrame(users_data).to_excel(writer, sheet_name="users", index=False)
        else:
            try:
                pd.read_excel(EXCEL_FILE, sheet_name="users")
                try:
                    supply_requests_df = pd.read_excel(EXCEL_FILE, sheet_name="supply_requests")
                    required_columns = ["product_id", "quantity", "request_date", "status"]
                    missing_columns = [col for col in required_columns if col not in supply_requests_df.columns]
                    if missing_columns:
                        supply_requests_df = pd.DataFrame(columns=required_columns)
                        with pd.ExcelWriter(EXCEL_FILE, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
                            supply_requests_df.to_excel(writer, sheet_name="supply_requests", index=False)
                            for sheet in ["inventory", "suppliers", "machines", "dispatch_history", "users"]:
                                pd.read_excel(EXCEL_FILE, sheet_name=sheet).to_excel(writer, sheet_name=sheet, index=False)
                except Exception:
                    supply_requests_df = pd.DataFrame(columns=["product_id", "quantity", "request_date", "status"])
                    with pd.ExcelWriter(EXCEL_FILE, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
                        supply_requests_df.to_excel(writer, sheet_name="supply_requests", index=False)
                        for sheet in ["inventory", "suppliers", "machines", "dispatch_history", "users"]:
                            pd.read_excel(EXCEL_FILE, sheet_name=sheet).to_excel(writer, sheet_name=sheet, index=False)
            except Exception:
                users_data = [
                    {"username": "admin", "password": "password123"},
                    {"username": "user1", "password": "abc"}
                ]
                with pd.ExcelWriter(EXCEL_FILE, mode="a", engine="openpyxl", if_sheet_exists="overlay") as writer:
                    pd.DataFrame(users_data).to_excel(writer, sheet_name="users", index=False)

    def configure_styles(self):
        self.style.configure("TFrame", background=AMAZON_LIGHT)
        self.style.configure("TLabel", font=("Helvetica", 10), background=AMAZON_LIGHT, foreground=AMAZON_TEXT)
        self.style.configure("TButton", font=("Helvetica", 10, "bold"), background=AMAZON_DARK, foreground="white", padding=10)
        self.style.map("TButton", background=[("active", AMAZON_ORANGE)])
        self.style.configure("Compact.TButton", font=("Helvetica", 10, "bold"), background=AMAZON_DARK, foreground="white", padding=3)
        self.style.map("Compact.TButton", background=[("active", AMAZON_ORANGE)])
        self.style.configure("TNotebook", background=AMAZON_LIGHT)
        self.style.configure("TNotebook.Tab", font=("Helvetica", 12, "bold"), padding=[10, 5])
        self.style.map("TNotebook.Tab", background=[("selected", AMAZON_DARK)], foreground=[("selected", "white")])
        self.style.configure("TEntry", font=("Helvetica", 10), padding=5)
        self.style.configure("TLabelframe", background=AMAZON_LIGHT, bordercolor=AMAZON_DARK, borderwidth=2)
        self.style.configure("TLabelframe.Label", background=AMAZON_LIGHT, foreground=AMAZON_DARK, font=("Helvetica", 12, "bold"))
        self.style.configure("DateEntry.TEntry", background=AMAZON_DARK, foreground="white")

    def create_scrollable_frame(self, parent):
        canvas = tk.Canvas(parent, bg=AMAZON_LIGHT, highlightthickness=0)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        parent.grid_rowconfigure(0, weight=1)
        parent.grid_columnconfigure(0, weight=1)

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        return scrollable_frame

    def setup_reception_tab(self):
        scrollable_frame = self.create_scrollable_frame(self.reception_frame)
        register_frame = ttk.LabelFrame(scrollable_frame, text="Registrar Producto")
        register_frame.grid(row=0, column=0, padx=20, pady=10, sticky="ew")
        left_frame = ttk.Frame(register_frame)
        left_frame.grid(row=0, column=0, padx=10, pady=5, sticky="nsew")
        right_frame = ttk.Frame(register_frame)
        right_frame.grid(row=0, column=1, padx=10, pady=5, sticky="nsew")

        fields_left = [
            ("ID Producto:", "product_id_entry"),
            ("ID Contenedor:", "container_entry"),
            ("Nombre:", "product_name_entry"),
            ("Fecha de Entrada:", "entry_date_entry"),
            ("Proveedor:", "supplier_entry")
        ]
        for idx, (label_text, attr_name) in enumerate(fields_left):
            ttk.Label(left_frame, text=label_text).grid(row=idx, column=0, padx=5, pady=5, sticky="w")
            if attr_name == "entry_date_entry":
                setattr(self, attr_name, DateEntry(left_frame, date_pattern="dd/mm/yyyy", background=AMAZON_DARK, foreground="white", width=20))
            else:
                setattr(self, attr_name, ttk.Entry(left_frame, width=20))
            getattr(self, attr_name).grid(row=idx, column=1, padx=5, pady=5, sticky="ew")

        fields_right = [
            ("Cantidad:", "quantity_entry"),
            ("Peso (kg):", "weight_entry"),
            ("Dimensiones (LxWxH cm):", "dimensions_entry"),
            ("Precio de Compra ($):", "purchase_price_entry"),
            ("Código SKU:", "sku_entry")
        ]
        for idx, (label_text, attr_name) in enumerate(fields_right):
            ttk.Label(right_frame, text=label_text).grid(row=idx, column=0, padx=5, pady=5, sticky="w")
            setattr(self, attr_name, ttk.Entry(right_frame, width=20))
            getattr(self, attr_name).grid(row=idx, column=1, padx=5, pady=5, sticky="ew")

        button_frame = ttk.Frame(register_frame)
        button_frame.grid(row=1, column=0, columnspan=2, pady=15)
        ttk.Button(button_frame, text="Registrar", command=self.register_product).grid(row=0, column=0, padx=5)
        ttk.Button(button_frame, text="Cancelar", command=self.clear_register_form).grid(row=0, column=1, padx=5)
        ttk.Button(button_frame, text="Eliminar", command=self.open_delete_window).grid(row=0, column=2, padx=5)

        inventory_frame = ttk.LabelFrame(scrollable_frame, text="Actualizar Inventario")
        inventory_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        inventory_button_frame = ttk.Frame(inventory_frame)
        inventory_button_frame.grid(row=0, column=0, pady=5, sticky="nsew")
        button_container = ttk.Frame(inventory_button_frame)
        button_container.pack(expand=True)
        ttk.Button(button_container, text="Actualizar Inventario", command=self.update_inventory, style="Compact.TButton").pack(side="left", padx=0)
        ttk.Button(button_container, text="Abrir Excel", command=self.open_excel_file, style="Compact.TButton").pack(side="left", padx=0)
        self.inventory_output = tk.Text(inventory_frame, height=3, width=60, font=("Helvetica", 10), bg="#e0e0e0", fg="#333333")
        self.inventory_output.grid(row=1, column=0, padx=5, pady=5, sticky="ew")

        report_frame = ttk.LabelFrame(scrollable_frame, text="Generar Informe")
        report_frame.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        ttk.Label(report_frame, text="Fecha Inicio:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.report_start_date_entry = DateEntry(report_frame, date_pattern="dd/mm/yyyy", background=AMAZON_DARK, foreground="white", width=20)
        self.report_start_date_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Label(report_frame, text="Fecha Fin:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.report_end_date_entry = DateEntry(report_frame, date_pattern="dd/mm/yyyy", background=AMAZON_DARK, foreground="white", width=20)
        self.report_end_date_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(report_frame, text="Generar Informe", command=self.generate_report).grid(row=2, column=0, columnspan=2, pady=10)
        self.report_output = tk.Text(report_frame, height=3, width=60, font=("Helvetica", 10), bg="#e0e0e0", fg="#333333")
        self.report_output.grid(row=3, column=0, columnspan=2, padx=5, pady=5, sticky="ew")
        self.report_canvas = tk.Canvas(report_frame, bg=AMAZON_LIGHT, height=300, width=600)
        self.report_canvas.grid(row=4, column=0, columnspan=2, padx=5, pady=5, sticky="ew")
        ttk.Frame(scrollable_frame, height=20).grid(row=3, column=0)

        scrollable_frame.grid_columnconfigure(0, weight=1)
        register_frame.grid_columnconfigure((0, 1), weight=1)
        left_frame.grid_columnconfigure(1, weight=1)
        right_frame.grid_columnconfigure(1, weight=1)
        inventory_frame.grid_columnconfigure(0, weight=1)
        report_frame.grid_columnconfigure(1, weight=1)

    def open_excel_file(self):
        try:
            excel_file_path = os.path.abspath(EXCEL_FILE)
            if not os.path.exists(excel_file_path):
                messagebox.showerror("Error", f"El archivo {excel_file_path} no existe.", parent=self.root)
                return
            if platform.system() == "Windows":
                os.startfile(excel_file_path)
            elif platform.system() == "Darwin":
                subprocess.run(["open", excel_file_path], check=True)
            else:
                subprocess.run(["xdg-open", excel_file_path], check=True)
            messagebox.showinfo("Éxito", "El archivo Excel se ha abierto. Por favor, ciérrelo antes de realizar nuevas operaciones en la aplicación.", parent=self.root)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir el archivo: {str(e)}", parent=self.root)

    def setup_inventory_tab(self):
        scrollable_frame = self.create_scrollable_frame(self.inventory_frame)
        scrollable_frame.grid_columnconfigure(0, weight=1)
        scrollable_frame.grid_columnconfigure(1, weight=1)
        FRAME_WIDTH = 450
        FRAME_HEIGHT = 250

        add_supplier_frame = ttk.LabelFrame(scrollable_frame, text="Agregar Proveedor", width=FRAME_WIDTH, height=FRAME_HEIGHT)
        add_supplier_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        add_supplier_frame.grid_propagate(False)
        fields = [
            ("ID Proveedor:", "add_supplier_id_entry"),
            ("Nombre:", "add_supplier_name_entry"),
            ("Tiempo de Entrega:", "add_supplier_delivery_entry"),
            ("Costo ($):", "add_supplier_cost_entry"),
        ]
        for idx, (label_text, attr_name) in enumerate(fields):
            ttk.Label(add_supplier_frame, text=label_text).grid(row=idx, column=0, padx=5, pady=5, sticky="w")
            setattr(self, attr_name, ttk.Entry(add_supplier_frame, width=20))
            getattr(self, attr_name).grid(row=idx, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(add_supplier_frame, text="Agregar", command=self.add_supplier).grid(row=6, column=0, columnspan=2, pady=10, sticky="ew", padx=5)
        self.add_supplier_output = tk.Text(add_supplier_frame, height=6, width=50, font=("Helvetica", 10), bg="#e0e0e0", fg="#333333")
        self.add_supplier_output.grid(row=7, column=0, columnspan=2, padx=5, pady=5, sticky="ew")
        add_supplier_frame.grid_columnconfigure(1, weight=1)

        search_frame = ttk.LabelFrame(scrollable_frame, text="Buscar Producto", width=FRAME_WIDTH, height=FRAME_HEIGHT)
        search_frame.grid(row=1, column=0, padx=10, pady=10, sticky="ew")
        search_frame.grid_propagate(False)
        ttk.Label(search_frame, text="ID Producto:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.search_product_id_entry = ttk.Entry(search_frame, width=20)
        self.search_product_id_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(search_frame, text="Buscar", command=self.search_product).grid(row=1, column=0, columnspan=2, pady=10, sticky="ew", padx=5)
        self.search_output = tk.Text(search_frame, height=6, width=50, font=("Helvetica", 10), bg="#e0e0e0", fg="#333333")
        self.search_output.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky="ew")
        search_frame.grid_columnconfigure(1, weight=1)

        supplier_frame = ttk.LabelFrame(scrollable_frame, text="Consultar Proveedor", width=FRAME_WIDTH, height=FRAME_HEIGHT)
        supplier_frame.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        supplier_frame.grid_propagate(False)
        ttk.Label(supplier_frame, text="ID Proveedor:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.supplier_id_entry = ttk.Entry(supplier_frame, width=20)
        self.supplier_id_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(supplier_frame, text="Consultar", command=self.view_supplier).grid(row=1, column=0, columnspan=2, pady=10, sticky="ew", padx=5)
        self.supplier_output = tk.Text(supplier_frame, height=6, width=50, font=("Helvetica", 10), bg="#e0e0e0", fg="#333333")
        self.supplier_output.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky="ew")
        supplier_frame.grid_columnconfigure(1, weight=1)

        supply_frame = ttk.LabelFrame(scrollable_frame, text="Generar Solicitud de Abastecimiento", width=FRAME_WIDTH, height=FRAME_HEIGHT)
        supply_frame.grid(row=1, column=1, padx=10, pady=10, sticky="ew")
        supply_frame.grid_propagate(False)
        ttk.Label(supply_frame, text="ID Producto:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.supply_product_id_entry = ttk.Entry(supply_frame, width=20)
        self.supply_product_id_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Label(supply_frame, text="Cantidad:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.supply_quantity_entry = ttk.Entry(supply_frame, width=20)
        self.supply_quantity_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(supply_frame, text="Enviar Solicitud", command=self.send_supply_request).grid(row=2, column=0, columnspan=2, pady=10, sticky="ew", padx=5)
        self.supply_output = tk.Text(supply_frame, height=3, width=50, font=("Helvetica", 10), bg="#e0e0e0", fg="#333333")
        self.supply_output.grid(row=3, column=0, columnspan=2, padx=5, pady=5, sticky="ew")
        supply_frame.grid_columnconfigure(1, weight=1)
        ttk.Frame(scrollable_frame, height=20).grid(row=2, column=0, columnspan=2)

    def setup_preparation_tab(self):
        scrollable_frame = self.create_scrollable_frame(self.preparation_frame)
        scrollable_frame.grid_columnconfigure(0, weight=1)
        scrollable_frame.grid_columnconfigure(1, weight=1)
        FRAME_WIDTH = 450
        FRAME_HEIGHT = 250

        picking_frame = ttk.LabelFrame(scrollable_frame, text="Instrucciones de Picking", width=FRAME_WIDTH, height=FRAME_HEIGHT)
        picking_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        picking_frame.grid_propagate(False)
        ttk.Button(picking_frame, text="Ver Instrucciones de Picking", command=self.show_picking_instructions).grid(row=0, column=0, pady=5, sticky="ew", padx=5)
        self.picking_output = tk.Text(picking_frame, height=6, width=50, font=("Helvetica", 10), bg="#e0e0e0", fg="#333333")
        self.picking_output.grid(row=1, column=0, padx=5, pady=5, sticky="ew")
        picking_frame.grid_columnconfigure(0, weight=1)

        packing_frame = ttk.LabelFrame(scrollable_frame, text="Instrucciones de Packing", width=FRAME_WIDTH, height=FRAME_HEIGHT)
        packing_frame.grid(row=1, column=0, padx=10, pady=10, sticky="ew")
        packing_frame.grid_propagate(False)
        ttk.Label(packing_frame, text="Código Contenedor:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.packing_container_entry = ttk.Entry(packing_frame, width=20)
        self.packing_container_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(packing_frame, text="Ver Instrucciones", command=self.show_packing_instructions).grid(row=1, column=0, columnspan=2, pady=5, sticky="ew", padx=5)
        self.packing_output = tk.Text(packing_frame, height=6, width=50, font=("Helvetica", 10), bg="#e0e0e0", fg="#333333")
        self.packing_output.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky="ew")
        packing_frame.grid_columnconfigure(1, weight=1)

        view_container_frame = ttk.LabelFrame(scrollable_frame, text="Visualizar Contenedor", width=FRAME_WIDTH, height=FRAME_HEIGHT)
        view_container_frame.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        view_container_frame.grid_propagate(False)
        ttk.Label(view_container_frame, text="ID Contenedor:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.view_container_entry = ttk.Entry(view_container_frame, width=20)
        self.view_container_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(view_container_frame, text="Visualizar", command=self.view_container).grid(row=1, column=0, columnspan=2, pady=5, sticky="ew", padx=5)
        self.view_container_output = tk.Text(view_container_frame, height=6, width=50, font=("Helvetica", 10), bg="#e0e0e0", fg="#333333")
        self.view_container_output.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky="ew")
        view_container_frame.grid_columnconfigure(1, weight=1)

        issue_frame = ttk.LabelFrame(scrollable_frame, text="Reportar Inconveniente", width=FRAME_WIDTH, height=FRAME_HEIGHT)
        issue_frame.grid(row=1, column=1, padx=10, pady=10, sticky="ew")
        issue_frame.grid_propagate(False)
        ttk.Label(issue_frame, text="Descripción:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.issue_entry = ttk.Entry(issue_frame, width=20)
        self.issue_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(issue_frame, text="Reportar", command=self.report_issue).grid(row=1, column=0, columnspan=2, pady=5, sticky="ew", padx=5)
        self.issue_output = tk.Text(issue_frame, height=6, width=50, font=("Helvetica", 10), bg="#e0e0e0", fg="#333333")
        self.issue_output.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky="ew")
        issue_frame.grid_columnconfigure(1, weight=1)
        ttk.Frame(scrollable_frame, height=20).grid(row=2, column=0, columnspan=2)

    def setup_dispatch_tab(self):
        scrollable_frame = self.create_scrollable_frame(self.dispatch_frame)
        scrollable_frame.grid_columnconfigure(0, weight=1)
        scrollable_frame.grid_columnconfigure(1, weight=1)
        FRAME_WIDTH = 450
        FRAME_HEIGHT = 250

        period_frame = ttk.LabelFrame(scrollable_frame, text="Seleccionar Período", width=FRAME_WIDTH, height=FRAME_HEIGHT)
        period_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        period_frame.grid_propagate(False)
        ttk.Label(period_frame, text="Fecha Inicio:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.start_date_entry = DateEntry(period_frame, date_pattern="dd/mm/yyyy", background=AMAZON_DARK, foreground="white", width=20)
        self.start_date_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Label(period_frame, text="Fecha Fin:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.end_date_entry = DateEntry(period_frame, date_pattern="dd/mm/yyyy", background=AMAZON_DARK, foreground="white", width=20)
        self.end_date_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(period_frame, text="Consultar", command=self.select_period).grid(row=2, column=0, columnspan=2, pady=10, sticky="ew", padx=5)
        self.period_output = tk.Text(period_frame, height=6, width=50, font=("Helvetica", 10), bg="#e0e0e0", fg="#333333")
        self.period_output.grid(row=3, column=0, columnspan=2, padx=5, pady=5, sticky="ew")
        period_frame.grid_columnconfigure(1, weight=1)

        kpi_frame = ttk.LabelFrame(scrollable_frame, text="Consultar KPIs", width=FRAME_WIDTH, height=FRAME_HEIGHT)
        kpi_frame.grid(row=1, column=0, padx=10, pady=10, sticky="ew")
        kpi_frame.grid_propagate(False)
        ttk.Button(kpi_frame, text="Consultar KPIs", command=self.show_kpis).grid(row=0, column=0, pady=5, sticky="ew", padx=5)
        self.kpi_output = tk.Text(kpi_frame, height=6, width=50, font=("Helvetica", 10), bg="#e0e0e0", fg="#333333")
        self.kpi_output.grid(row=1, column=0, padx=5, pady=5, sticky="ew")
        kpi_frame.grid_columnconfigure(0, weight=1)

        maintenance_frame = ttk.LabelFrame(scrollable_frame, text="Consultar Mantenimiento", width=FRAME_WIDTH, height=FRAME_HEIGHT)
        maintenance_frame.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        maintenance_frame.grid_propagate(False)
        ttk.Label(maintenance_frame, text="ID Máquina:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.machine_id_entry = ttk.Entry(maintenance_frame, width=20)
        self.machine_id_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(maintenance_frame, text="Consultar", command=self.check_maintenance).grid(row=1, column=0, columnspan=2, pady=5, sticky="ew", padx=5)
        self.maintenance_output = tk.Text(maintenance_frame, height=6, width=50, font=("Helvetica", 10), bg="#e0e0e0", fg="#333333")
        self.maintenance_output.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky="ew")
        maintenance_frame.grid_columnconfigure(1, weight=1)

        sorter_frame = ttk.LabelFrame(scrollable_frame, text="Consultar Sorter", width=FRAME_WIDTH, height=FRAME_HEIGHT)
        sorter_frame.grid(row=1, column=1, padx=10, pady=10, sticky="ew")
        sorter_frame.grid_propagate(False)
        ttk.Button(sorter_frame, text="Consultar Estado", command=self.check_sorter).grid(row=0, column=0, pady=5, sticky="ew", padx=5)
        self.sorter_output = tk.Text(sorter_frame, height=6, width=50, font=("Helvetica", 10), bg="#e0e0e0", fg="#333333")
        self.sorter_output.grid(row=1, column=0, padx=5, pady=5, sticky="ew")
        sorter_frame.grid_columnconfigure(0, weight=1)
        ttk.Frame(scrollable_frame, height=20).grid(row=2, column=0, columnspan=2)

    def clear_register_form(self):
        self.product_id_entry.delete(0, tk.END)
        self.container_entry.delete(0, tk.END)
        self.product_name_entry.delete(0, tk.END)
        self.entry_date_entry.set_date(datetime.now())
        self.supplier_entry.delete(0, tk.END)
        self.quantity_entry.delete(0, tk.END)
        self.weight_entry.delete(0, tk.END)
        self.dimensions_entry.delete(0, tk.END)
        self.purchase_price_entry.delete(0, tk.END)
        self.sku_entry.delete(0, tk.END)
        messagebox.showinfo("Cancelado", "Formulario restaurado.", parent=self.root)

    def open_delete_window(self):
        inventory_df = pd.read_excel(EXCEL_FILE, sheet_name="inventory")
        if inventory_df.empty:
            messagebox.showwarning("Advertencia", "No hay productos en el inventario para eliminar.", parent=self.root)
            return
        delete_window = tk.Toplevel(self.root)
        delete_window.title("Eliminar Producto")
        delete_window.geometry("400x300")
        delete_window.configure(bg=AMAZON_LIGHT)
        frame = ttk.Frame(delete_window)
        frame.pack(padx=20, pady=20, fill="both", expand=True)
        ttk.Label(frame, text="Seleccionar Producto (ID):").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        normalized_ids = [str(id).strip() for id in inventory_df["id"].tolist()]
        self.delete_product_combobox = ttk.Combobox(frame, values=normalized_ids, state="readonly", width=20)
        self.delete_product_combobox.grid(row=0, column=1, padx=5, pady=5)
        ttk.Label(frame, text="Cantidad a Eliminar:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.delete_quantity_entry = ttk.Entry(frame, width=20)
        self.delete_quantity_entry.grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(frame, text="Eliminar", command=self.delete_product).grid(row=2, column=0, columnspan=2, pady=10)
        self.delete_output = tk.Text(frame, height=5, width=40, font=("Helvetica", 10), bg="#e0e0e0", fg="#333333")
        self.delete_output.grid(row=3, column=0, columnspan=2, padx=5, pady=5)

    def delete_product(self):
        product_id = self.delete_product_combobox.get().strip()
        quantity_to_delete = self.delete_quantity_entry.get().strip()
        if not product_id:
            messagebox.showerror("Error", "Por favor seleccione un producto.", parent=self.root)
            return
        if not quantity_to_delete:
            messagebox.showerror("Error", "Por favor ingrese la cantidad a eliminar.", parent=self.root)
            return
        try:
            quantity_to_delete = int(quantity_to_delete)
            if quantity_to_delete <= 0:
                raise ValueError("La cantidad debe ser mayor que 0.")
        except ValueError as e:
            messagebox.showerror("Error", f"La cantidad debe ser un número entero positivo: {str(e)}", parent=self.root)
            return
        try:
            inventory_df = pd.read_excel(EXCEL_FILE, sheet_name="inventory")
            inventory_df["id"] = inventory_df["id"].astype(str).str.strip()
            product_id_normalized = product_id.strip()
            product = inventory_df[inventory_df["id"] == product_id_normalized]
            if product.empty:
                messagebox.showerror("Error", f"Producto no encontrado (ID: {product_id_normalized}).", parent=self.root)
                return
            current_quantity = product.iloc[0]["quantity"]
            if quantity_to_delete > current_quantity:
                messagebox.showerror("Error", f"Solo hay {current_quantity} unidades disponibles.", parent=self.root)
                return
            self.delete_output.delete(1.0, tk.END)
            if quantity_to_delete == current_quantity:
                inventory_df = inventory_df[inventory_df["id"] != product_id_normalized]
                self.delete_output.insert(tk.END, f"Producto {product.iloc[0]['name']} (ID: {product_id_normalized}) eliminado completamente.")
            else:
                inventory_df.loc[inventory_df["id"] == product_id_normalized, "quantity"] -= quantity_to_delete
                self.delete_output.insert(tk.END, f"Se eliminaron {quantity_to_delete} unidades del producto {product.iloc[0]['name']} (ID: {product_id_normalized}).\nQuedan {inventory_df.loc[inventory_df['id'] == product_id_normalized, 'quantity'].iloc[0]} unidades.")
            with pd.ExcelWriter(EXCEL_FILE, mode="a", if_sheet_exists="replace", engine="openpyxl") as writer:
                inventory_df.to_excel(writer, sheet_name="inventory", index=False)
                for sheet in ["suppliers", "supply_requests", "machines", "dispatch_history", "users"]:
                    pd.read_excel(EXCEL_FILE, sheet_name=sheet).to_excel(writer, sheet_name=sheet, index=False)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo eliminar el producto: {str(e)}", parent=self.root)
            return
        self.delete_product_combobox.set("")
        self.delete_quantity_entry.delete(0, tk.END)

    def add_supplier(self):
        supplier_id = self.add_supplier_id_entry.get().strip()
        name = self.add_supplier_name_entry.get().strip()
        delivery_time = self.add_supplier_delivery_entry.get().strip()
        cost = self.add_supplier_cost_entry.get().strip()
        missing_fields = []
        if not supplier_id:
            missing_fields.append("ID Proveedor")
        if not name:
            missing_fields.append("Nombre")
        if not delivery_time:
            missing_fields.append("Tiempo de Entrega")
        if not cost:
            missing_fields.append("Costo")
        if missing_fields:
            messagebox.showerror("Error", f"Por favor complete los siguientes campos: {', '.join(missing_fields)}.", parent=self.root)
            return
        suppliers_df = pd.read_excel(EXCEL_FILE, sheet_name="suppliers")
        if supplier_id in suppliers_df["id"].astype(str).values:
            messagebox.showerror("Error", "El ID del proveedor ya existe. Use un ID diferente.", parent=self.root)
            return
        try:
            cost = float(cost)
        except ValueError:
            messagebox.showerror("Error", "El costo debe ser numérico.", parent=self.root)
            return
        new_supplier = pd.DataFrame([{
            "id": supplier_id,
            "name": name,
            "delivery_time": delivery_time,
            "cost": cost
        }])
        suppliers_df = pd.concat([suppliers_df, new_supplier], ignore_index=True)
        with pd.ExcelWriter(EXCEL_FILE, mode="a", if_sheet_exists="replace", engine="openpyxl") as writer:
            suppliers_df.to_excel(writer, sheet_name="suppliers", index=False)
            for sheet in ["inventory", "supply_requests", "machines", "dispatch_history", "users"]:
                pd.read_excel(EXCEL_FILE, sheet_name=sheet).to_excel(writer, sheet_name=sheet, index=False)
        self.add_supplier_output.delete(1.0, tk.END)
        self.add_supplier_output.insert(tk.END, f"Proveedor {name} (ID: {supplier_id}) agregado exitosamente.")
        self.add_supplier_id_entry.delete(0, tk.END)
        self.add_supplier_name_entry.delete(0, tk.END)
        self.add_supplier_delivery_entry.delete(0, tk.END)
        self.add_supplier_cost_entry.delete(0, tk.END)

    def register_product(self):
        product_id = self.product_id_entry.get().strip()
        container = self.container_entry.get().strip()
        product_name = self.product_name_entry.get().strip()
        entry_date = self.entry_date_entry.get().strip()
        supplier = self.supplier_entry.get().strip()
        quantity = self.quantity_entry.get().strip()
        weight = self.weight_entry.get().strip()
        dimensions = self.dimensions_entry.get().strip()
        purchase_price = self.purchase_price_entry.get().strip()
        sku = self.sku_entry.get().strip()
        missing_fields = []
        if not product_id:
            missing_fields.append("ID Producto")
        if not container:
            missing_fields.append("ID Contenedor")
        if not product_name:
            missing_fields.append("Nombre")
        if not entry_date:
            missing_fields.append("Fecha de Entrada")
        if not supplier:
            missing_fields.append("Proveedor")
        if not quantity:
            missing_fields.append("Cantidad")
        if not weight:
            missing_fields.append("Peso")
        if not dimensions:
            missing_fields.append("Dimensiones")
        if not purchase_price:
            missing_fields.append("Precio de Compra")
        if not sku:
            missing_fields.append("Código SKU")
        if missing_fields:
            messagebox.showerror("Error", f"Por favor complete los siguientes campos: {', '.join(missing_fields)}.", parent=self.root)
            return
        if not (7 <= len(sku) <= 14):
            messagebox.showerror("Error", "El Código SKU debe tener entre 7 y 14 dígitos.", parent=self.root)
            return
        if not sku.isdigit():
            messagebox.showerror("Error", "El Código SKU debe contener solo dígitos.", parent=self.root)
            return
        suppliers_df = pd.read_excel(EXCEL_FILE, sheet_name="suppliers")
        if supplier not in suppliers_df["id"].astype(str).values:
            messagebox.showerror("Error", "Proveedor no registrado.", parent=self.root)
            return
        try:
            quantity = int(quantity)
            weight = float(weight)
            purchase_price = float(purchase_price)
        except ValueError:
            messagebox.showerror("Error", "Cantidad, peso y precio deben ser numéricos.", parent=self.root)
            return
        location = container
        product = {
            "id": product_id,
            "container": str(container).strip(),
            "name": product_name,
            "entry_date": entry_date,
            "supplier": supplier,
            "quantity": quantity,
            "weight": weight,
            "dimensions": dimensions,
            "purchase_price": purchase_price,
            "sku": sku,
            "location": location
        }
        try:
            inventory_df = pd.read_excel(EXCEL_FILE, sheet_name="inventory")
            inventory_df = pd.concat([inventory_df, pd.DataFrame([product])], ignore_index=True)
            inventory_df["container"] = inventory_df["container"].astype(str)
            with pd.ExcelWriter(EXCEL_FILE, mode="a", if_sheet_exists="replace", engine="openpyxl") as writer:
                inventory_df.to_excel(writer, sheet_name="inventory", index=False)
                for sheet in ["suppliers", "supply_requests", "machines", "dispatch_history", "users"]:
                    pd.read_excel(EXCEL_FILE, sheet_name=sheet).to_excel(writer, sheet_name=sheet, index=False)
            self.inventory_output.delete(1.0, tk.END)
            self.inventory_output.insert(tk.END, f"Producto {product_name} registrado con éxito en la ubicación {container}.")
            self.product_id_entry.delete(0, tk.END)
            self.container_entry.delete(0, tk.END)
            self.product_name_entry.delete(0, tk.END)
            self.entry_date_entry.set_date(datetime.now())
            self.supplier_entry.delete(0, tk.END)
            self.quantity_entry.delete(0, tk.END)
            self.weight_entry.delete(0, tk.END)
            self.dimensions_entry.delete(0, tk.END)
            self.purchase_price_entry.delete(0, tk.END)
            self.sku_entry.delete(0, tk.END)
            messagebox.showinfo("Éxito", f"Producto {product_name} (ID: {product_id}) registrado exitosamente en la ubicación {container}.", parent=self.root)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo registrar el producto: {str(e)}", parent=self.root)
            return

    def update_inventory(self):
        try:
            inventory_df = pd.read_excel(EXCEL_FILE, sheet_name="inventory")
            total_products = len(inventory_df)
            total_quantity = inventory_df["quantity"].sum()
            self.inventory_output.delete(1.0, tk.END)
            self.inventory_output.insert(tk.END, f"Inventario Actualizado:\nTotal de Productos: {total_products}\nCantidad Total: {total_quantity}")
        except Exception as e:
            self.inventory_output.delete(1.0, tk.END)
            self.inventory_output.insert(tk.END, f"Error al actualizar el inventario: {str(e)}")

    def generate_report(self):
        try:
            # Leer todas las hojas del Excel
            inventory_df = pd.read_excel(EXCEL_FILE, sheet_name="inventory")
            suppliers_df = pd.read_excel(EXCEL_FILE, sheet_name="suppliers")
            supply_requests_df = pd.read_excel(EXCEL_FILE, sheet_name="supply_requests")
            dispatch_history_df = pd.read_excel(EXCEL_FILE, sheet_name="dispatch_history")

            # Calcular métricas
            total_products = len(inventory_df)
            total_quantity = inventory_df["quantity"].sum() if not inventory_df.empty else 0
            unique_suppliers = suppliers_df["name"].nunique() if not suppliers_df.empty else 0
            pending_supply_requests = len(supply_requests_df[supply_requests_df["status"] != "Completado"]) if not supply_requests_df.empty else 0
            total_dispatches = len(dispatch_history_df) if not dispatch_history_df.empty else 0

            # Crear el PDF con FPDF
            pdf_filename = "informe_logistica.pdf"
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", "B", 16)
            pdf.cell(0, 10, "Informe General de Logística", 0, 1, "C")
            pdf.ln(10)

            # Sección de Inventario
            pdf.set_font("Arial", "B", 12)
            pdf.cell(0, 10, "Inventario", 0, 1)
            pdf.set_font("Arial", "", 10)
            pdf.cell(0, 7, f"Total de Productos Registrados: {total_products}", 0, 1)
            pdf.cell(0, 7, f"Cantidad Total de Unidades: {total_quantity}", 0, 1)
            pdf.ln(5)

            # Sección de Proveedores
            pdf.set_font("Arial", "B", 12)
            pdf.cell(0, 10, "Proveedores", 0, 1)
            pdf.set_font("Arial", "", 10)
            pdf.cell(0, 7, f"Total de Proveedores Únicos: {unique_suppliers}", 0, 1)
            pdf.ln(5)

            # Sección de Abastecimiento
            pdf.set_font("Arial", "B", 12)
            pdf.cell(0, 10, "Abastecimiento", 0, 1)
            pdf.set_font("Arial", "", 10)
            pdf.cell(0, 7, f"Solicitudes Pendientes: {pending_supply_requests}", 0, 1)
            pdf.ln(5)

            # Sección de Despacho
            pdf.set_font("Arial", "B", 12)
            pdf.cell(0, 10, "Despacho", 0, 1)
            pdf.set_font("Arial", "", 10)
            pdf.cell(0, 7, f"Total de Órdenes Despachadas: {total_dispatches}", 0, 1)

            # Guardar el PDF
            pdf.output(pdf_filename)

            # Mostrar mensaje de éxito
            self.report_output.delete(1.0, tk.END)
            self.report_output.insert(tk.END, f"Informe generado con éxito: {pdf_filename}")

            # Abrir el PDF automáticamente
            self.open_pdf_file(pdf_filename)

        except Exception as e:
            self.report_output.delete(1.0, tk.END)
            self.report_output.insert(tk.END, f"Error al generar el informe: {str(e)}")

    def open_pdf_file(self, filename):
        try:
            if not os.path.exists(filename):
                messagebox.showerror("Error", f"El archivo {filename} no existe.", parent=self.root)
                return
            if platform.system() == "Windows":
                os.startfile(filename)
            elif platform.system() == "Darwin":
                subprocess.run(["open", filename], check=True)
            else:
                subprocess.run(["xdg-open", filename], check=True)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir el archivo PDF: {str(e)}", parent=self.root)

    def search_product(self):
        product_id = self.search_product_id_entry.get().strip()
        self.search_output.delete(1.0, tk.END)
        if not product_id:
            self.search_output.insert(tk.END, "Por favor, ingrese un ID de producto.")
            return
        try:
            inventory_df = pd.read_excel(EXCEL_FILE, sheet_name="inventory")
            inventory_df["id"] = inventory_df["id"].astype(str).str.strip()
            product_found = inventory_df[inventory_df["id"] == product_id]
            if not product_found.empty:
                info = product_found.iloc[0]
                output_text = (
                    f"Producto Encontrado:\n"
                    f"ID: {info['id']}\n"
                    f"Nombre: {info['name']}\n"
                    f"Cantidad: {info['quantity']}\n"
                    f"Ubicación: {info['container']}\n"
                    f"Proveedor: {info['supplier']}\n"
                    f"Fecha de Entrada: {info['entry_date']}\n"
                    f"Peso: {info['weight']} kg\n"
                    f"Dimensiones: {info['dimensions']} cm\n"
                    f"Precio de Compra: ${info['purchase_price']}\n"
                    f"SKU: {info['sku']}"
                )
                self.search_output.insert(tk.END, output_text)
            else:
                self.search_output.insert(tk.END, f"Producto con ID '{product_id}' no encontrado.")
        except Exception as e:
            self.search_output.insert(tk.END, f"Error al buscar producto: {str(e)}")

    def view_supplier(self):
        supplier_id = self.supplier_id_entry.get().strip()
        self.supplier_output.delete(1.0, tk.END)
        if not supplier_id:
            self.supplier_output.insert(tk.END, "Por favor, ingrese un ID de proveedor.")
            return
        try:
            suppliers_df = pd.read_excel(EXCEL_FILE, sheet_name="suppliers")
            supplier_found = suppliers_df[suppliers_df["id"].astype(str) == supplier_id]
            if not supplier_found.empty:
                info = supplier_found.iloc[0]
                output_text = (
                    f"Proveedor Encontrado:\n"
                    f"ID: {info['id']}\n"
                    f"Nombre: {info['name']}\n"
                    f"Tiempo de Entrega: {info['delivery_time']}\n"
                    f"Costo: ${info['cost']}"
                )
                self.supplier_output.insert(tk.END, output_text)
            else:
                self.supplier_output.insert(tk.END, f"Proveedor con ID '{supplier_id}' no encontrado.")
        except Exception as e:
            self.supplier_output.insert(tk.END, f"Error al consultar proveedor: {str(e)}")

    def send_supply_request(self):
        product_id = self.supply_product_id_entry.get().strip()
        quantity = self.supply_quantity_entry.get().strip()
        self.supply_output.delete(1.0, tk.END)
        if not product_id or not quantity:
            self.supply_output.insert(tk.END, "Por favor, complete ID Producto y Cantidad.")
            return
        try:
            quantity = int(quantity)
            if quantity <= 0:
                raise ValueError
            inventory_df = pd.read_excel(EXCEL_FILE, sheet_name="inventory")
            inventory_df["id"] = inventory_df["id"].astype(str).str.strip()
            product_exists = product_id in inventory_df["id"].values
            if not product_exists:
                self.supply_output.insert(tk.END, f"Producto con ID '{product_id}' no existe en el inventario.")
                return
            supply_requests_df = pd.read_excel(EXCEL_FILE, sheet_name="supply_requests")
            new_request = pd.DataFrame([{
                "product_id": product_id,
                "quantity": quantity,
                "request_date": datetime.now().strftime("%Y-%m-%d"),
                "status": "Pendiente"
            }])
            supply_requests_df = pd.concat([supply_requests_df, new_request], ignore_index=True)
            with pd.ExcelWriter(EXCEL_FILE, mode="a", if_sheet_exists="replace", engine="openpyxl") as writer:
                supply_requests_df.to_excel(writer, sheet_name="supply_requests", index=False)
                for sheet in ["inventory", "suppliers", "machines", "dispatch_history", "users"]:
                    pd.read_excel(EXCEL_FILE, sheet_name=sheet).to_excel(writer, sheet_name=sheet, index=False)
            self.supply_output.insert(tk.END, f"Solicitud de abastecimiento para {quantity} unidades del producto {product_id} enviada.")
            self.supply_product_id_entry.delete(0, tk.END)
            self.supply_quantity_entry.delete(0, tk.END)
        except ValueError:
            self.supply_output.insert(tk.END, "La cantidad debe ser un número entero positivo.")
        except Exception as e:
            self.supply_output.insert(tk.END, f"Error al enviar solicitud: {str(e)}")

    def show_picking_instructions(self):
        self.picking_output.delete(1.0, tk.END)
        instructions = (
            "Instrucciones de Picking:\n"
            "1. Diríjase a la sección A, pasillo 3, estante 5.\n"
            "2. Recoja el producto con ID 'PROD001'.\n"
            "3. Verifique la cantidad (2 unidades).\n"
            "4. Coloque los productos en el contenedor de picking 'CONT001'."
        )
        self.picking_output.insert(tk.END, instructions)

    def show_packing_instructions(self):
        container_code = self.packing_container_entry.get().strip()
        self.packing_output.delete(1.0, tk.END)
        if not container_code:
            self.packing_output.insert(tk.END, "Por favor, ingrese un Código de Contenedor.")
            return
        instructions = (
            f"Instrucciones de Packing para Contenedor '{container_code}':\n"
            f"1. Verifique que todos los productos del contenedor '{container_code}' estén presentes.\n"
            "2. Seleccione la caja de tamaño adecuado (ej. caja mediana).\n"
            "3. Rellene los espacios vacíos con material de amortiguación.\n"
            "4. Sella la caja de forma segura.\n"
            "5. Imprima y adhiera la etiqueta de envío."
        )
        self.packing_output.insert(tk.END, instructions)

    def view_container(self):
        container_id = self.view_container_entry.get().strip()
        self.view_container_output.delete(1.0, tk.END)
        if not container_id:
            self.view_container_output.insert(tk.END, "Por favor, ingrese un ID de Contenedor.")
            return
        try:
            inventory_df = pd.read_excel(EXCEL_FILE, sheet_name="inventory")
            inventory_df["container"] = inventory_df["container"].astype(str).str.strip()
            products_in_container = inventory_df[inventory_df["container"] == container_id]
            if not products_in_container.empty:
                output_text = f"Productos en Contenedor '{container_id}':\n"
                for index, row in products_in_container.iterrows():
                    output_text += f"- ID: {row['id']}, Nombre: {row['name']}, Cantidad: {row['quantity']}\n"
                self.view_container_output.insert(tk.END, output_text)
            else:
                self.view_container_output.insert(tk.END, f"Contenedor '{container_id}' no encontrado o vacío.")
        except Exception as e:
            self.view_container_output.insert(tk.END, f"Error al visualizar contenedor: {str(e)}")

    def report_issue(self):
        issue_description = self.issue_entry.get().strip()
        self.issue_output.delete(1.0, tk.END)
        if not issue_description:
            self.issue_output.insert(tk.END, "Por favor, ingrese una descripción del inconveniente.")
            return
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        report_message = f"Inconveniente reportado el {timestamp}: {issue_description}"
        self.issue_output.insert(tk.END, report_message)
        self.issue_entry.delete(0, tk.END)
        messagebox.showinfo("Reporte de Inconveniente", "Inconveniente reportado exitosamente.", parent=self.root)

    def select_period(self):
        start_date_str = self.start_date_entry.get().strip()
        end_date_str = self.end_date_entry.get().strip()
        self.period_output.delete(1.0, tk.END)
        if not start_date_str or not end_date_str:
            self.period_output.insert(tk.END, "Por favor, seleccione una fecha de inicio y fin.")
            return
        try:
            start_date = datetime.strptime(start_date_str, "%d/%m/%Y")
            end_date = datetime.strptime(end_date_str, "%d/%m/%Y")
            if start_date > end_date:
                self.period_output.insert(tk.END, "La fecha de inicio no puede ser posterior a la fecha de fin.")
                return
            dispatch_history_df = pd.read_excel(EXCEL_FILE, sheet_name="dispatch_history")
            dispatch_history_df["date"] = pd.to_datetime(dispatch_history_df["date"], format="%Y-%m-%d", errors="coerce")
            filtered_dispatches = dispatch_history_df[
                (dispatch_history_df["date"] >= start_date) & 
                (dispatch_history_df["date"] <= end_date)
            ]
            if not filtered_dispatches.empty:
                total_orders = len(filtered_dispatches)
                self.period_output.insert(tk.END, 
                                          f"Órdenes despachadas entre {start_date_str} y {end_date_str}:\n"
                                          f"Total de órdenes: {total_orders}\n\n")
                for index, row in filtered_dispatches.iterrows():
                    self.period_output.insert(tk.END, f"  - ID Orden: {row['order_id']}, Fecha: {row['date'].strftime('%Y-%m-%d')}, Items: {row['items']}\n")
            else:
                self.period_output.insert(tk.END, f"No se encontraron órdenes despachadas entre {start_date_str} y {end_date_str}.")
        except ValueError:
            self.period_output.insert(tk.END, "Formato de fecha inválido. Use DD/MM/YYYY.")
        except Exception as e:
            self.period_output.insert(tk.END, f"Error al seleccionar período: {str(e)}")

    def show_kpis(self):
        self.kpi_output.delete(1.0, tk.END)
        try:
            inventory_df = pd.read_excel(EXCEL_FILE, sheet_name="inventory")
            dispatch_history_df = pd.read_excel(EXCEL_FILE, sheet_name="dispatch_history")
            avg_dispatch_time = "N/A"
            if not dispatch_history_df.empty:
                avg_dispatch_time = "3.5 días"
            inventory_turnover = "N/A"
            if not inventory_df.empty and not dispatch_history_df.empty:
                inventory_turnover = "4.2 veces/año"
            order_accuracy_rate = "98%"
            kpi_text = (
                f"--- Indicadores Clave de Rendimiento (KPIs) ---\n"
                f"- Tiempo promedio de despacho: {avg_dispatch_time}\n"
                f"- Rotación de inventario: {inventory_turnover}\n"
                f"- Tasa de precisión de pedidos: {order_accuracy_rate}\n"
                f"- % de productos entregados: 95%\n"
                f"- Detección de anomalías: 90%"
            )
            self.kpi_output.insert(tk.END, kpi_text)
        except Exception as e:
            self.kpi_output.insert(tk.END, f"Error al consultar KPIs: {str(e)}")

    def check_maintenance(self):
        machine_id = self.machine_id_entry.get().strip()
        self.maintenance_output.delete(1.0, tk.END)
        if not machine_id:
            self.maintenance_output.insert(tk.END, "Por favor, ingrese un ID de máquina.")
            return
        try:
            machines_df = pd.read_excel(EXCEL_FILE, sheet_name="machines")
            machine = machines_df[machines_df["id"] == machine_id]
            if not machine.empty:
                info = machine.iloc[0]
                self.maintenance_output.insert(tk.END, (
                    f"Máquina: {machine_id}\n"
                    f"Modelo: {info['model']}\n"
                    f"Tipo: {info['type']}\n"
                    f"Estado: {info['status']}\n"
                    f"Próximo mantenimiento: {info['next_maintenance']}\n"
                    f"Dimensiones: {info['dimensions']}"
                ))
            else:
                self.maintenance_output.insert(tk.END, "Máquina no encontrada.")
        except Exception as e:
            self.maintenance_output.insert(tk.END, f"Error al consultar mantenimiento: {str(e)}")

    def check_sorter(self):
        self.sorter_output.delete(1.0, tk.END)
        sorter_status = {
            "status": "Activo",
            "last_used": "2025-04-25",
            "classified_orders": 150
        }
        output = (f"Estado de Sorter:\n"
                  f"Estado: {sorter_status['status']}\n"
                  f"Último uso: {sorter_status['last_used']}\n"
                  f"Órdenes clasificadas hoy: {sorter_status['classified_orders']}")
        self.sorter_output.insert(tk.END, output)

class LoginApp:
    def __init__(self, master):
        self.master = master
        master.title("Login")
        master.geometry("400x250")
        master.configure(bg=AMAZON_LIGHT)
        self.style = ttk.Style()
        self.style.theme_use("clam")
        self.configure_styles()
        self.frame = ttk.Frame(master, padding="20 20 20 20")
        self.frame.pack(expand=True)
        self.username_label = ttk.Label(self.frame, text="Usuario:")
        self.username_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.username_entry = ttk.Entry(self.frame, width=30)
        self.username_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.password_label = ttk.Label(self.frame, text="Contraseña:")
        self.password_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.password_entry = ttk.Entry(self.frame, show="*", width=30)
        self.password_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        self.login_button = ttk.Button(self.frame, text="Iniciar Sesión", command=self.check_login)
        self.login_button.grid(row=2, column=0, columnspan=2, pady=10)
        self.frame.grid_columnconfigure(1, weight=1)

    def configure_styles(self):
        self.style.configure("TFrame", background=AMAZON_LIGHT)
        self.style.configure("TLabel", font=("Helvetica", 10), background=AMAZON_LIGHT, foreground=AMAZON_TEXT)
        self.style.configure("TEntry", font=("Helvetica", 10), padding=5)
        self.style.configure("TButton", font=("Helvetica", 10, "bold"), background=AMAZON_DARK, foreground="white", padding=8)
        self.style.map("TButton", background=[("active", AMAZON_ORANGE)])

    def check_login(self):
        username = self.username_entry.get().strip()
        password = self.password_entry.get().strip()
        try:
            users_df = pd.read_excel(EXCEL_FILE, sheet_name="users")
            users_df["username"] = users_df["username"].astype(str).str.strip()
            users_df["password"] = users_df["password"].astype(str).str.strip()
            authenticated = False
            for index, row in users_df.iterrows():
                if row["username"] == username and row["password"] == password:
                    authenticated = True
                    break
            if authenticated:
                messagebox.showinfo("Inicio de Sesión", "¡Inicio de sesión exitoso!")
                self.master.destroy()
                self.launch_main_app()
            else:
                messagebox.showerror("Error de Inicio de Sesión", "Usuario o contraseña incorrectos.")
        except FileNotFoundError:
            messagebox.showerror("Error", f"El archivo {EXCEL_FILE} no se encontró.")
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error: {str(e)}")

    def launch_main_app(self):
        root = tk.Tk()
        app = LogisticsApp(root)
        root.mainloop()

if __name__ == "__main__":
    root_login = tk.Tk()
    login_app = LoginApp(root_login)
    root_login.mainloop()