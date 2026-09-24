"""
Generador de números de unicidad para productos Stahls.

VERSIÓN 1.2

El programa solicita:

- Número o nombre del pedido.
- Licenciatario.
- Uno o varios productos.
- Cantidad de hologramas de cada producto.
- Plantilla Excel con la estructura ZPL.
- Carpeta de destino.

FUNCIONAMIENTO

1. Crea una carpeta con el número o nombre del pedido.
2. Genera una única secuencia de códigos para todos los productos.
3. Genera un Excel maestro o LOG.
4. El Excel contiene una hoja por producto.
5. Cada hoja relaciona cada número de unicidad con su lote.
6. Genera archivos TXT de un máximo de 5.050 etiquetas.
7. Cada línea de los TXT contiene el ZPL completo de una etiqueta.
8. Comprueba que no existan códigos duplicados antes de guardar.

ESTRUCTURA DEL CÓDIGO

    PREFIJO + yyyyMMddHHmmssSSS

Ejemplo:

    FCB20260921083045123

El prefijo se obtiene del licenciatario:
- Se convierte a mayúsculas.
- Se eliminan espacios y caracteres especiales.
"""

import datetime
import math
import queue
import re
import threading
import time
import tkinter as tk

from pathlib import Path
from tkinter import filedialog, messagebox, ttk


# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================

WINDOW_TITLE = "Generador de unicidades Stahls - V1.2"
WINDOW_SIZE = "850x680"

MAX_CODES_PER_FILE = 5050
MAX_LICENSEE_LENGTH = 50
MAX_PRODUCT_LENGTH = 80
MAX_ORDER_LENGTH = 80

PROGRESS_UPDATE_INTERVAL = 25

DEFAULT_TEMPLATE_NAME = "plantilla holagramas.xlsx"


# ============================================================
# FUNCIONES DE NORMALIZACIÓN Y VALIDACIÓN
# ============================================================

def normalize_text(value):
    """
    Elimina espacios iniciales y finales.

    Args:
        value (str): texto que se desea normalizar.

    Returns:
        str: texto limpio.
    """
    if not isinstance(value, str):
        return ""

    return value.strip()


def sanitize_file_name(value):
    """
    Convierte un texto en un nombre válido para archivos y carpetas.

    Sustituye caracteres no permitidos en Windows por guiones bajos.

    Args:
        value (str): texto original.

    Returns:
        str: nombre seguro para archivos y carpetas.
    """
    sanitized_value = normalize_text(value)

    sanitized_value = re.sub(
        r'[<>:"/\\|?*]',
        "_",
        sanitized_value,
    )

    sanitized_value = re.sub(
        r"\s+",
        "_",
        sanitized_value,
    )

    sanitized_value = sanitized_value.strip("._ ")

    return sanitized_value


def sanitize_sheet_name(value, existing_names=None):
    """
    Convierte el nombre de un producto en un nombre válido de hoja Excel.

    Excel admite un máximo de 31 caracteres por nombre de hoja.

    Args:
        value (str): nombre del producto.
        existing_names (set[str] | None): nombres ya utilizados.

    Returns:
        str: nombre de hoja válido y no repetido.
    """
    if existing_names is None:
        existing_names = set()

    sanitized_name = normalize_text(value)

    sanitized_name = re.sub(
        r"[:\\/?*\[\]]",
        "_",
        sanitized_name,
    )

    sanitized_name = sanitized_name[:31].strip()

    if not sanitized_name:
        sanitized_name = "Producto"

    base_name = sanitized_name
    counter = 2

    while sanitized_name.lower() in {
        name.lower()
        for name in existing_names
    }:
        suffix = f"_{counter}"
        available_length = 31 - len(suffix)

        sanitized_name = (
            base_name[:available_length] + suffix
        )

        counter += 1

    return sanitized_name


def build_prefix_from_licensee(licensee):
    """
    Construye el prefijo de unicidad a partir del licenciatario.

    Solo conserva letras y números.

    Ejemplos:
        "FC Barcelona" -> "FCBARCELONA"
        "Real Madrid C.F." -> "REALMADRIDCF"

    Args:
        licensee (str): licenciatario introducido.

    Returns:
        str: prefijo normalizado.
    """
    prefix = re.sub(
        r"[^A-Za-z0-9]",
        "",
        licensee,
    ).upper()

    if not prefix:
        raise ValueError(
            "No se ha podido construir un prefijo válido "
            "a partir del licenciatario."
        )

    return prefix


def validate_order(order_number):
    """
    Valida el número o nombre del pedido.

    Args:
        order_number (str): identificador del pedido.

    Returns:
        str: pedido validado.

    Raises:
        ValueError: si el pedido no es válido.
    """
    validated_order = normalize_text(order_number)

    if not validated_order:
        raise ValueError(
            "El número o nombre del pedido no puede estar vacío."
        )

    if len(validated_order) > MAX_ORDER_LENGTH:
        raise ValueError(
            f"El pedido no puede superar "
            f"{MAX_ORDER_LENGTH} caracteres."
        )

    if not sanitize_file_name(validated_order):
        raise ValueError(
            "El pedido no contiene caracteres válidos."
        )

    return validated_order


def validate_licensee(licensee):
    """
    Valida el licenciatario.

    Args:
        licensee (str): licenciatario.

    Returns:
        str: licenciatario validado.

    Raises:
        ValueError: si no es válido.
    """
    validated_licensee = normalize_text(licensee)

    if not validated_licensee:
        raise ValueError(
            "El licenciatario no puede estar vacío."
        )

    if len(validated_licensee) > MAX_LICENSEE_LENGTH:
        raise ValueError(
            f"El licenciatario no puede superar "
            f"{MAX_LICENSEE_LENGTH} caracteres."
        )

    build_prefix_from_licensee(validated_licensee)

    return validated_licensee


def validate_product_name(product_name):
    """
    Valida el nombre del producto.

    Args:
        product_name (str): nombre del producto.

    Returns:
        str: nombre validado.

    Raises:
        ValueError: si no es válido.
    """
    validated_product = normalize_text(product_name)

    if not validated_product:
        raise ValueError(
            "El nombre del producto no puede estar vacío."
        )

    if len(validated_product) > MAX_PRODUCT_LENGTH:
        raise ValueError(
            f"El producto no puede superar "
            f"{MAX_PRODUCT_LENGTH} caracteres."
        )

    return validated_product


def validate_quantity(quantity):
    """
    Valida que la cantidad sea un entero positivo.

    Args:
        quantity (int): cantidad solicitada.

    Returns:
        int: cantidad validada.

    Raises:
        ValueError: si no es válida.
    """
    if not isinstance(quantity, int):
        raise ValueError(
            "La cantidad debe ser un número entero."
        )

    if quantity <= 0:
        raise ValueError(
            "La cantidad debe ser mayor que cero."
        )

    return quantity


# ============================================================
# LECTURA DE LA PLANTILLA
# ============================================================

def load_zpl_template(template_path):
    """
    Lee PLANT1, PLANT2, PLANT3, PLANT4 y PLANT5 de la plantilla.

    Se espera esta estructura:

        A1: PLANT1
        B1: PLANT2
        C1: PLANT3
        D1: PLANT4
        E1: PLANT5

        A2:E2: fragmentos ZPL.

    Args:
        template_path (str | Path): ruta de la plantilla Excel.

    Returns:
        dict[str, str]: fragmentos de la plantilla.

    Raises:
        FileNotFoundError: si el archivo no existe.
        ValueError: si la plantilla no tiene la estructura correcta.
        ImportError: si openpyxl no está instalado.
    """
    try:
        from openpyxl import load_workbook
    except ImportError as error:
        raise ImportError(
            "No se ha podido importar openpyxl.\n\n"
            "Instálalo con:\n"
            "python -m pip install openpyxl"
        ) from error

    template_path = Path(template_path)

    if not template_path.exists():
        raise FileNotFoundError(
            f"No se encuentra la plantilla:\n{template_path}"
        )

    workbook = load_workbook(
        template_path,
        read_only=True,
        data_only=False,
    )

    worksheet = workbook.active

    expected_headers = [
        "PLANT1",
        "PLANT2",
        "PLANT3",
        "PLANT4",
        "PLANT5",
    ]

    actual_headers = [
        worksheet.cell(
            row=1,
            column=column_number,
        ).value
        for column_number in range(1, 6)
    ]

    for expected, actual in zip(
        expected_headers,
        actual_headers,
    ):
        if normalize_text(str(actual or "")).upper() != expected:
            workbook.close()

            raise ValueError(
                "La plantilla no tiene la estructura esperada.\n\n"
                "Las columnas A:E deben contener:\n"
                "PLANT1, PLANT2, PLANT3, PLANT4 y PLANT5."
            )

    template_parts = {}

    for column_number, header in enumerate(
        expected_headers,
        start=1,
    ):
        value = worksheet.cell(
            row=2,
            column=column_number,
        ).value

        if value is None:
            workbook.close()

            raise ValueError(
                f"El fragmento {header} está vacío "
                "en la fila 2 de la plantilla."
            )

        template_parts[header] = str(value)

    workbook.close()

    return template_parts


def build_zpl_line(
    template_parts,
    unique_number,
    licensee,
    product,
):
    """
    Construye el ZPL completo de una etiqueta.

    Fórmula equivalente:

        PLANT1
        + número de unicidad
        + PLANT2
        + licenciatario
        + PLANT3
        + producto
        + PLANT4
        + número de unicidad
        + PLANT5

    Args:
        template_parts (dict[str, str]): fragmentos ZPL.
        unique_number (str): número de unicidad.
        licensee (str): licenciatario.
        product (str): producto.

    Returns:
        str: ZPL completo de una etiqueta.
    """
    return (
        template_parts["PLANT1"]
        + unique_number
        + template_parts["PLANT2"]
        + licensee
        + template_parts["PLANT3"]
        + product
        + template_parts["PLANT4"]
        + unique_number
        + template_parts["PLANT5"]
    )


# ============================================================
# GENERACIÓN DE UNICIDADES
# ============================================================

def get_current_timestamp():
    """
    Obtiene fecha y hora con precisión de milisegundos.

    Formato:
        yyyyMMddHHmmssSSS

    Returns:
        str: timestamp de 17 dígitos.
    """
    current_time = datetime.datetime.now()

    date_and_time = current_time.strftime(
        "%Y%m%d%H%M%S"
    )

    milliseconds = current_time.microsecond // 1000

    return f"{date_and_time}{milliseconds:03d}"


def generate_unique_number(prefix, last_timestamp=None):
    """
    Genera un código y espera hasta disponer de un milisegundo nuevo.

    Args:
        prefix (str): prefijo de unicidad.
        last_timestamp (str | None): último timestamp utilizado.

    Returns:
        tuple[str, str]:
            - Código completo.
            - Timestamp utilizado.
    """
    while True:
        current_timestamp = get_current_timestamp()

        if (
            last_timestamp is None
            or current_timestamp > last_timestamp
        ):
            complete_code = (
                f"{prefix}{current_timestamp}"
            )

            return complete_code, current_timestamp

        time.sleep(0.0001)


def generate_order_codes(
    products,
    prefix,
    progress_callback=None,
):
    """
    Genera una única secuencia para todos los productos del pedido.

    La secuencia no se reinicia cuando cambia el producto.

    Args:
        products (list[dict]): productos y cantidades.
        prefix (str): prefijo de unicidad.
        progress_callback (callable | None): callback de progreso.

    Returns:
        tuple[list[dict], list[str]]:
            - Productos con códigos y lotes.
            - Lista global de códigos generados.
    """
    total_quantity = sum(
        product["quantity"]
        for product in products
    )

    processed_products = []
    all_unique_numbers = []

    last_timestamp = None
    generated_count = 0

    for product in products:
        product_name = product["name"]
        product_quantity = product["quantity"]

        product_records = []

        for product_index in range(product_quantity):
            unique_number, current_timestamp = (
                generate_unique_number(
                    prefix=prefix,
                    last_timestamp=last_timestamp,
                )
            )

            last_timestamp = current_timestamp
            all_unique_numbers.append(unique_number)

            lot_number = (
                product_index // MAX_CODES_PER_FILE
            ) + 1

            product_records.append(
                {
                    "unique_number": unique_number,
                    "lot_number": lot_number,
                }
            )

            generated_count += 1

            if progress_callback is not None:
                should_update = (
                    generated_count
                    % PROGRESS_UPDATE_INTERVAL
                    == 0
                    or generated_count
                    == total_quantity
                )

                if should_update:
                    progress_callback(
                        generated_count,
                        total_quantity,
                        product_name,
                    )

        processed_products.append(
            {
                "name": product_name,
                "quantity": product_quantity,
                "records": product_records,
            }
        )

    return processed_products, all_unique_numbers


def verify_no_duplicates(unique_numbers):
    """
    Comprueba que no existan duplicados en todo el pedido.

    Args:
        unique_numbers (list[str]): códigos del pedido.

    Raises:
        RuntimeError: si encuentra duplicados.
    """
    total_codes = len(unique_numbers)
    total_unique_codes = len(set(unique_numbers))

    if total_codes != total_unique_codes:
        duplicated_count = (
            total_codes - total_unique_codes
        )

        raise RuntimeError(
            "Se han detectado códigos duplicados.\n\n"
            f"Duplicados: {duplicated_count}\n\n"
            "No se ha creado ningún archivo."
        )


# ============================================================
# CREACIÓN DE CARPETAS
# ============================================================

def prepare_order_directories(
    destination_directory,
    order_number,
):
    """
    Crea la estructura del pedido.

    Estructura:

        DESTINO/
        └── PEDIDO/
            └── LOTES/

    Args:
        destination_directory (str | Path): carpeta base.
        order_number (str): número o nombre del pedido.

    Returns:
        tuple[Path, Path]:
            - Carpeta del pedido.
            - Carpeta de lotes.
    """
    safe_order = sanitize_file_name(order_number)

    order_directory = (
        Path(destination_directory) / safe_order
    )

    lots_directory = order_directory / "LOTES"

    order_directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    lots_directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    return order_directory, lots_directory


# ============================================================
# GENERACIÓN DEL EXCEL MAESTRO
# ============================================================

def create_master_log(
    processed_products,
    order_number,
    licensee,
    output_path,
):
    """
    Crea el Excel maestro del pedido.

    Cada producto tiene su propia hoja.

    Columnas:
        A: NUMERO_UNICIDAD
        B: LOTE

    Args:
        processed_products (list[dict]): productos procesados.
        order_number (str): pedido.
        licensee (str): licenciatario.
        output_path (str | Path): archivo de salida.
    """
    try:
        from openpyxl import Workbook
        from openpyxl.styles import (
            Alignment,
            Font,
            PatternFill,
        )
        from openpyxl.worksheet.table import (
            Table,
            TableStyleInfo,
        )
    except ImportError as error:
        raise ImportError(
            "No se ha podido importar openpyxl.\n\n"
            "Instálalo con:\n"
            "python -m pip install openpyxl"
        ) from error

    workbook = Workbook()

    default_sheet = workbook.active
    workbook.remove(default_sheet)

    existing_sheet_names = set()

    header_fill = PatternFill(
        fill_type="solid",
        fgColor="1F4E78",
    )

    header_font = Font(
        color="FFFFFF",
        bold=True,
    )

    information_fill = PatternFill(
        fill_type="solid",
        fgColor="D9EAF7",
    )

    for product_index, product in enumerate(
        processed_products,
        start=1,
    ):
        product_name = product["name"]
        records = product["records"]

        sheet_name = sanitize_sheet_name(
            product_name,
            existing_names=existing_sheet_names,
        )

        existing_sheet_names.add(sheet_name)

        worksheet = workbook.create_sheet(
            title=sheet_name
        )

        # Información general
        worksheet["A1"] = "PEDIDO"
        worksheet["B1"] = order_number

        worksheet["A2"] = "LICENCIATARIO"
        worksheet["B2"] = licensee

        worksheet["A3"] = "PRODUCTO"
        worksheet["B3"] = product_name

        worksheet["A4"] = "CANTIDAD"
        worksheet["B4"] = len(records)

        worksheet["A5"] = "TOTAL LOTES"
        worksheet["B5"] = math.ceil(
            len(records) / MAX_CODES_PER_FILE
        )

        for row_number in range(1, 6):
            worksheet.cell(
                row=row_number,
                column=1,
            ).font = Font(bold=True)

            worksheet.cell(
                row=row_number,
                column=1,
            ).fill = information_fill

        # Encabezados de la tabla
        header_row = 7

        worksheet.cell(
            row=header_row,
            column=1,
            value="NUMERO_UNICIDAD",
        )

        worksheet.cell(
            row=header_row,
            column=2,
            value="LOTE",
        )

        for column_number in range(1, 3):
            header_cell = worksheet.cell(
                row=header_row,
                column=column_number,
            )

            header_cell.fill = header_fill
            header_cell.font = header_font
            header_cell.alignment = Alignment(
                horizontal="center"
            )

        # Datos
        safe_product_name = sanitize_file_name(
            product_name
        ).upper()

        for row_offset, record in enumerate(
            records,
            start=1,
        ):
            excel_row = header_row + row_offset

            lot_name = (
                f"{safe_product_name}_"
                f"LOTE_{record['lot_number']:03d}"
            )

            unique_cell = worksheet.cell(
                row=excel_row,
                column=1,
                value=record["unique_number"],
            )

            unique_cell.number_format = "@"

            worksheet.cell(
                row=excel_row,
                column=2,
                value=lot_name,
            )

        final_row = header_row + len(records)

        if records:
            table_reference = (
                f"A{header_row}:B{final_row}"
            )

            table_name = (
                f"TablaProducto{product_index}"
            )

            table = Table(
                displayName=table_name,
                ref=table_reference,
            )

            table_style = TableStyleInfo(
                name="TableStyleMedium2",
                showFirstColumn=False,
                showLastColumn=False,
                showRowStripes=True,
                showColumnStripes=False,
            )

            table.tableStyleInfo = table_style
            worksheet.add_table(table)

        worksheet.column_dimensions["A"].width = 35
        worksheet.column_dimensions["B"].width = 35

        worksheet.freeze_panes = "A8"

    workbook.save(output_path)
    workbook.close()


# ============================================================
# GENERACIÓN DE TXT POR LOTES
# ============================================================

def create_lot_txt_files(
    processed_products,
    template_parts,
    licensee,
    lots_directory,
    generation_identifier,
):
    """
    Crea los TXT de lotes para todos los productos.

    Cada TXT contiene como máximo 5.050 etiquetas.
    Cada línea contiene el ZPL completo de una etiqueta.

    Args:
        processed_products (list[dict]): productos procesados.
        template_parts (dict[str, str]): fragmentos ZPL.
        licensee (str): licenciatario.
        lots_directory (Path): carpeta LOTES.
        generation_identifier (str): fecha y hora de generación.

    Returns:
        listinformación de los archivos creados.
    """
    generated_files = []

    for product in processed_products:
        product_name = product["name"]
        records = product["records"]

        safe_product_name = sanitize_file_name(
            product_name
        ).upper()

        total_lots = math.ceil(
            len(records) / MAX_CODES_PER_FILE
        )

        for lot_number in range(1, total_lots + 1):
            start_index = (
                lot_number - 1
            ) * MAX_CODES_PER_FILE

            end_index = (
                start_index + MAX_CODES_PER_FILE
            )

            lot_records = records[
                start_index:end_index
            ]

            file_name = (
                f"{safe_product_name}_"
                f"LOTE_{lot_number:03d}_"
                f"{generation_identifier}.txt"
            )

            file_path = (
                lots_directory / file_name
            )

            zpl_lines = []

            for record in lot_records:
                zpl_line = build_zpl_line(
                    template_parts=template_parts,
                    unique_number=record[
                        "unique_number"
                    ],
                    licensee=licensee,
                    product=product_name,
                )

                zpl_lines.append(zpl_line)

            with file_path.open(
                mode="w",
                encoding="utf-8",
                newline="\n",
            ) as output_file:
                output_file.write(
                    "\n".join(zpl_lines)
                )

                output_file.write("\n")

            generated_files.append(
                {
                    "path": file_path,
                    "product": product_name,
                    "lot_number": lot_number,
                    "quantity": len(lot_records),
                }
            )

    return generated_files


# ============================================================
# FILA DINÁMICA DE PRODUCTO
# ============================================================

class ProductRow:
    """Representa una fila de producto en la interfaz."""

    def __init__(
        self,
        parent,
        row_number,
        remove_callback,
    ):
        self.parent = parent
        self.remove_callback = remove_callback

        self.name_variable = tk.StringVar()
        self.quantity_variable = tk.StringVar()

        self.name_entry = ttk.Entry(
            parent,
            textvariable=self.name_variable,
            width=42,
        )

        self.quantity_entry = ttk.Entry(
            parent,
            textvariable=self.quantity_variable,
            width=14,
        )

        self.remove_button = ttk.Button(
            parent,
            text="Eliminar",
            command=self.remove,
            width=10,
        )

        self.grid(row_number)

    def grid(self, row_number):
        """Coloca o recoloca la fila."""
        self.name_entry.grid(
            row=row_number,
            column=0,
            sticky="ew",
            padx=(0, 10),
            pady=4,
        )

        self.quantity_entry.grid(
            row=row_number,
            column=1,
            sticky="ew",
            padx=(0, 10),
            pady=4,
        )

        self.remove_button.grid(
            row=row_number,
            column=2,
            sticky="ew",
            pady=4,
        )

    def remove(self):
        """Solicita eliminar la fila."""
        self.remove_callback(self)

    def destroy(self):
        """Elimina los controles de la fila."""
        self.name_entry.destroy()
        self.quantity_entry.destroy()
        self.remove_button.destroy()

    def set_enabled(self, enabled):
        """Activa o desactiva la fila."""
        state = "normal" if enabled else "disabled"

        self.name_entry.configure(state=state)
        self.quantity_entry.configure(state=state)
        self.remove_button.configure(state=state)

    def get_data(self):
        """
        Obtiene y valida el producto.

        Returns:
            dict: nombre y cantidad.
        """
        product_name = validate_product_name(
            self.name_variable.get()
        )

        quantity_text = (
            self.quantity_variable.get().strip()
        )

        if not quantity_text:
            raise ValueError(
                f"Debes indicar la cantidad del producto "
                f"'{product_name}'."
            )

        try:
            quantity = int(quantity_text)
        except ValueError as error:
            raise ValueError(
                f"La cantidad de '{product_name}' "
                "debe ser un número entero."
            ) from error

        quantity = validate_quantity(quantity)

        return {
            "name": product_name,
            "quantity": quantity,
        }


# ============================================================
# INTERFAZ GRÁFICA
# ============================================================

class UniqueCodeGeneratorApp:
    """Aplicación gráfica del generador V1.2."""

    def __init__(self, root):
        self.root = root

        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.minsize(780, 620)

        self.order_variable = tk.StringVar()
        self.licensee_variable = tk.StringVar()

        self.template_variable = tk.StringVar(
            value=str(
                Path(__file__).with_name(
                    DEFAULT_TEMPLATE_NAME
                )
            )
        )

        self.destination_variable = tk.StringVar()

        self.status_variable = tk.StringVar(
            value=(
                "Introduce los datos del pedido "
                "y añade sus productos."
            )
        )

        self.summary_variable = tk.StringVar(
            value="Total: 0 hologramas | 0 lotes"
        )

        self.product_rows = []
        self.message_queue = queue.Queue()
        self.generation_in_progress = False

        self.configure_styles()
        self.create_widgets()

        self.add_product_row()

        self.root.after(
            50,
            self.process_thread_messages,
        )

        self.root.protocol(
            "WM_DELETE_WINDOW",
            self.close_application,
        )

    def configure_styles(self):
        """Configura los estilos."""
        style = ttk.Style()

        available_themes = style.theme_names()

        if "vista" in available_themes:
            style.theme_use("vista")
        elif "clam" in available_themes:
            style.theme_use("clam")

        style.configure(
            "Title.TLabel",
            font=("Segoe UI", 18, "bold"),
        )

        style.configure(
            "Subtitle.TLabel",
            font=("Segoe UI", 10),
            foreground="#555555",
        )

        style.configure(
            "Section.TLabel",
            font=("Segoe UI", 11, "bold"),
        )

        style.configure(
            "Generate.TButton",
            font=("Segoe UI", 11, "bold"),
            padding=10,
        )

    def create_widgets(self):
        """Crea todos los elementos de la ventana."""
        main_frame = ttk.Frame(
            self.root,
            padding=22,
        )

        main_frame.pack(
            fill="both",
            expand=True,
        )

        title_label = ttk.Label(
            main_frame,
            text="Generador de unicidades",
            style="Title.TLabel",
        )

        title_label.pack(pady=(0, 4))

        subtitle_label = ttk.Label(
            main_frame,
            text=(
                "Versión 1.2 · Pedido con varios productos "
                "y lotes de un máximo de 5.050 etiquetas"
            ),
            style="Subtitle.TLabel",
        )

        subtitle_label.pack(pady=(0, 20))

        order_frame = ttk.LabelFrame(
            main_frame,
            text="Datos del pedido",
            padding=15,
        )

        order_frame.pack(
            fill="x",
            pady=(0, 12),
        )

        ttk.Label(
            order_frame,
            text="Número o nombre del pedido:",
        ).grid(
            row=0,
            column=0,
            sticky="w",
            padx=(0, 12),
            pady=6,
        )

        self.order_entry = ttk.Entry(
            order_frame,
            textvariable=self.order_variable,
        )

        self.order_entry.grid(
            row=0,
            column=1,
            sticky="ew",
            pady=6,
        )

        ttk.Label(
            order_frame,
            text="Licenciatario:",
        ).grid(
            row=1,
            column=0,
            sticky="w",
            padx=(0, 12),
            pady=6,
        )

        self.licensee_entry = ttk.Entry(
            order_frame,
            textvariable=self.licensee_variable,
        )

        self.licensee_entry.grid(
            row=1,
            column=1,
            sticky="ew",
            pady=6,
        )

        order_frame.columnconfigure(1, weight=1)

        products_frame = ttk.LabelFrame(
            main_frame,
            text="Productos del pedido",
            padding=15,
        )

        products_frame.pack(
            fill="both",
            expand=True,
            pady=(0, 12),
        )

        product_headers = ttk.Frame(
            products_frame,
        )

        product_headers.pack(fill="x")

        ttk.Label(
            product_headers,
            text="Producto",
            style="Section.TLabel",
        ).grid(
            row=0,
            column=0,
            sticky="w",
        )

        ttk.Label(
            product_headers,
            text="Cantidad",
            style="Section.TLabel",
        ).grid(
            row=0,
            column=1,   #Revisar el 23/09/2026 si se puede poner decimales en este campo con el objetivo de centrar el texto
            sticky="w",
        )

        product_headers.columnconfigure(
            0,
            weight=1,
        )

        product_headers.columnconfigure(
            1,
            minsize=130,
        )

        product_headers.columnconfigure(
            2,
            minsize=90,
        )

        self.product_rows_frame = ttk.Frame(
            products_frame,
        )

        self.product_rows_frame.pack(
            fill="both",
            expand=True,
            pady=(5, 8),
        )

        self.product_rows_frame.columnconfigure(
            0,
            weight=1,
        )

        self.add_product_button = ttk.Button(
            products_frame,
            text="+ Añadir producto",
            command=self.add_product_row,
        )

        self.add_product_button.pack(
            anchor="w",
            pady=(5, 0),
        )

        configuration_frame = ttk.LabelFrame(
            main_frame,
            text="Archivos",
            padding=15,
        )

        configuration_frame.pack(
            fill="x",
            pady=(0, 12),
        )

        ttk.Label(
            configuration_frame,
            text="Plantilla Excel:",
        ).grid(
            row=0,
            column=0,
            sticky="w",
            padx=(0, 12),
            pady=6,
        )

        self.template_entry = ttk.Entry(
            configuration_frame,
            textvariable=self.template_variable,
        )

        self.template_entry.grid(
            row=0,
            column=1,
            sticky="ew",
            pady=6,
        )

        self.template_button = ttk.Button(
            configuration_frame,
            text="Seleccionar",
            command=self.select_template,
        )

        self.template_button.grid(
            row=0,
            column=2,
            padx=(10, 0),
            pady=6,
        )

        ttk.Label(
            configuration_frame,
            text="Carpeta de destino:",
        ).grid(
            row=1,
            column=0,
            sticky="w",
            padx=(0, 12),
            pady=6,
        )

        self.destination_entry = ttk.Entry(
            configuration_frame,
            textvariable=self.destination_variable,
        )

        self.destination_entry.grid(
            row=1,
            column=1,
            sticky="ew",
            pady=6,
        )

        self.destination_button = ttk.Button(
            configuration_frame,
            text="Seleccionar",
            command=self.select_destination,
        )

        self.destination_button.grid(
            row=1,
            column=2,
            padx=(10, 0),
            pady=6,
        )

        configuration_frame.columnconfigure(
            1,
            weight=1,
        )

        summary_label = ttk.Label(
            main_frame,
            textvariable=self.summary_variable,
            anchor="center",
        )

        summary_label.pack(fill="x", pady=(0, 8))

        self.progress_bar = ttk.Progressbar(
            main_frame,
            orient="horizontal",
            mode="determinate",
            maximum=100,
        )

        self.progress_bar.pack(
            fill="x",
            pady=(0, 8),
        )

        self.status_label = ttk.Label(
            main_frame,
            textvariable=self.status_variable,
            anchor="center",
        )

        self.status_label.pack(
            fill="x",
            pady=(0, 12),
        )

        self.generate_button = ttk.Button(
            main_frame,
            text="Generar pedido",
            command=self.start_generation,
            style="Generate.TButton",
        )

        self.generate_button.pack(fill="x")

        self.order_entry.focus_set()

    def add_product_row(self):
        """Añade una fila de producto."""
        product_row = ProductRow(
            parent=self.product_rows_frame,
            row_number=len(self.product_rows),
            remove_callback=self.remove_product_row,
        )

        product_row.quantity_variable.trace_add(
            "write",
            lambda *args: self.update_summary(),
        )

        self.product_rows.append(product_row)

        product_row.name_entry.focus_set()

        self.update_summary()

    def remove_product_row(self, product_row):
        """Elimina una fila de producto."""
        if len(self.product_rows) <= 1:
            messagebox.showwarning(
                "Producto obligatorio",
                (
                    "El pedido debe contener al menos "
                    "un producto."
                ),
            )
            return

        product_row.destroy()
        self.product_rows.remove(product_row)

        self.reorder_product_rows()
        self.update_summary()

    def reorder_product_rows(self):
        """Recoloca las filas después de eliminar una."""
        for row_number, product_row in enumerate(
            self.product_rows
        ):
            product_row.grid(row_number)

    def update_summary(self):
        """Actualiza el total de hologramas y lotes."""
        total_quantity = 0
        total_lots = 0

        for product_row in self.product_rows:
            quantity_text = (
                product_row.quantity_variable
                .get()
                .strip()
            )

            try:
                quantity = int(quantity_text)

                if quantity > 0:
                    total_quantity += quantity

                    total_lots += math.ceil(
                        quantity
                        / MAX_CODES_PER_FILE
                    )

            except ValueError:
                continue

        self.summary_variable.set(
            f"Total: {total_quantity} hologramas "
            f"| {total_lots} lotes"
        )

    def select_template(self):
        """Selecciona la plantilla Excel."""
        selected_file = filedialog.askopenfilename(
            title="Selecciona la plantilla",
            filetypes=[
                ("Archivos Excel", "*.xlsx"),
                ("Todos los archivos", "*.*"),
            ],
        )

        if selected_file:
            self.template_variable.set(
                selected_file
            )

    def select_destination(self):
        """Selecciona la carpeta de destino."""
        selected_directory = (
            filedialog.askdirectory(
                title=(
                    "Selecciona la carpeta "
                    "de destino"
                )
            )
        )

        if selected_directory:
            self.destination_variable.set(
                selected_directory
            )

    def get_form_data(self):
        """
        Lee y valida todos los campos.

        Returns:
            tuple: datos validados.
        """
        order_number = validate_order(
            self.order_variable.get()
        )

        licensee = validate_licensee(
            self.licensee_variable.get()
        )

        products = []

        product_names = set()

        for row_number, product_row in enumerate(
            self.product_rows,
            start=1,
        ):
            try:
                product_data = (
                    product_row.get_data()
                )
            except ValueError as error:
                raise ValueError(
                    f"Fila de producto {row_number}:\n"
                    f"{error}"
                ) from error

            comparable_name = (
                product_data["name"]
                .strip()
                .lower()
            )

            if comparable_name in product_names:
                raise ValueError(
                    "No puedes introducir dos veces "
                    f"el producto "
                    f"'{product_data['name']}'.\n\n"
                    "Agrupa la cantidad en una sola fila."
                )

            product_names.add(comparable_name)
            products.append(product_data)

        template_path = Path(
            self.template_variable.get().strip()
        )

        if not template_path.exists():
            raise ValueError(
                "No se encuentra la plantilla Excel."
            )

        destination_directory = Path(
            self.destination_variable.get().strip()
        )

        if not str(destination_directory).strip():
            raise ValueError(
                "Debes seleccionar una carpeta "
                "de destino."
            )

        return (
            order_number,
            licensee,
            products,
            template_path,
            destination_directory,
        )

    def start_generation(self):
        """Valida e inicia la generación."""
        if self.generation_in_progress:
            return

        try:
            (
                order_number,
                licensee,
                products,
                template_path,
                destination_directory,
            ) = self.get_form_data()

            total_quantity = sum(
                product["quantity"]
                for product in products
            )

            total_lots = sum(
                math.ceil(
                    product["quantity"]
                    / MAX_CODES_PER_FILE
                )
                for product in products
            )

        except ValueError as error:
            messagebox.showerror(
                "Datos incorrectos",
                str(error),
            )
            return

        self.generation_in_progress = True
        self.set_interface_enabled(False)

        self.progress_bar["value"] = 0

        self.status_variable.set(
            f"Generando 0 de {total_quantity} "
            f"códigos en {total_lots} lotes..."
        )

        generation_thread = threading.Thread(
            target=self.generate_order,
            args=(
                order_number,
                licensee,
                products,
                template_path,
                destination_directory,
            ),
            daemon=True,
        )

        generation_thread.start()

    def notify_progress(
        self,
        current,
        total,
        product_name,
    ):
        """Envía el progreso al hilo principal."""
        self.message_queue.put(
            (
                "progress",
                {
                    "current": current,
                    "total": total,
                    "product": product_name,
                },
            )
        )

    def generate_order(
        self,
        order_number,
        licensee,
        products,
        template_path,
        destination_directory,
    ):
        """Genera todos los archivos del pedido."""
        try:
            template_parts = load_zpl_template(
                template_path
            )

            prefix = build_prefix_from_licensee(
                licensee
            )

            (
                processed_products,
                all_unique_numbers,
            ) = generate_order_codes(
                products=products,
                prefix=prefix,
                progress_callback=self.notify_progress,
            )

            verify_no_duplicates(
                all_unique_numbers
            )

            (
                order_directory,
                lots_directory,
            ) = prepare_order_directories(
                destination_directory=(
                    destination_directory
                ),
                order_number=order_number,
            )

            generation_identifier = (
                datetime.datetime.now().strftime(
                    "%Y%m%d_%H%M%S"
                )
            )

            safe_order = sanitize_file_name(
                order_number
            )

            master_log_path = (
                order_directory
                / (
                    f"LOG_{safe_order}_"
                    f"{generation_identifier}.xlsx"
                )
            )

            create_master_log(
                processed_products=(
                    processed_products
                ),
                order_number=order_number,
                licensee=licensee,
                output_path=master_log_path,
            )

            generated_lot_files = (
                create_lot_txt_files(
                    processed_products=(
                        processed_products
                    ),
                    template_parts=template_parts,
                    licensee=licensee,
                    lots_directory=lots_directory,
                    generation_identifier=(
                        generation_identifier
                    ),
                )
            )

            self.message_queue.put(
                (
                    "completed",
                    {
                        "order_directory": (
                            order_directory
                        ),
                        "master_log": (
                            master_log_path
                        ),
                        "lot_files": (
                            generated_lot_files
                        ),
                        "total_codes": len(
                            all_unique_numbers
                        ),
                        "first_code": (
                            all_unique_numbers[0]
                        ),
                        "last_code": (
                            all_unique_numbers[-1]
                        ),
                        "products": len(
                            processed_products
                        ),
                    },
                )
            )

        except Exception as error:
            self.message_queue.put(
                (
                    "error",
                    {
                        "message": str(error),
                    },
                )
            )

    def process_thread_messages(self):
        """Procesa los mensajes del hilo secundario."""
        try:
            while True:
                (
                    message_type,
                    message_data,
                ) = self.message_queue.get_nowait()

                if message_type == "progress":
                    self.update_progress(
                        current=message_data[
                            "current"
                        ],
                        total=message_data["total"],
                        product_name=message_data[
                            "product"
                        ],
                    )

                elif message_type == "completed":
                    self.generation_completed(
                        message_data
                    )

                elif message_type == "error":
                    self.generation_failed(
                        message_data["message"]
                    )

        except queue.Empty:
            pass

        try:
            if self.root.winfo_exists():
                self.root.after(
                    50,
                    self.process_thread_messages,
                )
        except tk.TclError:
            pass

    def update_progress(
        self,
        current,
        total,
        product_name,
    ):
        """Actualiza el progreso visual."""
        percentage = (
            current / total
        ) * 100

        self.progress_bar["value"] = (
            percentage
        )

        self.status_variable.set(
            f"Generando {current} de {total} "
            f"códigos · Producto: {product_name}"
        )

    def generation_completed(self, result):
        """Muestra el resumen final."""
        self.generation_in_progress = False
        self.set_interface_enabled(True)

        self.progress_bar["value"] = 100

        total_codes = result["total_codes"]
        total_lots = len(result["lot_files"])
        total_products = result["products"]

        self.status_variable.set(
            f"Completado: {total_codes} códigos, "
            f"{total_products} productos y "
            f"{total_lots} lotes."
        )

        lot_summary = []

        for lot_information in result[
            "lot_files"
        ]:
            lot_summary.append(
                f"{lot_information['path'].name}: "
                f"{lot_information['quantity']} códigos"
            )

        visible_lots = lot_summary[:10]

        lot_text = "\n".join(visible_lots)

        if len(lot_summary) > 10:
            lot_text += (
                f"\n... y "
                f"{len(lot_summary) - 10} "
                "archivo(s) más."
            )

        messagebox.showinfo(
            "Pedido generado",
            (
                f"Códigos generados: {total_codes}\n"
                f"Productos: {total_products}\n"
                f"Lotes TXT: {total_lots}\n\n"
                f"Primer código:\n"
                f"{result['first_code']}\n\n"
                f"Último código:\n"
                f"{result['last_code']}\n\n"
                f"Excel maestro:\n"
                f"{result['master_log'].name}\n\n"
                f"Carpeta del pedido:\n"
                f"{result['order_directory']}\n\n"
                f"Archivos de lotes:\n"
                f"{lot_text}"
            ),
        )

    def generation_failed(self, error_message):
        """Muestra un error de generación."""
        self.generation_in_progress = False
        self.set_interface_enabled(True)

        self.status_variable.set(
            "Se ha producido un error."
        )

        messagebox.showerror(
            "Error de generación",
            error_message,
        )

    def set_interface_enabled(self, enabled):
        """Activa o desactiva la interfaz."""
        state = (
            "normal"
            if enabled
            else "disabled"
        )

        self.order_entry.configure(state=state)
        self.licensee_entry.configure(
            state=state
        )

        self.template_entry.configure(
            state=state
        )

        self.destination_entry.configure(
            state=state
        )

        self.template_button.configure(
            state=state
        )

        self.destination_button.configure(
            state=state
        )

        self.add_product_button.configure(
            state=state
        )

        self.generate_button.configure(
            state=state
        )

        for product_row in self.product_rows:
            product_row.set_enabled(enabled)

    def close_application(self):
        """Gestiona el cierre de la aplicación."""
        if self.generation_in_progress:
            should_close = messagebox.askyesno(
                "Generación en curso",
                (
                    "Hay una generación en curso.\n\n"
                    "Si cierras ahora, el pedido podría "
                    "quedar incompleto.\n\n"
                    "¿Quieres cerrar la aplicación?"
                ),
            )

            if not should_close:
                return

        self.root.destroy()


# ============================================================
# INICIO
# ============================================================

def main():
    """Inicia la aplicación."""
    root = tk.Tk()

    UniqueCodeGeneratorApp(root)

    root.mainloop()


if __name__ == "__main__":
    main()