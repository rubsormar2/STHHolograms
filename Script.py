"""
Generador de códigos únicos para productos Stahls.

Estructura del código:
    PREFIJO + yyyyMMddHHmmssSSS

Ejemplo:
    FCB20260918094530123

Donde:
- PREFIJO: valor introducido por el usuario.
- yyyy: año.
- MM: mes.
- dd: día.
- HH: hora.
- mm: minutos.
- ss: segundos.
- SSS: milisegundos.

Características de esta versión:
- Interfaz gráfica desarrollada con Tkinter.
- Generación de un único código por milisegundo.
- Prefijo configurable.
- Exportación a TXT o Excel.
- Máximo de 5.050 códigos por archivo.
- División automática en varios archivos.
- Comprobación de duplicados antes de guardar.
- Barra de progreso.
- Generación ejecutada en un hilo independiente.
"""

import datetime
import math
import queue
import threading
import time
import tkinter as tk

from pathlib import Path
from tkinter import filedialog, messagebox, ttk


# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================

WINDOW_TITLE = "Generador de códigos únicos Stahls"
WINDOW_SIZE = "580x500"

MAX_PREFIX_LENGTH = 20
MAX_CODES_PER_FILE = 5050

PROGRESS_UPDATE_INTERVAL = 25


# ============================================================
# VALIDACIONES
# ============================================================

def validate_quantity(quantity):
    """
    Valida que la cantidad solicitada sea un número entero positivo.

    Args:
        quantity (int): cantidad solicitada.

    Returns:
        int: cantidad validada.

    Raises:
        ValueError: si la cantidad no es válida.
    """
    if not isinstance(quantity, int):
        raise ValueError("La cantidad debe ser un número entero.")

    if quantity <= 0:
        raise ValueError("La cantidad debe ser mayor que 0.")

    return quantity


def validate_prefix(prefix):
    """
    Valida y normaliza el prefijo.

    El prefijo:
    - No puede estar vacío.
    - Se convierte automáticamente a mayúsculas.
    - Solo puede contener letras y números.
    - No puede superar la longitud máxima configurada.

    Args:
        prefix (str): prefijo escrito por el usuario.

    Returns:
        str: prefijo validado.

    Raises:
        ValueError: si el prefijo no es válido.
    """
    if not isinstance(prefix, str):
        raise ValueError("El prefijo debe ser texto.")

    validated_prefix = prefix.strip().upper()

    if not validated_prefix:
        raise ValueError("El prefijo no puede estar vacío.")

    if len(validated_prefix) > MAX_PREFIX_LENGTH:
        raise ValueError(
            f"El prefijo no puede superar los "
            f"{MAX_PREFIX_LENGTH} caracteres."
        )

    if not validated_prefix.isalnum():
        raise ValueError(
            "El prefijo solo puede contener letras y números."
        )

    return validated_prefix


# ============================================================
# GENERACIÓN DE CÓDIGOS
# ============================================================

def get_current_timestamp():
    """
    Obtiene el timestamp actual con precisión de milisegundos.

    Estructura:
        yyyyMMddHHmmssSSS

    Ejemplo:
        20260918094530123

    Returns:
        str: timestamp de 17 dígitos.
    """
    current_time = datetime.datetime.now()

    date_and_time = current_time.strftime("%Y%m%d%H%M%S")
    milliseconds = current_time.microsecond // 1000

    return f"{date_and_time}{milliseconds:03d}"


def generate_unique_number(prefix, last_timestamp=None):
    """
    Genera un código único utilizando el prefijo y el reloj del sistema.

    Si el milisegundo actual ya se utilizó, el programa espera hasta
    que el reloj alcance un milisegundo posterior.

    Args:
        prefix (str): prefijo validado.
        last_timestamp (str | None): último timestamp utilizado.

    Returns:
        tuple[str, str]:
            - Código completo.
            - Timestamp utilizado, sin el prefijo.
    """
    while True:
        current_timestamp = get_current_timestamp()

        if (
            last_timestamp is None
            or current_timestamp > last_timestamp
        ):
            complete_code = f"{prefix}{current_timestamp}"

            return complete_code, current_timestamp

        # Se vuelve a consultar el reloj hasta que cambie
        # el milisegundo.
        time.sleep(0.0001)


def generate_unique_numbers(
    quantity,
    prefix,
    progress_callback=None,
):
    """
    Genera una lista completa de códigos únicos.

    La variable last_timestamp se mantiene durante toda la generación.
    Por tanto, la posterior división en archivos no reinicia el
    timestamp ni afecta a la unicidad.

    Args:
        quantity (int): cantidad total de códigos.
        prefix (str): prefijo solicitado.
        progress_callback (callable | None): función para comunicar
            el progreso.

    Returns:
        listcódigos únicos generados.
    """
    validated_quantity = validate_quantity(quantity)
    validated_prefix = validate_prefix(prefix)

    generated_numbers = []
    last_timestamp = None

    for index in range(validated_quantity):
        unique_number, current_timestamp = generate_unique_number(
            prefix=validated_prefix,
            last_timestamp=last_timestamp,
        )

        generated_numbers.append(unique_number)
        last_timestamp = current_timestamp

        generated_count = index + 1

        # Evita enviar miles de actualizaciones innecesarias
        # a la interfaz gráfica.
        if progress_callback is not None:
            if (
                generated_count % PROGRESS_UPDATE_INTERVAL == 0
                or generated_count == validated_quantity
            ):
                progress_callback(
                    generated_count,
                    validated_quantity,
                )

    return generated_numbers


def verify_no_duplicates(unique_numbers):
    """
    Comprueba que no existan duplicados en el lote completo.

    Esta comprobación se realiza antes de dividir los códigos
    en diferentes archivos.

    Args:
        unique_numbers (list[str]): códigos generados.

    Raises:
        RuntimeError: si se encuentra algún código duplicado.
    """
    total_codes = len(unique_numbers)
    total_unique_codes = len(set(unique_numbers))

    if total_codes != total_unique_codes:
        duplicated_count = total_codes - total_unique_codes

        raise RuntimeError(
            "Se han detectado códigos duplicados.\n\n"
            f"Duplicados encontrados: {duplicated_count}\n\n"
            "No se ha generado ningún archivo."
        )


# ============================================================
# DIVISIÓN EN BLOQUES
# ============================================================

def split_into_batches(
    unique_numbers,
    batch_size=MAX_CODES_PER_FILE,
):
    """
    Divide los códigos en bloques de tamaño máximo configurable.

    Args:
        unique_numbers (list[str]): lista completa.
        batch_size (int): máximo de códigos por bloque.

    Returns:
        list[list[str]]: bloques de códigos.
    """
    if batch_size <= 0:
        raise ValueError(
            "El tamaño máximo de cada bloque debe ser mayor que 0."
        )

    return [
        unique_numbers[index:index + batch_size]
        for index in range(
            0,
            len(unique_numbers),
            batch_size,
        )
    ]


# ============================================================
# NOMBRES DE ARCHIVO
# ============================================================

def build_base_file_name(prefix, quantity):
    """
    Construye el nombre base del lote.

    Ejemplo:
        codigos_FCB_10100_20260918_094530

    Args: 
        prefix (str): prefijo del lote.
        quantity (int): cantidad total solicitada.

    Returns:
        str: nombre base.
    """
    generation_date = datetime.datetime.now().strftime(
        "%Y%m%d_%H%M%S"
    )

    return (
        f"codigos_{prefix}_{quantity}_{generation_date}"
    )


def build_part_file_name(
    base_name,
    part_number,
    total_parts,
    extension,
):
    """
    Construye el nombre de cada archivo.

    Ejemplo:
        codigos_FCB_10100_20260918_094530_parte_01_de_02.xlsx

    Args:
        base_name (str): nombre base del lote.
        part_number (int): número de archivo.
        total_parts (int): total de archivos.
        extension (str): extensión sin punto.

    Returns:
        str: nombre completo del archivo.
    """
    number_width = max(2, len(str(total_parts)))

    formatted_part = str(part_number).zfill(number_width)
    formatted_total = str(total_parts).zfill(number_width)

    return (
        f"{base_name}_parte_{formatted_part}"
        f"_de_{formatted_total}.{extension}"
    )


# ============================================================
# EXPORTACIÓN A TXT
# ============================================================

def save_as_txt_batches(
    unique_numbers,
    output_directory,
    base_name,
):
    """
    Guarda los códigos en uno o varios archivos TXT.

    Cada archivo tendrá un máximo de 5.050 códigos.

    Args:
        unique_numbers (list[str]): códigos generados.
        output_directory (str | Path): carpeta de destino.
        base_name (str): nombre base del lote.

    Returns:
        listarchivos creados.
    """
    destination = Path(output_directory)
    destination.mkdir(parents=True, exist_ok=True)

    batches = split_into_batches(unique_numbers)
    total_parts = len(batches)
    generated_files = []

    for part_number, batch in enumerate(batches, start=1):
        file_name = build_part_file_name(
            base_name=base_name,
            part_number=part_number,
            total_parts=total_parts,
            extension="txt",
        )

        file_path = destination / file_name

        with file_path.open(
            mode="w",
            encoding="utf-8",
        ) as output_file:
            output_file.write("\n".join(batch))
            output_file.write("\n")

        generated_files.append(file_path)

    return generated_files


# ============================================================
# EXPORTACIÓN A EXCEL
# ============================================================

def save_as_excel_batches(
    unique_numbers,
    output_directory,
    base_name,
):
    """
    Guarda los códigos en uno o varios archivos Excel.

    Cada archivo tendrá un máximo de 5.050 códigos,
    sin contar la fila de encabezado.

    Args:
        unique_numbers (list[str]): códigos generados.
        output_directory (str | Path): carpeta de destino.
        base_name (str): nombre base del lote.

    Returns:
        listarchivos creados.

    Raises:
        ImportError: si openpyxl no está instalado.
    """
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font
    except ImportError as error:
        raise ImportError(
            "No se ha podido importar openpyxl.\n\n"
            "Instálalo en el mismo intérprete de Python con:\n"
            "python -m pip install openpyxl"
        ) from error

    destination = Path(output_directory)
    destination.mkdir(parents=True, exist_ok=True)

    batches = split_into_batches(unique_numbers)
    total_parts = len(batches)
    generated_files = []

    for part_number, batch in enumerate(batches, start=1):
        file_name = build_part_file_name(
            base_name=base_name,
            part_number=part_number,
            total_parts=total_parts,
            extension="xlsx",
        )

        file_path = destination / file_name

        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Códigos únicos"

        # Encabezado
        worksheet["A1"] = "CODIGO"
        worksheet["A1"].font = Font(bold=True)
        worksheet.column_dimensions["A"].width = 35
        worksheet.freeze_panes = "A2"

        # Contenido
        for row_number, unique_number in enumerate(
            batch,
            start=2,
        ):
            cell = worksheet.cell(
                row=row_number,
                column=1,
                value=unique_number,
            )

            # Los códigos se almacenan como texto para evitar que
            # Excel modifique los valores largos.
            cell.number_format = "@"

        workbook.save(file_path)
        workbook.close()

        generated_files.append(file_path)

    return generated_files


# ============================================================
# INTERFAZ GRÁFICA
# ============================================================

class UniqueCodeGeneratorApp:
    """Interfaz gráfica del generador de códigos únicos."""

    def __init__(self, root):
        self.root = root

        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.resizable(False, False)

        self.prefix_variable = tk.StringVar()
        self.quantity_variable = tk.StringVar(value="5050")
        self.format_variable = tk.StringVar(value="txt")

        self.status_variable = tk.StringVar(
            value="Introduce el prefijo y la cantidad."
        )

        self.files_variable = tk.StringVar(
            value="Archivos previstos: 1"
        )

        self.generation_in_progress = False
        self.message_queue = queue.Queue()

        self.configure_styles()
        self.create_widgets()
        self.configure_events()

        # Revisa periódicamente los mensajes enviados
        # por el hilo de generación.
        self.root.after(50, self.process_thread_messages)

    def configure_styles(self):
        """Configura los estilos visuales."""
        style = ttk.Style()

        available_themes = style.theme_names()

        if "vista" in available_themes:
            style.theme_use("vista")
        elif "clam" in available_themes:
            style.theme_use("clam")

        style.configure(
            "Title.TLabel",
            font=("Segoe UI", 17, "bold"),
        )

        style.configure(
            "Subtitle.TLabel",
            font=("Segoe UI", 10),
            foreground="#555555",
        )

        style.configure(
            "Information.TLabel",
            font=("Segoe UI", 9),
            foreground="#444444",
        )

        style.configure(
            "Generate.TButton",
            font=("Segoe UI", 11, "bold"),
            padding=9,
        )

    def create_widgets(self):
        """Crea todos los elementos de la ventana."""
        main_frame = ttk.Frame(
            self.root,
            padding=25,
        )
        main_frame.pack(
            fill="both",
            expand=True,
        )

        title_label = ttk.Label(
            main_frame,
            text="Generador de códigos únicos",
            style="Title.TLabel",
        )
        title_label.pack(pady=(0, 5))

        subtitle_label = ttk.Label(
            main_frame,
            text=(
                "PREFIJO + año, mes, día, hora, minutos, "
                "segundos y milisegundos"
            ),
            style="Subtitle.TLabel",
            wraplength=510,
            justify="center",
        )
        subtitle_label.pack(pady=(0, 25))

        form_frame = ttk.Frame(main_frame)
        form_frame.pack(fill="x")

        prefix_label = ttk.Label(
            form_frame,
            text="Prefijo:",
        )
        prefix_label.grid(
            row=0,
            column=0,
            sticky="w",
            padx=(0, 15),
            pady=8,
        )

        self.prefix_entry = ttk.Entry(
            form_frame,
            textvariable=self.prefix_variable,
            width=35,
        )
        self.prefix_entry.grid(
            row=0,
            column=1,
            sticky="ew",
            pady=8,
        )

        quantity_label = ttk.Label(
            form_frame,
            text="Cantidad:",
        )
        quantity_label.grid(
            row=1,
            column=0,
            sticky="w",
            padx=(0, 15),
            pady=8,
        )

        self.quantity_entry = ttk.Entry(
            form_frame,
            textvariable=self.quantity_variable,
            width=35,
        )
        self.quantity_entry.grid(
            row=1,
            column=1,
            sticky="ew",
            pady=8,
        )

        format_label = ttk.Label(
            form_frame,
            text="Formato:",
        )
        format_label.grid(
            row=2,
            column=0,
            sticky="w",
            padx=(0, 15),
            pady=8,
        )

        format_frame = ttk.Frame(form_frame)
        format_frame.grid(
            row=2,
            column=1,
            sticky="w",
            pady=8,
        )

        self.txt_radio_button = ttk.Radiobutton(
            format_frame,
            text="Archivo TXT",
            variable=self.format_variable,
            value="txt",
        )
        self.txt_radio_button.pack(
            side="left",
            padx=(0, 20),
        )

        self.excel_radio_button = ttk.Radiobutton(
            format_frame,
            text="Archivo Excel",
            variable=self.format_variable,
            value="xlsx",
        )
        self.excel_radio_button.pack(side="left")

        form_frame.columnconfigure(1, weight=1)

        limit_label = ttk.Label(
            main_frame,
            text=(
                "Cada archivo contendrá un máximo de "
                f"{MAX_CODES_PER_FILE:,} códigos."
            ).replace(",", "."),
            style="Information.TLabel",
        )
        limit_label.pack(pady=(15, 3))

        files_label = ttk.Label(
            main_frame,
            textvariable=self.files_variable,
            style="Information.TLabel",
        )
        files_label.pack(pady=(0, 12))

        self.progress_bar = ttk.Progressbar(
            main_frame,
            orient="horizontal",
            mode="determinate",
            maximum=100,
        )
        self.progress_bar.pack(
            fill="x",
            pady=(8, 8),
        )

        self.status_label = ttk.Label(
            main_frame,
            textvariable=self.status_variable,
            anchor="center",
            wraplength=510,
        )
        self.status_label.pack(
            fill="x",
            pady=(0, 20),
        )

        self.generate_button = ttk.Button(
            main_frame,
            text="Generar códigos",
            command=self.start_generation,
            style="Generate.TButton",
        )
        self.generate_button.pack(fill="x")

        self.prefix_entry.focus_set()

    def configure_events(self):
        """Configura los eventos de la interfaz."""
        self.quantity_variable.trace_add(
            "write",
            self.update_expected_file_count,
        )

        self.root.bind(
            "<Return>",
            lambda event: self.start_generation(),
        )

        self.root.protocol(
            "WM_DELETE_WINDOW",
            self.close_application,
        )

    def update_expected_file_count(self, *args):
        """Actualiza el número previsto de archivos."""
        quantity_text = self.quantity_variable.get().strip()

        try:
            quantity = int(quantity_text)

            if quantity <= 0:
                raise ValueError

            total_files = math.ceil(
                quantity / MAX_CODES_PER_FILE
            )

            self.files_variable.set(
                f"Archivos previstos: {total_files}"
            )

        except ValueError:
            self.files_variable.set(
                "Archivos previstos: -"
            )

    def get_form_data(self):
        """
        Obtiene y valida los valores del formulario.

        Returns:
            tuple[str, int, str]: prefijo, cantidad y formato.
        """
        prefix = validate_prefix(
            self.prefix_variable.get()
        )

        quantity_text = self.quantity_variable.get().strip()

        if not quantity_text:
            raise ValueError(
                "Debes introducir una cantidad."
            )

        try:
            quantity = int(quantity_text)
        except ValueError as error:
            raise ValueError(
                "La cantidad debe ser un número entero."
            ) from error

        quantity = validate_quantity(quantity)

        output_format = self.format_variable.get()

        if output_format not in ("txt", "xlsx"):
            raise ValueError(
                "Debes seleccionar un formato de salida."
            )

        return prefix, quantity, output_format

    def ask_output_directory(self):
        """Solicita la carpeta de destino."""
        return filedialog.askdirectory(
            title="Selecciona la carpeta de destino"
        )

    def start_generation(self):
        """Inicia la generación de códigos."""
        if self.generation_in_progress:
            return

        try:
            prefix, quantity, output_format = (
                self.get_form_data()
            )
        except ValueError as error:
            messagebox.showerror(
                "Datos incorrectos",
                str(error),
            )
            return

        output_directory = self.ask_output_directory()

        if not output_directory:
            self.status_variable.set(
                "Generación cancelada."
            )
            return

        total_files = math.ceil(
            quantity / MAX_CODES_PER_FILE
        )

        self.generation_in_progress = True
        self.set_interface_enabled(False)

        self.progress_bar["value"] = 0

        self.status_variable.set(
            f"Generando 0 de {quantity} códigos. "
            f"Se crearán {total_files} archivo(s)."
        )

        generation_thread = threading.Thread(
            target=self.generate_and_save,
            args=(
                prefix,
                quantity,
                output_format,
                output_directory,
            ),
            daemon=True,
        )

        generation_thread.start()

    def notify_progress(self, current, total):
        """
        Envía el progreso desde el hilo secundario.

        Args:
            current (int): códigos generados.
            total (int): cantidad total.
        """
        self.message_queue.put(
            (
                "progress",
                {
                    "current": current,
                    "total": total,
                },
            )
        )

    def generate_and_save(
        self,
        prefix,
        quantity,
        output_format,
        output_directory,
    ):
        """
        Genera, comprueba y guarda los códigos.

        Este método se ejecuta en un hilo secundario.
        """
        try:
            generated_numbers = generate_unique_numbers(
                quantity=quantity,
                prefix=prefix,
                progress_callback=self.notify_progress,
            )

            # Se comprueba el lote completo antes de dividirlo.
            verify_no_duplicates(generated_numbers)

            base_name = build_base_file_name(
                prefix=prefix,
                quantity=quantity,
            )

            if output_format == "txt":
                generated_files = save_as_txt_batches(
                    unique_numbers=generated_numbers,
                    output_directory=output_directory,
                    base_name=base_name,
                )
            else:
                generated_files = save_as_excel_batches(
                    unique_numbers=generated_numbers,
                    output_directory=output_directory,
                    base_name=base_name,
                )

            self.message_queue.put(
                (
                    "completed",
                    {
                        "generated_numbers": generated_numbers,
                        "generated_files": generated_files,
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
        """
        Procesa mensajes del hilo secundario de forma segura.
        """
        try:
            while True:
                message_type, message_data = (
                    self.message_queue.get_nowait()
                )

                if message_type == "progress":
                    self.update_progress(
                        current=message_data["current"],
                        total=message_data["total"],
                    )

                elif message_type == "completed":
                    self.generation_completed(
                        generated_numbers=message_data[
                            "generated_numbers"
                        ],
                        generated_files=message_data[
                            "generated_files"
                        ],
                    )

                elif message_type == "error":
                    self.generation_failed(
                        message_data["message"]
                    )

        except queue.Empty:
            pass

        if self.root.winfo_exists():
            self.root.after(
                50,
                self.process_thread_messages,
            )

    def update_progress(self, current, total):
        """Actualiza la barra de progreso."""
        percentage = (current / total) * 100

        self.progress_bar["value"] = percentage

        self.status_variable.set(
            f"Generando {current} de {total} códigos..."
        )

    def generation_completed(
        self,
        generated_numbers,
        generated_files,
    ):
        """Muestra el resumen final."""
        self.generation_in_progress = False
        self.set_interface_enabled(True)

        self.progress_bar["value"] = 100

        total_codes = len(generated_numbers)
        total_files = len(generated_files)

        first_code = generated_numbers[0]
        last_code = generated_numbers[-1]

        file_names = "\n".join(
            file_path.name
            for file_path in generated_files
        )

        destination_directory = (
            generated_files[0].parent
        )

        self.status_variable.set(
            f"Completado: {total_codes} códigos "
            f"en {total_files} archivo(s)."
        )

        messagebox.showinfo(
            "Generación completada",
            (
                f"Códigos generados: {total_codes}\n"
                f"Archivos creados: {total_files}\n\n"
                f"Primer código:\n{first_code}\n\n"
                f"Último código:\n{last_code}\n\n"
                f"Carpeta de destino:\n"
                f"{destination_directory}\n\n"
                f"Archivos generados:\n{file_names}"
            ),
        )

    def generation_failed(self, error_message):
        """Muestra los errores de generación o guardado."""
        self.generation_in_progress = False
        self.set_interface_enabled(True)

        self.status_variable.set(
            "Se produjo un error durante la generación."
        )

        messagebox.showerror(
            "Error",
            error_message,
        )

    def set_interface_enabled(self, enabled):
        """Activa o desactiva los controles."""
        state = "normal" if enabled else "disabled"

        self.prefix_entry.configure(state=state)
        self.quantity_entry.configure(state=state)
        self.txt_radio_button.configure(state=state)
        self.excel_radio_button.configure(state=state)
        self.generate_button.configure(state=state)

    def close_application(self):
        """Gestiona el cierre de la aplicación."""
        if self.generation_in_progress:
            should_close = messagebox.askyesno(
                "Generación en curso",
                (
                    "Hay una generación en curso.\n\n"
                    "¿Quieres cerrar la aplicación?"
                ),
            )

            if not should_close:
                return

        self.root.destroy()


# ============================================================
# INICIO DE LA APLICACIÓN
# ============================================================

def main():
    """Inicia la interfaz gráfica."""
    root = tk.Tk()
    UniqueCodeGeneratorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()