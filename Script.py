"""
Generador de códigos únicos para productos Stahls.

Estructura del código:
    PREFIJO + yyyyMMddHHmmssSSS

Ejemplo:
    FCB20260918083045123

Donde:
- PREFIJO: introducido por el usuario.
- yyyy: año.
- MM: mes.
- dd: día.
- HH: hora.
- mm: minutos.
- ss: segundos.
- SSS: milisegundos.

El programa genera como máximo un código por milisegundo.

Los códigos pueden exportarse a:
- Archivo TXT.
- Archivo Excel XLSX.
"""

import datetime
import threading
import time
import tkinter as tk

from pathlib import Path
from tkinter import filedialog, messagebox, ttk


# Configuración general
WINDOW_TITLE = "Generador de códigos únicos Stahls"
WINDOW_SIZE = "540x460"
MAX_PREFIX_LENGTH = 20


def validate_quantity(quantity):
    """
    Valida que la cantidad solicitada sea un entero positivo.

    Args:
        quantity (int): cantidad que se desea validar.

    Returns:
        int: cantidad validada.

    Raises:
        ValueError: si la cantidad no es un entero positivo.
    """
    if not isinstance(quantity, int):
        raise ValueError("La cantidad debe ser un número entero.")

    if quantity <= 0:
        raise ValueError("La cantidad debe ser mayor que 0.")

    return quantity


def validate_prefix(prefix):
    """
    Valida y normaliza el prefijo.

    Args:
        prefix (str): prefijo introducido por el usuario.

    Returns:
        str: prefijo validado y convertido a mayúsculas.

    Raises:
        ValueError: si el prefijo está vacío, supera la longitud máxima
        o contiene caracteres no permitidos.
    """
    if not isinstance(prefix, str):
        raise ValueError("El prefijo debe ser texto.")

    prefix = prefix.strip().upper()

    if not prefix:
        raise ValueError("El prefijo no puede estar vacío.")

    if len(prefix) > MAX_PREFIX_LENGTH:
        raise ValueError(
            f"El prefijo no puede superar los "
            f"{MAX_PREFIX_LENGTH} caracteres."
        )

    if not prefix.isalnum():
        raise ValueError(
            "El prefijo solo puede contener letras y números."
        )

    return prefix


def generate_unique_number(prefix, last_timestamp=None):
    """
    Genera un código utilizando el prefijo y el instante actual.

    El programa espera hasta que el reloj del sistema alcance un
    milisegundo diferente al utilizado en el código anterior.

    Args:
        prefix (str): prefijo validado.
        last_timestamp (str | None): timestamp del código anterior,
            sin incluir el prefijo.

    Returns:
        tuple[str, str]: código completo y timestamp utilizado.
    """
    while True:
        current_time = datetime.datetime.now()

        date_and_time = current_time.strftime("%Y%m%d%H%M%S")
        milliseconds = current_time.microsecond // 1000
        timestamp = f"{date_and_time}{milliseconds:03d}"

        if last_timestamp is None or timestamp > last_timestamp:
            complete_code = f"{prefix}{timestamp}"
            return complete_code, timestamp

        # Espera breve antes de volver a consultar el reloj.
        time.sleep(0.0001)


def generate_unique_numbers(quantity, prefix, progress_callback=None):
    """
    Genera la cantidad solicitada de códigos únicos.

    Args:
        quantity (int): cantidad de códigos que se desea generar.
        prefix (str): prefijo de los códigos.
        progress_callback (callable | None): función utilizada para
            actualizar el progreso de la interfaz.

    Returns:
        listlista de códigos generados.
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

        if progress_callback is not None:
            progress_callback(index + 1, validated_quantity)

    return generated_numbers


def save_as_txt(unique_numbers, file_path):
    """
    Guarda los códigos en un archivo TXT.

    Cada código se almacena en una línea diferente.

    Args:
        unique_numbers (list[str]): códigos generados.
        file_path (str | Path): ruta del archivo de salida.
    """
    output_path = Path(file_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with output_path.open("w", encoding="utf-8") as file:
        file.write("\n".join(unique_numbers) + "\n")


def save_as_excel(unique_numbers, file_path):
    """
    Guarda los códigos en un archivo Excel.

    Args:
        unique_numbers (list[str]): códigos generados.
        file_path (str | Path): ruta del archivo de salida.

    Raises:
        ImportError: si openpyxl no está instalado.
    """
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font
    except ImportError as error:
        raise ImportError(
            "Para exportar a Excel debes instalar openpyxl con:\n"
            "pip install openpyxl"
        ) from error

    output_path = Path(file_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Códigos únicos"

    worksheet["A1"] = "CODIGO"
    worksheet["A1"].font = Font(bold=True)
    worksheet.column_dimensions["A"].width = 35

    for row_number, unique_number in enumerate(
        unique_numbers,
        start=2,
    ):
        cell = worksheet.cell(
            row=row_number,
            column=1,
            value=unique_number,
        )

        # Fuerza el almacenamiento como texto.
        cell.number_format = "@"

    workbook.save(output_path)


class UniqueCodeGeneratorApp:
    """Interfaz gráfica del generador de códigos únicos."""

    def __init__(self, root):
        self.root = root
        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.resizable(False, False)

        self.prefix_variable = tk.StringVar()
        self.quantity_variable = tk.StringVar(value="5000")
        self.format_variable = tk.StringVar(value="txt")
        self.status_variable = tk.StringVar(
            value="Introduce los datos para comenzar."
        )

        self.generated_numbers = []
        self.generation_in_progress = False

        self.configure_styles()
        self.create_widgets()

    def configure_styles(self):
        """Configura los estilos visuales de la aplicación."""
        style = ttk.Style()

        try:
            style.theme_use("vista")
        except tk.TclError:
            pass

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
            "Generate.TButton",
            font=("Segoe UI", 11, "bold"),
            padding=8,
        )

    def create_widgets(self):
        """Crea todos los controles de la ventana."""
        main_frame = ttk.Frame(self.root, padding=25)
        main_frame.pack(fill="both", expand=True)

        title_label = ttk.Label(
            main_frame,
            text="Generador de códigos únicos",
            style="Title.TLabel",
        )
        title_label.pack(pady=(0, 5))

        subtitle_label = ttk.Label(
            main_frame,
            text=(
                "Estructura: PREFIJO + "
                "año, mes, día, hora, minutos, segundos y milisegundos"
            ),
            style="Subtitle.TLabel",
            wraplength=470,
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
            width=30,
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
            width=30,
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

        self.progress_bar = ttk.Progressbar(
            main_frame,
            orient="horizontal",
            mode="determinate",
            maximum=100,
        )
        self.progress_bar.pack(
            fill="x",
            pady=(25, 8),
        )

        self.status_label = ttk.Label(
            main_frame,
            textvariable=self.status_variable,
            anchor="center",
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

    def get_form_data(self):
        """
        Lee y valida los valores introducidos en el formulario.

        Returns:
            tuple[str, int, str]: prefijo, cantidad y formato.
        """
        prefix = validate_prefix(
            self.prefix_variable.get()
        )

        quantity_text = self.quantity_variable.get().strip()

        if not quantity_text:
            raise ValueError("Debes introducir una cantidad.")

        try:
            quantity = int(quantity_text)
        except ValueError as error:
            raise ValueError(
                "La cantidad debe ser un número entero."
            ) from error

        quantity = validate_quantity(quantity)
        output_format = self.format_variable.get()

        return prefix, quantity, output_format

    def start_generation(self):
        """Valida el formulario e inicia la generación."""
        if self.generation_in_progress:
            return

        try:
            prefix, quantity, output_format = self.get_form_data()
        except ValueError as error:
            messagebox.showerror(
                "Datos incorrectos",
                str(error),
            )
            return

        default_file_name = self.build_default_file_name(
            prefix=prefix,
            quantity=quantity,
            output_format=output_format,
        )

        file_path = self.ask_output_path(
            output_format=output_format,
            default_file_name=default_file_name,
        )

        if not file_path:
            self.status_variable.set(
                "Generación cancelada."
            )
            return

        self.generation_in_progress = True
        self.set_interface_enabled(False)

        self.progress_bar["value"] = 0
        self.status_variable.set(
            f"Generando 0 de {quantity} códigos..."
        )

        generation_thread = threading.Thread(
            target=self.generate_and_save,
            args=(
                prefix,
                quantity,
                output_format,
                file_path,
            ),
            daemon=True,
        )

        generation_thread.start()

    def generate_and_save(
        self,
        prefix,
        quantity,
        output_format,
        file_path,
    ):
        """
        Genera y guarda los códigos en un hilo secundario.

        Args:
            prefix (str): prefijo validado.
            quantity (int): cantidad validada.
            output_format (str): formato TXT o XLSX.
            file_path (str): archivo seleccionado.
        """
        try:
            generated_numbers = generate_unique_numbers(
                quantity=quantity,
                prefix=prefix,
                progress_callback=self.schedule_progress_update,
            )

            if output_format == "txt":
                save_as_txt(
                    generated_numbers,
                    file_path,
                )
            else:
                save_as_excel(
                    generated_numbers,
                    file_path,
                )

            self.root.after(
                0,
                self.generation_completed,
                generated_numbers,
                file_path,
            )

        except Exception as error:
            self.root.after(
                0,
                self.generation_failed,
                str(error),
            )

    def schedule_progress_update(self, current, total):
        """
        Solicita a Tkinter la actualización segura del progreso.

        Args:
            current (int): cantidad generada.
            total (int): cantidad total solicitada.
        """
        self.root.after(
            0,
            self.update_progress,
            current,
            total,
        )

    def update_progress(self, current, total):
        """Actualiza la barra y el mensaje de progreso."""
        percentage = (current / total) * 100
        self.progress_bar["value"] = percentage

        self.status_variable.set(
            f"Generando {current} de {total} códigos..."
        )

    def generation_completed(
        self,
        generated_numbers,
        file_path,
    ):
        """Muestra el resultado de una generación correcta."""
        self.generated_numbers = generated_numbers
        self.generation_in_progress = False
        self.set_interface_enabled(True)

        self.progress_bar["value"] = 100

        first_code = generated_numbers[0]
        last_code = generated_numbers[-1]

        self.status_variable.set(
            f"Completado: {len(generated_numbers)} códigos generados."
        )

        messagebox.showinfo(
            "Generación completada",
            (
                f"Se han generado "
                f"{len(generated_numbers)} códigos.\n\n"
                f"Primer código:\n{first_code}\n\n"
                f"Último código:\n{last_code}\n\n"
                f"Archivo guardado en:\n{file_path}"
            ),
        )

    def generation_failed(self, error_message):
        """Muestra un error producido durante la generación."""
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
        """
        Activa o desactiva los controles durante la generación.

        Args:
            enabled (bool): estado de los controles.
        """
        state = "normal" if enabled else "disabled"

        self.prefix_entry.configure(state=state)
        self.quantity_entry.configure(state=state)
        self.txt_radio_button.configure(state=state)
        self.excel_radio_button.configure(state=state)
        self.generate_button.configure(state=state)

    @staticmethod
    def build_default_file_name(
        prefix,
        quantity,
        output_format,
    ):
        """Construye el nombre sugerido para el archivo."""
        current_time = datetime.datetime.now()
        date_text = current_time.strftime("%Y%m%d_%H%M%S")

        return (
            f"codigos_{prefix}_{quantity}_"
            f"{date_text}.{output_format}"
        )

    def ask_output_path(
        self,
        output_format,
        default_file_name,
    ):
        """Abre la ventana para seleccionar el archivo de salida."""
        if output_format == "txt":
            return filedialog.asksaveasfilename(
                title="Guardar códigos únicos",
                initialfile=default_file_name,
                defaultextension=".txt",
                filetypes=[
                    ("Archivo de texto", "*.txt"),
                    ("Todos los archivos", "*.*"),
                ],
            )

        return filedialog.asksaveasfilename(
            title="Guardar códigos únicos",
            initialfile=default_file_name,
            defaultextension=".xlsx",
            filetypes=[
                ("Archivo Excel", "*.xlsx"),
                ("Todos los archivos", "*.*"),
            ],
        )


def main():
    """Inicia la aplicación gráfica."""
    root = tk.Tk()
    UniqueCodeGeneratorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()