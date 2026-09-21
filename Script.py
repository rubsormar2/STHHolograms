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

def validar_cantidad(cantidad):
    """
    Valida que la cantidad solicitada sea un número entero positivo.

    Args:
        cantidad (int): cantidad solicitada.

    Returns:
        int: cantidad validada.

    Raises:
        ValueError: si la cantidad no es válida.
    """
    if not isinstance(cantidad, int):
        raise ValueError("La cantidad debe ser un número entero.")

    if cantidad <= 0:
        raise ValueError("La cantidad debe ser mayor que 0.")

    return cantidad


def validar_prefijo(prefijo):
    """
    Valida y normaliza el prefijo.

    El prefijo:
    - No puede estar vacío.
    - Se convierte automáticamente a mayúsculas.
    - Solo puede contener letras y números.
    - No puede superar la longitud máxima configurada.

    Args:
        prefijo (str): prefijo escrito por el usuario.

    Returns:
        str: prefijo validado.

    Raises:
        ValueError: si el prefijo no es válido.
    """
    if not isinstance(prefijo, str):
        raise ValueError("El prefijo debe ser texto.")

    validated_prefix = prefijo.strip().upper()

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

def obtener_timestamp_actual():
    """
    Obtiene el timestamp actual con precisión de milisegundos.

    Estructura:
        yyyyMMddHHmmssSSS

    Ejemplo:
        20260918094530123

    Returns:
        str: timestamp de 17 dígitos.
    """
    tiempo_actual = datetime.datetime.now()

    fecha_hora = tiempo_actual.strftime("%Y%m%d%H%M%S")
    milisegundos = tiempo_actual.microsecond // 1000

    return f"{fecha_hora}{milisegundos:03d}"


def generar_numero_unico(prefijo, ultimo_timestamp=None):
    """
    Genera un código único utilizando el prefijo y el reloj del sistema.

    Si el milisegundo actual ya se utilizó, el programa espera hasta
    que el reloj alcance un milisegundo posterior.

    Args:
        prefijo (str): prefijo validado.
        ultimo_timestamp (str | None): último timestamp utilizado.

    Returns:
        tuple[str, str]:
            - Código completo.
            - Timestamp utilizado, sin el prefijo.
    """
    while True:
        timestamp_actual = obtener_timestamp_actual()

        if (
            ultimo_timestamp is None
            or timestamp_actual > ultimo_timestamp
        ):
            complete_code = f"{prefijo}{timestamp_actual}"

            return complete_code, timestamp_actual

        # Se vuelve a consultar el reloj hasta que cambie
        # el milisegundo.
        time.sleep(0.0001)


def generar_numeros_unicos(
    cantidad,
    prefijo,
    progress_callback=None,
):
    """
    Genera una lista completa de códigos únicos.

    La variable ultimo_timestamp se mantiene durante toda la generación.
    Por tanto, la posterior división en archivos no reinicia el
    timestamp ni afecta a la unicidad.

    Args:
        cantidad (int): cantidad total de códigos.
        prefijo (str): prefijo solicitado.
        progress_callback (callable | None): función para comunicar
            el progreso.

    Returns:
        listcódigos únicos generados.
    """
    validated_quantity = validar_cantidad(cantidad)
    validated_prefix = validar_prefijo(prefijo)

    numeros_generados = []
    ultimo_timestamp = None

    for index in range(validated_quantity):
        unique_number, timestamp_actual = generar_numero_unico(
            prefijo=validated_prefix,
            ultimo_timestamp=ultimo_timestamp,
        )

        numeros_generados.append(unique_number)
        ultimo_timestamp = timestamp_actual

        cantidad_generada = index + 1

        # Evita enviar miles de actualizaciones innecesarias
        # a la interfaz gráfica.
        if progress_callback is not None:
            if (
                cantidad_generada % PROGRESS_UPDATE_INTERVAL == 0
                or cantidad_generada == validated_quantity
            ):
                progress_callback(
                    cantidad_generada,
                    validated_quantity,
                )

    return numeros_generados


def verificar_sin_duplicados(unique_numbers):
    """
    Comprueba que no existan duplicados en el lote completo.

    Esta comprobación se realiza antes de dividir los códigos
    en diferentes archivos.

    Args:
        unique_numbers (list[str]): códigos generados.

    Raises:
        RuntimeError: si se encuentra algún código duplicado.
    """
    total_codigos = len(unique_numbers)
    total_codigos_unicos = len(set(unique_numbers))

    if total_codigos != total_codigos_unicos:
        cantidad_duplicados = total_codigos - total_codigos_unicos

        raise RuntimeError(
            "Se han detectado códigos duplicados.\n\n"
            f"Duplicados encontrados: {cantidad_duplicados}\n\n"
            "No se ha generado ningún archivo."
        )


# ============================================================
# DIVISIÓN EN BLOQUES
# ============================================================

def dividir_en_lotes(
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

def construir_nombre_base_archivo(prefijo, cantidad):
    """
    Construye el nombre base del lote.

    Ejemplo:
        codigos_FCB_10100_20260918_094530

    Args: 
        prefijo (str): prefijo del lote.
        cantidad (int): cantidad total solicitada.

    Returns:
        str: nombre base.
    """
    generation_date = datetime.datetime.now().strftime(
        "%Y%m%d_%H%M%S"
    )

    return (
        f"codigos_{prefijo}_{cantidad}_{generation_date}"
    )


def construir_nombre_archivo_parte(
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

def guardar_lotes_txt(
    unique_numbers,
    directorio_salida,
    base_name,
):
    """
    Guarda los códigos en uno o varios archivos TXT.

    Cada archivo tendrá un máximo de 5.050 códigos.

    Args:
        unique_numbers (list[str]): códigos generados.
        directorio_salida (str | Path): carpeta de destino.
        base_name (str): nombre base del lote.

    Returns:
        listarchivos creados.
    """
    destination = Path(directorio_salida)
    destination.mkdir(parents=True, exist_ok=True)

    batches = dividir_en_lotes(unique_numbers)
    total_parts = len(batches)
    archivos_generados = []

    for part_number, batch in enumerate(batches, start=1):
        nombre_archivo = construir_nombre_archivo_parte(
            base_name=base_name,
            part_number=part_number,
            total_parts=total_parts,
            extension="txt",
        )

        ruta_archivo = destination / nombre_archivo

        with ruta_archivo.open(
            mode="w",
            encoding="utf-8",
        ) as archivo_salida:
            archivo_salida.write("\n".join(batch))
            archivo_salida.write("\n")

        archivos_generados.append(ruta_archivo)

    return archivos_generados


# ============================================================
# EXPORTACIÓN A EXCEL
# ============================================================

def guardar_lotes_excel(
    unique_numbers,
    directorio_salida,
    base_name,
):
    """
    Guarda los códigos en uno o varios archivos Excel.

    Cada archivo tendrá un máximo de 5.050 códigos,
    sin contar la fila de encabezado.

    Args:
        unique_numbers (list[str]): códigos generados.
        directorio_salida (str | Path): carpeta de destino.
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

    destination = Path(directorio_salida)
    destination.mkdir(parents=True, exist_ok=True)

    batches = dividir_en_lotes(unique_numbers)
    total_parts = len(batches)
    archivos_generados = []

    for part_number, batch in enumerate(batches, start=1):
        nombre_archivo = construir_nombre_archivo_parte(
            base_name=base_name,
            part_number=part_number,
            total_parts=total_parts,
            extension="xlsx",
        )

        ruta_archivo = destination / nombre_archivo

        libro_trabajo = Workbook()
        hoja_trabajo = libro_trabajo.active
        hoja_trabajo.title = "Códigos únicos"

        # Encabezado
        hoja_trabajo["A1"] = "CODIGO"
        hoja_trabajo["A1"].font = Font(bold=True)
        hoja_trabajo.column_dimensions["A"].width = 35
        hoja_trabajo.freeze_panes = "A2"

        # Contenido
        for numero_fila, unique_number in enumerate(
            batch,
            start=2,
        ):
            cell = hoja_trabajo.cell(
                row=numero_fila,
                column=1,
                value=unique_number,
            )

            # Los códigos se almacenan como texto para evitar que
            # Excel modifique los valores largos.
            cell.number_format = "@"

        libro_trabajo.save(ruta_archivo)
        libro_trabajo.close()

        archivos_generados.append(ruta_archivo)

    return archivos_generados


# ============================================================
# INTERFAZ GRÁFICA
# ============================================================

class GeneradorCodigosUnicosApp:
    """Interfaz gráfica del generador de códigos únicos."""

    def __init__(self, root):
        self.root = root

        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.resizable(False, False)

        self.variable_prefijo = tk.StringVar()
        self.variable_cantidad = tk.StringVar(value="5050")
        self.variable_formato = tk.StringVar(value="txt")

        self.variable_estado = tk.StringVar(
            value="Introduce el prefijo y la cantidad."
        )

        self.variable_archivos = tk.StringVar(
            value="Archivos previstos: 1"
        )

        self.generacion_en_progreso = False
        self.cola_mensajes = queue.Queue()

        self.configurar_estilos()
        self.crear_widgets()
        self.configurar_eventos()

        # Revisa periódicamente los mensajes enviados
        # por el hilo de generación.
        self.root.after(50, self.procesar_mensajes_hilo)

    def configurar_estilos(self):
        """Configura los estilos visuales."""
        style = ttk.Style()

        temas_disponibles = style.theme_names()

        if "vista" in temas_disponibles:
            style.theme_use("vista")
        elif "clam" in temas_disponibles:
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

    def crear_widgets(self):
        """Crea todos los elementos de la ventana."""
        marco_principal = ttk.Frame(
            self.root,
            padding=25,
        )
        marco_principal.pack(
            fill="both",
            expand=True,
        )

        etiqueta_titulo = ttk.Label(
            marco_principal,
            text="Generador de códigos únicos",
            style="Title.TLabel",
        )
        etiqueta_titulo.pack(pady=(0, 5))

        etiqueta_subtitulo = ttk.Label(
            marco_principal,
            text=(
                "PREFIJO + año, mes, día, hora, minutos, "
                "segundos y milisegundos"
            ),
            style="Subtitle.TLabel",
            wraplength=510,
            justify="center",
        )
        etiqueta_subtitulo.pack(pady=(0, 25))

        marco_formulario = ttk.Frame(marco_principal)
        marco_formulario.pack(fill="x")

        etiqueta_prefijo = ttk.Label(
            marco_formulario,
            text="Prefijo:",
        )
        etiqueta_prefijo.grid(
            row=0,
            column=0,
            sticky="w",
            padx=(0, 15),
            pady=8,
        )

        self.campo_prefijo = ttk.Entry(
            marco_formulario,
            textvariable=self.variable_prefijo,
            width=35,
        )
        self.campo_prefijo.grid(
            row=0,
            column=1,
            sticky="ew",
            pady=8,
        )

        etiqueta_cantidad = ttk.Label(
            marco_formulario,
            text="Cantidad:",
        )
        etiqueta_cantidad.grid(
            row=1,
            column=0,
            sticky="w",
            padx=(0, 15),
            pady=8,
        )

        self.campo_cantidad = ttk.Entry(
            marco_formulario,
            textvariable=self.variable_cantidad,
            width=35,
        )
        self.campo_cantidad.grid(
            row=1,
            column=1,
            sticky="ew",
            pady=8,
        )

        etiqueta_formato = ttk.Label(
            marco_formulario,
            text="Formato:",
        )
        etiqueta_formato.grid(
            row=2,
            column=0,
            sticky="w",
            padx=(0, 15),
            pady=8,
        )

        format_frame = ttk.Frame(marco_formulario)
        format_frame.grid(
            row=2,
            column=1,
            sticky="w",
            pady=8,
        )

        self.boton_radio_txt = ttk.Radiobutton(
            format_frame,
            text="Archivo TXT",
            variable=self.variable_formato,
            value="txt",
        )
        self.boton_radio_txt.pack(
            side="left",
            padx=(0, 20),
        )

        self.boton_radio_excel = ttk.Radiobutton(
            format_frame,
            text="Archivo Excel",
            variable=self.variable_formato,
            value="xlsx",
        )
        self.boton_radio_excel.pack(side="left")

        marco_formulario.columnconfigure(1, weight=1)

        etiqueta_limite = ttk.Label(
            marco_principal,
            text=(
                "Cada archivo contendrá un máximo de "
                f"{MAX_CODES_PER_FILE:,} códigos."
            ).replace(",", "."),
            style="Information.TLabel",
        )
        etiqueta_limite.pack(pady=(15, 3))

        etiqueta_archivos = ttk.Label(
            marco_principal,
            textvariable=self.variable_archivos,
            style="Information.TLabel",
        )
        etiqueta_archivos.pack(pady=(0, 12))

        self.barra_progreso = ttk.Progressbar(
            marco_principal,
            orient="horizontal",
            mode="determinate",
            maximum=100,
        )
        self.barra_progreso.pack(
            fill="x",
            pady=(8, 8),
        )

        self.etiqueta_estado = ttk.Label(
            marco_principal,
            textvariable=self.variable_estado,
            anchor="center",
            wraplength=510,
        )
        self.etiqueta_estado.pack(
            fill="x",
            pady=(0, 20),
        )

        self.boton_generar = ttk.Button(
            marco_principal,
            text="Generar códigos",
            command=self.start_generation,
            style="Generate.TButton",
        )
        self.boton_generar.pack(fill="x")

        self.campo_prefijo.focus_set()

    def configurar_eventos(self):
        """Configura los eventos de la interfaz."""
        self.variable_cantidad.trace_add(
            "write",
            self.actualizar_cantidad_archivos_esperados,
        )

        self.root.bind(
            "<Return>",
            lambda event: self.start_generation(),
        )

        self.root.protocol(
            "WM_DELETE_WINDOW",
            self.cerrar_aplicacion,
        )

    def actualizar_cantidad_archivos_esperados(self, *args):
        """Actualiza el número previsto de archivos."""
        texto_cantidad = self.variable_cantidad.get().strip()

        try:
            cantidad = int(texto_cantidad)

            if cantidad <= 0:
                raise ValueError

            total_files = math.ceil(
                cantidad / MAX_CODES_PER_FILE
            )

            self.variable_archivos.set(
                f"Archivos previstos: {total_files}"
            )

        except ValueError:
            self.variable_archivos.set(
                "Archivos previstos: -"
            )

    def obtener_datos_formulario(self):
        """
        Obtiene y valida los valores del formulario.

        Returns:
            tuple[str, int, str]: prefijo, cantidad y formato.
        """
        prefijo = validar_prefijo(
            self.variable_prefijo.get()
        )

        texto_cantidad = self.variable_cantidad.get().strip()

        if not texto_cantidad:
            raise ValueError(
                "Debes introducir una cantidad."
            )

        try:
            cantidad = int(texto_cantidad)
        except ValueError as error:
            raise ValueError(
                "La cantidad debe ser un número entero."
            ) from error

        cantidad = validar_cantidad(cantidad)

        formato_salida = self.variable_formato.get()

        if formato_salida not in ("txt", "xlsx"):
            raise ValueError(
                "Debes seleccionar un formato de salida."
            )

        return prefijo, cantidad, formato_salida

    def solicitar_directorio_salida(self):
        """Solicita la carpeta de destino."""
        return filedialog.askdirectory(
            title="Selecciona la carpeta de destino"
        )

    def start_generation(self):
        """Inicia la generación de códigos."""
        if self.generacion_en_progreso:
            return

        try:
            prefijo, cantidad, formato_salida = (
                self.obtener_datos_formulario()
            )
        except ValueError as error:
            messagebox.showerror(
                "Datos incorrectos",
                str(error),
            )
            return

        directorio_salida = self.solicitar_directorio_salida()

        if not directorio_salida:
            self.variable_estado.set(
                "Generación cancelada."
            )
            return

        total_files = math.ceil(
            cantidad / MAX_CODES_PER_FILE
        )

        self.generacion_en_progreso = True
        self.configurar_interfaz_habilitada(False)

        self.barra_progreso["value"] = 0

        self.variable_estado.set(
            f"Generando 0 de {cantidad} códigos. "
            f"Se crearán {total_files} archivo(s)."
        )

        generation_thread = threading.Thread(
            target=self.generar_y_guardar,
            args=(
                prefijo,
                cantidad,
                formato_salida,
                directorio_salida,
            ),
            daemon=True,
        )

        generation_thread.start()

    def notificar_progreso(self, actual, total):
        """
        Envía el progreso desde el hilo secundario.

        Args:
            actual (int): códigos generados.
            total (int): cantidad total.
        """
        self.cola_mensajes.put(
            (
                "progress",
                {
                    "actual": actual,
                    "total": total,
                },
            )
        )

    def generar_y_guardar(
        self,
        prefijo,
        cantidad,
        formato_salida,
        directorio_salida,
    ):
        """
        Genera, comprueba y guarda los códigos.

        Este método se ejecuta en un hilo secundario.
        """
        try:
            numeros_generados = generar_numeros_unicos(
                cantidad=cantidad,
                prefijo=prefijo,
                progress_callback=self.notificar_progreso,
            )

            # Se comprueba el lote completo antes de dividirlo.
            verificar_sin_duplicados(numeros_generados)

            base_name = construir_nombre_base_archivo(
                prefijo=prefijo,
                cantidad=cantidad,
            )

            if formato_salida == "txt":
                archivos_generados = guardar_lotes_txt(
                    unique_numbers=numeros_generados,
                    directorio_salida=directorio_salida,
                    base_name=base_name,
                )
            else:
                archivos_generados = guardar_lotes_excel(
                    unique_numbers=numeros_generados,
                    directorio_salida=directorio_salida,
                    base_name=base_name,
                )

            self.cola_mensajes.put(
                (
                    "completed",
                    {
                        "numeros_generados": numeros_generados,
                        "archivos_generados": archivos_generados,
                    },
                )
            )

        except Exception as error:
            self.cola_mensajes.put(
                (
                    "error",
                    {
                        "mensaje": str(error),
                    },
                )
            )

    def procesar_mensajes_hilo(self):
        """
        Procesa mensajes del hilo secundario de forma segura.
        """
        try:
            while True:
                tipo_mensaje, datos_mensaje = (
                    self.cola_mensajes.get_nowait()
                )

                if tipo_mensaje == "progress":
                    self.actualizar_progreso(
                        actual=datos_mensaje["actual"],
                        total=datos_mensaje["total"],
                    )

                elif tipo_mensaje == "completed":
                    self.generacion_completada(
                        numeros_generados=datos_mensaje[
                            "numeros_generados"
                        ],
                        archivos_generados=datos_mensaje[
                            "archivos_generados"
                        ],
                    )

                elif tipo_mensaje == "error":
                    self.generacion_fallida(
                        datos_mensaje["mensaje"]
                    )

        except queue.Empty:
            pass

        if self.root.winfo_exists():
            self.root.after(
                50,
                self.procesar_mensajes_hilo,
            )

    def actualizar_progreso(self, actual, total):
        """Actualiza la barra de progreso."""
        percentage = (actual / total) * 100

        self.barra_progreso["value"] = percentage

        self.variable_estado.set(
            f"Generando {actual} de {total} códigos..."
        )

    def generacion_completada(
        self,
        numeros_generados,
        archivos_generados,
    ):
        """Muestra el resumen final."""
        self.generacion_en_progreso = False
        self.configurar_interfaz_habilitada(True)

        self.barra_progreso["value"] = 100

        total_codigos = len(numeros_generados)
        total_files = len(archivos_generados)

        first_code = numeros_generados[0]
        last_code = numeros_generados[-1]

        nombres_archivos = "\n".join(
            ruta_archivo.name
            for ruta_archivo in archivos_generados
        )

        directorio_destino = (
            archivos_generados[0].padre
        )

        self.variable_estado.set(
            f"Completado: {total_codigos} códigos "
            f"en {total_files} archivo(s)."
        )

        messagebox.showinfo(
            "Generación completada",
            (
                f"Códigos generados: {total_codigos}\n"
                f"Archivos creados: {total_files}\n\n"
                f"Primer código:\n{first_code}\n\n"
                f"Último código:\n{last_code}\n\n"
                f"Carpeta de destino:\n"
                f"{directorio_destino}\n\n"
                f"Archivos generados:\n{nombres_archivos}"
            ),
        )

    def generacion_fallida(self, mensaje_error):
        """Muestra los errores de generación o guardado."""
        self.generacion_en_progreso = False
        self.configurar_interfaz_habilitada(True)

        self.variable_estado.set(
            "Se produjo un error durante la generación."
        )

        messagebox.showerror(
            "Error",
            mensaje_error,
        )

    def configurar_interfaz_habilitada(self, enabled):
        """Activa o desactiva los controles."""
        state = "normal" if enabled else "disabled"

        self.campo_prefijo.configure(state=state)
        self.campo_cantidad.configure(state=state)
        self.boton_radio_txt.configure(state=state)
        self.boton_radio_excel.configure(state=state)
        self.boton_generar.configure(state=state)

    def cerrar_aplicacion(self):
        """Gestiona el cierre de la aplicación."""
        if self.generacion_en_progreso:
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

def principal():
    """Inicia la interfaz gráfica."""
    root = tk.Tk()
    GeneradorCodigosUnicosApp(root)
    root.mainloop()


if __name__ == "__main__":
    principal()