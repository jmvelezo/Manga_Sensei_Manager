import sys
import os
import shutil
import tempfile
import subprocess
import logging
import locale
import rarfile
from zipfile import ZipFile, BadZipFile
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog,
    QListWidget, QListWidgetItem, QAbstractItemView, QMessageBox,
    QLabel, QHBoxLayout, QCheckBox, QTextEdit, QTabWidget, QSplashScreen
)
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QIcon

class CBRCreator(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Manga Sensei CBR Manager')
        self.setGeometry(100, 100, 800, 600)

        self.tabs = QTabWidget()

        self.manual_conversion_tab = QWidget()
        self.init_manual_conversion_tab()
        self.tabs.addTab(self.manual_conversion_tab, 'Conversión y Actualización Manual')

        self.update_library_tab = QWidget()
        self.init_update_library_tab()
        self.tabs.addTab(self.update_library_tab, 'Actualizar Biblioteca')

   
        main_layout = QVBoxLayout()
        main_layout.addWidget(self.tabs)
        self.setLayout(main_layout)

    
        logo_path = os.path.join(sys._MEIPASS, 'logo.png') if hasattr(sys, '_MEIPASS') else 'logo.png'

    
        self.logo_label = QLabel()
        self.logo_label.setPixmap(QPixmap(logo_path).scaled(175, 175, Qt.KeepAspectRatio))
        logo_layout = QHBoxLayout()
        logo_layout.addWidget(self.logo_label)
        self.explanation_label = QLabel()
        self.explanation_label.setWordWrap(True)
        logo_layout.addWidget(self.explanation_label, stretch=1)
        main_layout.insertLayout(0, logo_layout)

    
        self.update_tutorial_text(0)

        self.tabs.currentChanged.connect(self.update_tutorial_text)
        
        
    
        credit_label = QLabel(
            '<p>Creado por: Verdura y el logo por LOSion - VER: 1.0 - '
            '<a href="https://github.com/jmvelezo" style="color: #0000EE;">GitHub</a></p>'
        )
        credit_label.setOpenExternalLinks(True)
        credit_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(credit_label)

    def update_tutorial_text(self, index):
       if index == 0:
        self.explanation_label.setText(
            '<p>En la pestaña "Conversión y Actualización Manual", puede realizar las siguientes acciones:</p>'
            '<ul>'
            '<li>Agregar carpetas: con esta opción selecciona la carpeta que contiene las imagenes para convertir en CBR, si le das en la ruta de salida, te pedira el nombre para el CBR y si le das en la casilla de "Generar archivos cbr en la misma carpeta" de la parte inferior, generará el CBR con el mismo nombre de la carpeta que contiene las imagenes</li>'
            '<li>Agregar archivo CBR: esto sirve para actualizar un CBR que ya tengas, por ej con comics que se van actualizando, selecciona primero tu archivo CBR y luego con seleccionar carpeta agrega la carpeta que contiene las imagenes para agregar al CBR. en el recuerdo veras los archivos, siempre debe quedar al inicio el CBR para que las nuevas imaganes se agregen luego. despues selecciona ruta de salida y lueco crear o actualizar</li>'
            '<li>COMBINACION: si agregas más de una carpeta al recuadro, se creará un sólo CBR en el orden de carga que aparece en el recuadro. </li>'
            '<li>La opcion crear archivos CBR individuales por carpeta sirve para seleccionar varias carpetas con imagenes y que no se cree un solo CBR combinado si no que cada una se creara de forma independiente en su carpeta original con el nombre de la carpeta raíz de las imagenes</li>'
            '</ul>'
        )
       elif index == 1:
        self.explanation_label.setText(
            '<p>En la pestaña "Actualizar Biblioteca", Esto está especialmente pensado para bibliotecas descargadas con el software Hitomi Downloader (MUY RECOMENDADO)... puede realizar las siguientes acciones:</p>'
            '<ul>'
            '<li>Selecciona una carpeta donde tengas tu biblioteca para convertir u ordenar. las carpetas que descargas deben tener el nombre del autor entre corchetes al inicio. EJ: [Kentaro Miura] Berserk Cap 1 a 4 (43534534), lo que hará esta opcion es crear una carpeta llamada "Kentaro Miura" y allí enviará los CBR creados a partir de las diversas carpetas de comics de este autor. es un procese en masa, si la carpeta ya está creada agregará los CBR nuevos a ella. </li>'
            '<li>Actualizar los archivos CBR dentro de la biblioteca seleccionada.</li>'
            '</ul>'
        )
        
        
    def init_update_library_tab(self):
        layout = QVBoxLayout()

   
        self.select_library_btn = QPushButton('Seleccionar Biblioteca')
        self.select_library_btn.clicked.connect(self.select_library)
        layout.addWidget(self.select_library_btn)

        self.log_viewer = QTextEdit()
        self.log_viewer.setReadOnly(True)
        layout.addWidget(self.log_viewer)

      
        self.start_process_btn = QPushButton('Iniciar Proceso')
        self.start_process_btn.clicked.connect(self.start_process)
        layout.addWidget(self.start_process_btn)

      
        self.clear_btn = QPushButton('Limpiar Proceso')
        self.clear_btn.clicked.connect(self.clear_process)
        layout.addWidget(self.clear_btn)

        self.update_library_tab.setLayout(layout)

      
        logger = logging.getLogger()
        logger.setLevel(logging.DEBUG)
        log_handler = self.QtLogHandler(self.log_viewer)
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        log_handler.setFormatter(formatter)
        logger.addHandler(log_handler)

    def init_manual_conversion_tab(self):
        layout = QVBoxLayout()

        self.folder_list = []
        self.cbr_file = None

       
        self.add_folder_btn = QPushButton('Agregar Carpeta')
        self.add_folder_btn.clicked.connect(self.add_folder)
        layout.addWidget(self.add_folder_btn)

       
        self.add_cbr_btn = QPushButton('Agregar Archivo CBR')
        self.add_cbr_btn.clicked.connect(self.add_cbr)
        layout.addWidget(self.add_cbr_btn)

       
        self.folder_list_widget = QListWidget()
        self.folder_list_widget.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.folder_list_widget.setDragDropMode(QAbstractItemView.InternalMove)
        layout.addWidget(self.folder_list_widget)

      
        self.individual_cbr_checkbox = QCheckBox("Crear archivos CBR individuales por carpeta")
        layout.addWidget(self.individual_cbr_checkbox)

       
        self.same_folder_checkbox = QCheckBox("Generar archivos CBR en la misma carpeta")
        layout.addWidget(self.same_folder_checkbox)

 
        self.output_label = QLabel('Archivo de Salida:')
        self.output_path = ''
        self.output_btn = QPushButton('Seleccionar Ubicación de Salida')
        self.output_btn.clicked.connect(self.select_output)
        output_layout = QHBoxLayout()
        output_layout.addWidget(self.output_label)
        output_layout.addWidget(self.output_btn)
        layout.addLayout(output_layout)

        self.create_btn = QPushButton('Crear o Actualizar CBR')
        self.create_btn.clicked.connect(self.create_cbr)
        layout.addWidget(self.create_btn)

        self.clear_manual_btn = QPushButton('Limpiar Proceso')
        self.clear_manual_btn.clicked.connect(self.clear_manual_process)
        layout.addWidget(self.clear_manual_btn)

        self.manual_conversion_tab.setLayout(layout)

    class QtLogHandler(logging.Handler):
        def __init__(self, log_viewer):
            super().__init__()
            self.log_viewer = log_viewer

        def emit(self, record):
            log_entry = self.format(record)
            self.log_viewer.append(log_entry)

    def select_library(self):
        self.library_folder = QFileDialog.getExistingDirectory(self, 'Seleccionar Biblioteca')
        if self.library_folder:
            logging.info(f'Biblioteca seleccionada: {self.library_folder}')

    def start_process(self):
        if not self.library_folder:
            QMessageBox.warning(self, 'Advertencia', 'No se ha seleccionado una biblioteca para procesar.')
            return

        for root, dirs, files in os.walk(self.library_folder):
            for directory in dirs:
                # Solo procesar las carpetas de la primera capa que comienzan con corchetes
                if root == self.library_folder and directory.startswith('['):
                    author_name = directory.split(']')[0].strip('[]')
                    author_folder = os.path.join(self.library_folder, author_name)
                    if not os.path.exists(author_folder):
                        os.makedirs(author_folder)
                        logging.info(f'Carpeta de autor creada: {author_folder}')

                    comic_folder_path = os.path.join(root, directory)
                    cbr_name = f"{directory}.cbr"
                    cbr_output_path = os.path.join(author_folder, cbr_name)

                    # Crear archivo CBR a partir de la carpeta del cómic
                    zip_output = cbr_output_path.replace('.cbr', '.zip')
                    with ZipFile(zip_output, 'w') as zipf:
                        for comic_root, _, comic_files in os.walk(comic_folder_path):
                            images = [f for f in comic_files if f.lower().endswith(('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.webp'))]
                            images.sort()
                            for img in images:
                                src = os.path.join(comic_root, img)
                                zipf.write(src, os.path.basename(src))
                                logging.info(f'Añadiendo imagen al ZIP: {src}')

                    shutil.move(zip_output, cbr_output_path)
                    logging.info(f'Archivo CBR creado: {cbr_output_path}')
                    shutil.rmtree(comic_folder_path)
                    logging.info(f'Carpeta de cómic eliminada: {comic_folder_path}')

        QMessageBox.information(self, 'Proceso Completado', 'Biblioteca procesada exitosamente.')

    def clear_process(self):
        self.library_folder = None
        self.log_viewer.clear()
        logging.info('Proceso limpiado')
        QMessageBox.information(self, 'Proceso Limpiado', 'Se ha limpiado el proceso actual.')

    def add_folder(self):
        folder = QFileDialog.getExistingDirectory(self, 'Seleccionar Carpeta')
        if folder:
            self.folder_list.append(folder)
            item = QListWidgetItem(folder)
            self.folder_list_widget.addItem(item)

    def add_cbr(self):
        file, _ = QFileDialog.getOpenFileName(self, "Seleccionar Archivo CBR", "", "Archivos CBR (*.cbr)")
        if file:
            self.cbr_file = file
            item = QListWidgetItem(file)
            self.folder_list_widget.addItem(item)
            QMessageBox.information(self, 'Archivo CBR seleccionado', f'Se ha seleccionado el archivo: {file}')

    def select_output(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file, _ = QFileDialog.getSaveFileName(
            self, "Seleccionar Archivo de Salida", "", "Archivos CBR (*.cbr)", options=options)
        if file:
            if not file.lower().endswith('.cbr'):
                file += '.cbr'
            self.output_path = file
            self.output_label.setText(f'Archivo de Salida: {self.output_path}')

    def create_cbr(self):
        if self.folder_list_widget.count() == 0 and not self.cbr_file:
            QMessageBox.warning(self, 'Advertencia', 'No se han seleccionado carpetas ni archivo CBR.')
            return

        # Si la opción de generar archivos CBR en la misma carpeta está activada
        if self.same_folder_checkbox.isChecked():
            # Crear un archivo CBR en la misma carpeta por cada carpeta seleccionada
            for i in range(self.folder_list_widget.count()):
                folder = self.folder_list_widget.item(i).text()
                folder_name = os.path.basename(folder)
                cbr_output_path = os.path.join(folder, f"{folder_name}.cbr")
                logging.debug(f"Creando archivo CBR individual en la carpeta: {folder}")

                zip_output = cbr_output_path.replace('.cbr', '.zip')
                with ZipFile(zip_output, 'w') as zipf:
                    for root, dirs, files in os.walk(folder):
                        images = [f for f in files if f.lower().endswith(('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.webp'))]
                        images.sort()
                        for img in images:
                            src = os.path.join(root, img)
                            zipf.write(src, os.path.basename(src))

                shutil.move(zip_output, cbr_output_path)

            QMessageBox.information(self, 'Éxito', 'Archivos CBR generados en las carpetas correspondientes.')
            return

        # Crear archivos CBR individuales heredando el nombre de la carpeta automáticamente
        if self.individual_cbr_checkbox.isChecked():
            for i in range(self.folder_list_widget.count()):
                folder = self.folder_list_widget.item(i).text()
                folder_name = os.path.basename(folder)
                cbr_output_path = os.path.join(os.path.dirname(folder), f"{folder_name}.cbr")
                logging.debug(f"Creando archivo CBR individual para la carpeta: {folder}")

                # Crear el archivo ZIP y luego renombrarlo a CBR
                zip_output = cbr_output_path.replace('.cbr', '.zip')
                with ZipFile(zip_output, 'w') as zipf:
                    for root, dirs, files in os.walk(folder):
                        images = [f for f in files if f.lower().endswith(('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.webp'))]
                        images.sort()
                        for img in images:
                            src = os.path.join(root, img)
                            zipf.write(src, os.path.basename(src))

                shutil.move(zip_output, cbr_output_path)

            QMessageBox.information(self, 'Éxito', 'Archivos CBR individuales creados para cada carpeta.')
            return

        # Si no se seleccionó la opción de generar archivos en la misma carpeta
        if not self.output_path:
            QMessageBox.warning(self, 'Advertencia', 'No se ha seleccionado el archivo de salida.')
            return

        # Verificar si el comando 'zip' está disponible
        if shutil.which('zip') is None:
            QMessageBox.critical(
                self, 'Error',
                'El comando "zip" no está disponible. Por favor, instala las herramientas de línea de comandos de ZIP.')
            return

        # Crear directorio temporal
        temp_dir = tempfile.mkdtemp()
        logging.debug(f'Directorio temporal creado: {temp_dir}')
        image_counter = 1
        file_list = []

        try:
            # Proceso normal de secuencia de imágenes en un solo archivo CBR
            if self.cbr_file:
                try:
                    with ZipFile(self.cbr_file, 'r') as zip_ref:
                        zip_ref.extractall(temp_dir)
                        logging.debug(f'Descomprimido como ZIP')
                except BadZipFile:
                    logging.debug(f'No es un archivo ZIP, intentando como RAR')
                    with rarfile.RarFile(self.cbr_file) as rar_ref:
                        rar_ref.extractall(temp_dir)
                        logging.debug(f'Descomprimido como RAR')

                existing_files = [f for f in os.listdir(temp_dir) if f.lower().endswith(('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.webp'))]
                if existing_files:
                    existing_files.sort()
                    last_image = existing_files[-1]
                    image_counter = int(os.path.splitext(last_image)[0]) + 1
                logging.debug(f'Última imagen existente: {last_image}, siguiente contador: {image_counter}')
                file_list.extend([os.path.join(temp_dir, f) for f in existing_files])

            for i in range(self.folder_list_widget.count()):
                folder = self.folder_list_widget.item(i).text()
                logging.debug(f'Procesando carpeta: {folder}')
                for root, dirs, files in os.walk(folder):
                    images = [f for f in files if f.lower().endswith(('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.webp'))]
                    images.sort()
                    for img in images:
                        src = os.path.join(root, img)
                        ext = os.path.splitext(img)[1]
                        dst_filename = f'{image_counter:04d}{ext}'
                        dst = os.path.join(temp_dir, dst_filename)
                        shutil.copy2(src, dst)
                        file_list.append(dst)
                        image_counter += 1

            file_list.sort()
            logging.debug(f'Archivos a agregar al CBR: {file_list}')

            if not file_list:
                QMessageBox.warning(self, 'Advertencia', 'No se encontraron imágenes para agregar al archivo CBR.')
                return

            zip_output = self.output_path.replace('.cbr', '.zip')
            with ZipFile(zip_output, 'w') as zipf:
                for file in file_list:
                    zipf.write(file, os.path.basename(file))

            shutil.move(zip_output, self.output_path)
            QMessageBox.information(self, 'Éxito', f'Archivo CBR actualizado en {self.output_path}')
        except Exception as e:
            logging.error(f'Error al crear o actualizar el CBR: {e}')
            QMessageBox.critical(self, 'Error', str(e))
        finally:
            shutil.rmtree(temp_dir)
            logging.debug(f'Directorio temporal eliminado: {temp_dir}')

    def clear_manual_process(self):
        self.folder_list.clear()
        self.cbr_file = None
        self.output_path = ''
        self.folder_list_widget.clear()
        self.output_label.setText('Archivo de Salida:')
        self.individual_cbr_checkbox.setChecked(False)
        self.same_folder_checkbox.setChecked(False)
        QMessageBox.information(self, 'Proceso Limpiado', 'Se ha limpiado el proceso actual.')

if __name__ == '__main__':
    import logging
    locale.setlocale(locale.LC_ALL, 'C')

    app = QApplication(sys.argv)
    icon_path = os.path.join(sys._MEIPASS, 'logo.ico') if hasattr(sys, '_MEIPASS') else 'logo.ico'
    app.setWindowIcon(QIcon(icon_path))


    splash_logo_path = os.path.join(sys._MEIPASS, 'logo.png') if hasattr(sys, '_MEIPASS') else 'logo.png'

    splash_pix = QPixmap(splash_logo_path).scaled(375, 375, Qt.KeepAspectRatio)
    splash = QSplashScreen(splash_pix, Qt.WindowStaysOnTopHint)
    splash.setMask(splash_pix.mask())
    splash.show()

    window = CBRCreator()
    QTimer.singleShot(4000, lambda: (splash.close(), window.show()))  # Mostrar el logo por 4 segundos y luego abrir la app

    sys.exit(app.exec_())


