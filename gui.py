#TODO 1) Fix settings paths
#     2) Minimize logs, create progress bar
#     3) Change state when program has ended up succsecfully 

import sys
import os
import subprocess
from PyQt5.QtWidgets import QApplication, QFileDialog, QLabel, QPushButton, QHBoxLayout, QVBoxLayout, QWidget, QTextEdit, QDialog
from PyQt5.QtCore import Qt, QThread, pyqtSignal
import pathlib


class Worker(QThread):
    finished = pyqtSignal()

    def __init__(self, cmd):
        super().__init__()
        self.cmd = cmd

    def run(self):
        os.system(self.cmd)
        self.finished.emit()


catalog_path = None
full_load_path = None

class ProcessThread(QThread):
    finished = pyqtSignal(int)
    output = pyqtSignal(str)

    def __init__(self, cmd):
        super().__init__()
        self.cmd = cmd

    def run(self):
        process = subprocess.Popen(self.cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True)
        while process.poll() is None:
            line = process.stdout.readline().decode('utf-8').rstrip()
            if line:
                self.output.emit(line)
        self.finished.emit(process.returncode)


class SettingsDialog(QDialog):
    def __init__(self, parent, log_box):
        super().__init__(parent)
        self.setWindowTitle('Настройки')
        self.setFixedSize(500, 300)

        self.file1_button = QPushButton('Изменить файл каталога', self)
        self.file1_button.clicked.connect(self.showFileDialog1)
        self.file1_label = QLabel(self)
        self.file1_path = ""
        self.file1_button.setStyleSheet("QPushButton {background-color: gray; color: white; border: none; border-radius: 10px; padding: 5px;}")


        self.file2_button = QPushButton('Изменить файл полной выгрузки', self)
        self.file2_button.clicked.connect(self.showFileDialog2)
        self.file2_label = QLabel(self)
        self.file2_path = ""
        self.file2_button.setStyleSheet("QPushButton {background-color: gray; color: white; border: none; border-radius: 10px; padding: 5px;}")

        self.save_button = QPushButton('Сохранить', self)
        self.save_button.clicked.connect(self.saveSettings)
        self.save_button.setStyleSheet("QPushButton {background-color: green; color: white; border: none; border-radius: 10px; padding: 5px;}")

        self.cancel_button = QPushButton('Отмена', self)
        self.cancel_button.clicked.connect(self.reject)
        self.cancel_button.setStyleSheet("QPushButton {background-color: red; color: white; border: none; border-radius: 10px; padding: 5px;}")

        self.log_box = log_box

        layout = QVBoxLayout()
        layout.addWidget(self.file1_label)
        layout.addWidget(self.file1_button)
        layout.addWidget(self.file2_label)
        layout.addWidget(self.file2_button)
        button_layout = QVBoxLayout()
        button_layout.addWidget(self.save_button)
        button_layout.addWidget(self.cancel_button)
        layout.addLayout(button_layout)
        self.setLayout(layout)

    def showFileDialog1(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self,"Выбор файла:", "","All Files (*);;Python Files (*.py)", options=options)
        if fileName:
            self.file1_path = fileName
            self.file1_label.setText(fileName)

    def showFileDialog2(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self,"Выбор файла:", "","All Files (*);;Python Files (*.py)", options=options)
        if fileName:
            self.file2_path = fileName
            self.file2_label.setText(fileName)

    def saveSettings(self):
        if self.file1_path != "":
            catalog_path = self.file1_path
            self.log_box.append("Изменен файл каталога")
        if self.file2_path != "":
            full_load_path = self.file2_path
            self.log_box.append("Изменен файл полной выгрузки")
        self.accept()

    def getSettings(self):
        return self.file1_path, self.file2_path


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

        self.process_thread = None

    def closeEvent(self, event):
        QApplication.instance().aboutToQuit.emit()
        super().closeEvent(event)

    def initUI(self):
        self.setWindowTitle('OZON конфигуратор')
        self.resize(800, 500)

        self.button1 = QPushButton('Выберите конфигурационный файл', self)
        self.button1.clicked.connect(self.showFileDialog1)
        self.button1.setStyleSheet("QPushButton {background-color: gray; color: white; border: none; border-radius: 10px; padding: 5px;}")
        self.button2 = QPushButton('Выберите шаблон', self)
        self.button2.clicked.connect(self.showFileDialog2)
        self.button2.setStyleSheet("QPushButton {background-color: gray; color: white; border: none; border-radius: 10px; padding: 5px;}")
        self.settings_button = QPushButton('Настройки', self)
        self.settings_button.clicked.connect(self.showSettings)
        self.settings_button.setStyleSheet("QPushButton {background-color: gray; color: white; border: none; border-radius: 10px; padding: 5px;}")

        button_layout = QVBoxLayout()
        button_layout.addWidget(self.button1)
        button_layout.addWidget(self.button2)

        top_layout = QVBoxLayout()
        top_layout.addWidget(self.button1)
        top_layout.addWidget(self.button2)
        top_layout.addStretch(1)
        top_layout.addWidget(self.settings_button)

        self.start_button = QPushButton('Начать перенос данных', self)
        self.start_button.clicked.connect(self.onStartStopClicked)
        self.start_button.setStyleSheet("QPushButton {background-color: green; color: white; border: none; border-radius: 10px; padding: 5px;}")
       
        self.label1 = QLabel(self)
        self.label2 = QLabel(self)

        self.log_box = QTextEdit(self)
        self.log_box.setReadOnly(True)


        layout = QVBoxLayout()
        layout.addLayout(top_layout)
        layout.addLayout(button_layout)
        layout.addWidget(self.label1)
        layout.addWidget(self.label2)
        layout.addWidget(self.log_box)
        layout.addWidget(self.start_button)
        self.setLayout(layout)

    def redirect_stdout(self):
        sys.stdout = QTextStream(self.log_box)
        sys.stdout.setCodec("UTF-8")
        sys.stderr = QTextStream(self.log_box)
        sys.stderr.setCodec("UTF-8")
        sys.stdout.flush = lambda: None

    def showSettings(self):
        self.settings_window = SettingsDialog(self, self.log_box)
        self.settings_window.show()

    def showFileDialog1(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self,"Выбор файла:", "","All Files (*);;Python Files (*.py)", options=options)
        if fileName:
            self.label1.setText(fileName)

    def showFileDialog2(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self,"Выбор файла:", "","All Files (*);;Python Files (*.py)", options=options)
        if fileName:
            self.label2.setText(fileName)

    def onStartStopClicked(self):
        if self.start_button.text() == 'Начать перенос данных':
            #if not catalog_path or not full_load_path:
                #self.log_box.append('Не выбраны файлы настроек')
                #return
            filePath1 = self.label1.text()
            filePath2 = self.label2.text()
            cmd = "python3 program_opt.py \"" + filePath1 + "\" \"" + filePath2 + "\""
            if catalog_path is not None:
                cmd += " \"" + catalog_path + "\""
            if full_load_path is not None:
                cmd += " \"" + full_load_path + "\""
            print(cmd)
            self.log_box.append('Запущен процесс загрузки')
            self.start_button.setStyleSheet("QPushButton {background-color: red; color: white; border: none; border-radius: 10px; padding: 5px;}")
            self.log_box.clear()
            self.thread = ProcessThread(cmd)
            self.thread.output.connect(self.update_log)
            self.thread.start()
            self.start_button.setText('Остановить процес переноса данных')
        else:
            self.start_button.setStyleSheet("QPushButton {background-color: green; color: white; border: none; border-radius: 10px; padding: 5px;}")
            self.log_box.append('Остановлен процесс загрузки')
            self.thread.terminate()
            self.thread.finished.connect(self.stop)
            self.start_button.setText('Начать перенос данных')


    def update_log(self, line):
        self.log_box.append(line)

    def stop(self):
        self.log_box.append("Программа завершена.")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()

    sys.exit(app.exec_())
