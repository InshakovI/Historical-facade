import sys
import sqlite3
from datetime import datetime
from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, 
                            QLabel, QPushButton, QWidget, QTableWidget, 
                            QTableWidgetItem, QMessageBox, QComboBox)
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QIcon, QPixmap, QFont, QImage
import qrcode
from pyzbar.pyzbar import decode
import cv2
import numpy as np

class QRScannerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("QR Scanner - Учет посетителей")
        self.setWindowIcon(QIcon("icon.png"))
        self.setGeometry(100, 100, 900, 600)
        
        # Инициализация базы данных
        self.init_db()
        
        # Создание интерфейса
        self.init_ui()
        
        # Настройка камеры
        self.camera = cv2.VideoCapture(0)
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_frame)
        self.scanning = False
        
    def init_db(self):
        self.conn = sqlite3.connect('visitors.db')
        self.cursor = self.conn.cursor()
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS visitors (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT,
                phone TEXT,
                email TEXT,
                qr_data TEXT UNIQUE,
                visit_time DATETIME,
                event TEXT
            )
        ''')
        self.conn.commit()
        
    def init_ui(self):
        # Главный виджет
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        
        # Главный макет
        main_layout = QHBoxLayout()
        main_widget.setLayout(main_layout)
        
        # Левая панель (сканер)
        left_panel = QWidget()
        left_layout = QVBoxLayout()
        left_panel.setLayout(left_layout)
        
        # Правая панель (информация)
        right_panel = QWidget()
        right_layout = QVBoxLayout()
        right_panel.setLayout(right_layout)
        
        # Добавляем панели в главный макет
        main_layout.addWidget(left_panel, 1)
        main_layout.addWidget(right_panel, 2)
        
        # Настройка левой панели
        self.camera_label = QLabel()
        self.camera_label.setAlignment(Qt.AlignCenter)
        self.camera_label.setMinimumSize(400, 300)
        self.camera_label.setStyleSheet("background-color: #f0f0f0; border: 2px solid #ccc;")
        
        self.scan_button = QPushButton("Начать сканирование")
        self.scan_button.setStyleSheet(
            "QPushButton {"
            "background-color: #4CAF50;"
            "border: none;"
            "color: white;"
            "padding: 10px 24px;"
            "text-align: center;"
            "font-size: 16px;"
            "border-radius: 4px;"
            "}"
            "QPushButton:hover {"
            "background-color: #45a049;"
            "}"
        )
        self.scan_button.clicked.connect(self.toggle_scan)
        
        left_layout.addWidget(QLabel("<h2 style='color: #333;'>Сканер QR-кодов</h2>"))
        left_layout.addWidget(self.camera_label)
        left_layout.addWidget(self.scan_button)
        left_layout.addStretch()
        
        # Настройка правой панели
        self.info_label = QLabel("<h3 style='color: #333;'>Информация о посетителе</h3>")
        
        self.visitor_table = QTableWidget()
        self.visitor_table.setColumnCount(5)
        self.visitor_table.setHorizontalHeaderLabels(["Имя", "Телефон", "Email", "Время посещения", "Мероприятие"])
        self.visitor_table.setStyleSheet("""
            QTableWidget {
                border: 1px solid #ddd;
                font-size: 14px;
            }
            QHeaderView::section {
                background-color: #f8f8f8;
                padding: 5px;
                border: 1px solid #ddd;
            }
        """)
        
        self.event_combo = QComboBox()
        self.event_combo.addItems(["Конференция", "Выставка", "Концерт", "Другое"])
        self.event_combo.setStyleSheet("""
            QComboBox {
                padding: 5px;
                font-size: 14px;
                border: 1px solid #ddd;
                border-radius: 4px;
            }
        """)
        
        self.export_button = QPushButton("Экспорт в Excel")
        self.export_button.setStyleSheet(
            "QPushButton {"
            "background-color: #2196F3;"
            "border: none;"
            "color: white;"
            "padding: 10px 24px;"
            "text-align: center;"
            "font-size: 16px;"
            "border-radius: 4px;"
            "}"
            "QPushButton:hover {"
            "background-color: #0b7dda;"
            "}"
        )
        self.export_button.clicked.connect(self.export_to_excel)
        
        right_layout.addWidget(self.info_label)
        right_layout.addWidget(QLabel("Мероприятие:"))
        right_layout.addWidget(self.event_combo)
        right_layout.addWidget(self.visitor_table)
        right_layout.addWidget(self.export_button)
        
        # Обновляем таблицу
        self.update_visitors_table()
        
    def toggle_scan(self):
        if not self.scanning:
            self.scanning = True
            self.scan_button.setText("Остановить сканирование")
            self.scan_button.setStyleSheet(
                "QPushButton {"
                "background-color: #f44336;"
                "border: none;"
                "color: white;"
                "padding: 10px 24px;"
                "text-align: center;"
                "font-size: 16px;"
                "border-radius: 4px;"
                "}"
                "QPushButton:hover {"
                "background-color: #d32f2f;"
                "}"
            )
            self.timer.start(20)
        else:
            self.scanning = False
            self.scan_button.setText("Начать сканирование")
            self.scan_button.setStyleSheet(
                "QPushButton {"
                "background-color: #4CAF50;"
                "border: none;"
                "color: white;"
                "padding: 10px 24px;"
                "text-align: center;"
                "font-size: 16px;"
                "border-radius: 4px;"
                "}"
                "QPushButton:hover {"
                "background-color: #45a049;"
                "}"
            )
            self.timer.stop()
            # Очищаем изображение
            self.camera_label.clear()
            self.camera_label.setText("Камера отключена")
        
    def update_frame(self):
        ret, frame = self.camera.read()
        if ret:
            # Конвертируем кадр в RGB
            frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            
            # Декодируем QR-код
            decoded_objects = decode(frame)
            
            # Если найден QR-код
            if decoded_objects:
                qr_data = decoded_objects[0].data.decode('utf-8')
                self.process_qr_code(qr_data)
                self.timer.stop()
                self.scanning = False
                self.scan_button.setText("Начать сканирование")
                self.scan_button.setStyleSheet(
                    "QPushButton {"
                    "background-color: #4CAF50;"
                    "border: none;"
                    "color: white;"
                    "padding: 10px 24px;"
                    "text-align: center;"
                    "font-size: 16px;"
                    "border-radius: 4px;"
                    "}"
                    "QPushButton:hover {"
                    "background-color: #45a049;"
                    "}"
                )
            
            # Отображаем кадр
            h, w, ch = frame_rgb.shape
            bytes_per_line = ch * w
            q_img = QImage(frame_rgb.data, w, h, bytes_per_line, QImage.Format_RGB888)
            self.camera_label.setPixmap(QPixmap.fromImage(q_img).scaled(
                self.camera_label.width(), 
                self.camera_label.height(), 
                Qt.KeepAspectRatio
            ))
    
    def process_qr_code(self, qr_data):
        try:
            # Проверяем, есть ли такой QR в базе
            self.cursor.execute("SELECT * FROM visitors WHERE qr_data=?", (qr_data,))
            visitor = self.cursor.fetchone()
            
            if visitor:
                # Обновляем время посещения
                visit_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                event = self.event_combo.currentText()
                self.cursor.execute(
                    "UPDATE visitors SET visit_time=?, event=? WHERE qr_data=?",
                    (visit_time, event, qr_data)
                )
                self.conn.commit()
                
                QMessageBox.information(
                    self, "Успешно", 
                    f"Посетитель {visitor[1]} уже зарегистрирован.\nВремя обновлено."
                )
            else:
                # Добавляем нового посетителя
                visit_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                event = self.event_combo.currentText()
                
                # В реальном приложении здесь можно добавить форму для ввода данных
                name = "Посетитель " + str(datetime.now().timestamp())[-4:]
                phone = "Не указан"
                email = "Не указан"
                
                self.cursor.execute(
                    "INSERT INTO visitors (name, phone, email, qr_data, visit_time, event) "
                    "VALUES (?, ?, ?, ?, ?, ?)",
                    (name, phone, email, qr_data, visit_time, event)
                )
                self.conn.commit()
                
                QMessageBox.information(
                    self, "Успешно", 
                    f"Новый посетитель {name} зарегистрирован!"
                )
            
            # Обновляем таблицу
            self.update_visitors_table()
            
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка: {str(e)}")
    
    def update_visitors_table(self):
        self.cursor.execute("SELECT name, phone, email, visit_time, event FROM visitors ORDER BY visit_time DESC LIMIT 50")
        visitors = self.cursor.fetchall()
        
        self.visitor_table.setRowCount(len(visitors))
        for row, visitor in enumerate(visitors):
            for col, data in enumerate(visitor):
                item = QTableWidgetItem(str(data))
                item.setTextAlignment(Qt.AlignCenter)
                self.visitor_table.setItem(row, col, item)
        
        self.visitor_table.resizeColumnsToContents()
    
    def export_to_excel(self):
        try:
            import pandas as pd
            self.cursor.execute("SELECT name, phone, email, visit_time, event FROM visitors")
            visitors = self.cursor.fetchall()
            
            df = pd.DataFrame(visitors, columns=["Имя", "Телефон", "Email", "Время посещения", "Мероприятие"])
            df.to_excel("visitors.xlsx", index=False)
            
            QMessageBox.information(self, "Успешно", "Данные экспортированы в visitors.xlsx")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось экспортировать: {str(e)}")
    
    def closeEvent(self, event):
        self.camera.release()
        self.conn.close()
        event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # Установка стиля
    app.setStyle("Fusion")
    
    window = QRScannerApp()
    window.show()
    
    sys.exit(app.exec_())