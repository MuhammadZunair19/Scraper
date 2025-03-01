import sys
import time
import pandas as pd
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QIcon, QFont
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLineEdit, QLabel, QTableWidget, \
    QTableWidgetItem, QCheckBox, QHBoxLayout, QFileDialog, QProgressBar, QHeaderView
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup


class ScraperThread(QThread):
    progress_signal = pyqtSignal(int)
    data_signal = pyqtSignal(list)

    def __init__(self, url, selected_tags, tags):
        super().__init__()
        self.url = url
        self.selected_tags = selected_tags
        self.tags = tags

    def run(self):
        options = Options()
        options.add_argument("--headless")
        options.add_argument("--disable-gpu")
        options.add_argument("--window-size=1920x1080")

        driver = webdriver.Chrome(options=options)
        driver.get(self.url)
        time.sleep(2)

        last_height = driver.execute_script("return document.body.scrollHeight")

        while True:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)
            new_height = driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                break
            last_height = new_height

        soup = BeautifulSoup(driver.page_source, 'html.parser')
        articles = soup.find_all('article', class_='store store-item border-accent single-image')
        driver.quit()

        data = []
        total_items = len(articles)
        for i, article in enumerate(articles):
            row = []
            for tag in self.selected_tags:
                element = article.select_one(self.tags[tag])
                text = element.text.strip() if element and tag != "Image" else (
                    element['src'] if element and element.has_attr('src') else "")
                row.append(text)
            data.append(row)
            self.progress_signal.emit(int((i / total_items) * 100))

        self.progress_signal.emit(100)
        self.data_signal.emit(data)


class ScraperApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("🕵️‍♂️ Web Scraper Pro")
        self.setGeometry(100, 100, 900, 600)
        self.setStyleSheet("background-color: #1e1e1e; color: white;")

        layout = QVBoxLayout()

        title_label = QLabel("🔍 Dynamic Web Scraper")
        title_label.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title_label)

        self.url_input = QLineEdit(self)
        self.url_input.setPlaceholderText("Enter website URL...")
        self.url_input.setStyleSheet("background: #333; color: white; padding: 8px; border-radius: 5px;")
        layout.addWidget(self.url_input)

        tag_layout = QHBoxLayout()
        self.checkboxes = {}
        self.tags = {"Title": 'h1.text-main', "Description": 'div.description', "Image": 'img',
                     "Price": 'div.product-price'}
        for tag in self.tags:
            cb = QCheckBox(tag, self)
            cb.setStyleSheet("color: white;")
            self.checkboxes[tag] = cb
            tag_layout.addWidget(cb)
        layout.addLayout(tag_layout)

        self.scrape_button = QPushButton("🚀 Start Scraping", self)
        self.scrape_button.setStyleSheet("background: #007BFF; color: white; padding: 10px; border-radius: 5px;")
        self.scrape_button.clicked.connect(self.scrape_website)
        layout.addWidget(self.scrape_button)

        self.progress_bar = QProgressBar(self)
        self.progress_bar.setStyleSheet("QProgressBar { background: #333; color: white; border-radius: 5px; } "
                                        "QProgressBar::chunk { background: #007BFF; }")
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)

        self.table = QTableWidget(self)
        self.table.setStyleSheet("background: #252526; color: white; gridline-color: #333;")
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.table)

        self.save_button = QPushButton("💾 Save to Excel", self)
        self.save_button.setStyleSheet("background: #28A745; color: white; padding: 10px; border-radius: 5px;")
        self.save_button.clicked.connect(self.save_to_excel)
        layout.addWidget(self.save_button)

        self.setLayout(layout)

    def scrape_website(self):
        url = self.url_input.text()
        if not url:
            return

        selected_tags = [tag for tag, cb in self.checkboxes.items() if cb.isChecked()]
        self.table.setColumnCount(len(selected_tags))
        self.table.setHorizontalHeaderLabels(selected_tags)

        self.progress_bar.setValue(0)

        self.scraper_thread = ScraperThread(url, selected_tags, self.tags)
        self.scraper_thread.progress_signal.connect(self.progress_bar.setValue)
        self.scraper_thread.data_signal.connect(self.populate_table)
        self.scraper_thread.start()

    def populate_table(self, data):
        self.table.setRowCount(len(data))
        for i, row in enumerate(data):
            for j, item in enumerate(row):
                self.table.setItem(i, j, QTableWidgetItem(item))
        self.progress_bar.setValue(100)

    def save_to_excel(self):
        if self.table.rowCount() == 0:
            print("No data available to save.")  # Debugging message
            return

        path, _ = QFileDialog.getSaveFileName(self, "Save File", "", "Excel Files (*.xlsx)")

        if not path:
            print("No file selected.")  # Debugging message
            return

        headers = [self.table.horizontalHeaderItem(i).text() for i in range(self.table.columnCount())]
        data = []

        for row in range(self.table.rowCount()):
            row_data = []
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                row_data.append(item.text() if item else "N/A")  # Handle empty cells
            data.append(row_data)

        df = pd.DataFrame(data, columns=headers)

        try:
            df.to_excel(path, index=False)
            print(f"Excel file saved successfully at: {path}")  # Debugging message
        except Exception as e:
            print(f"Error saving file: {e}")  # Print error for debugging


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ScraperApp()
    window.show()
    sys.exit(app.exec())
