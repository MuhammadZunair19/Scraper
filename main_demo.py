import sys
import time
import pandas as pd
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLineEdit, QLabel, QTableWidget, \
    QTableWidgetItem, QCheckBox, QHBoxLayout, QFileDialog
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup


class ScraperApp(QWidget):
    def __init__(self):
        super().__init__()
        self.table = None
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Dynamic Web Scraper with Infinite Scroll")
        self.setGeometry(100, 100, 800, 600)

        layout = QVBoxLayout()

        self.url_input = QLineEdit(self)
        self.url_input.setPlaceholderText("Enter website URL")
        layout.addWidget(self.url_input)

        # Checkboxes for selecting tags
        self.checkboxes = {}
        tag_layout = QHBoxLayout()
        self.tags = {"Title": 'h1.text-main', "Description": 'div.description', "Image": 'img',
                     "Price": 'div.product-price'}
        for tag in self.tags:
            cb = QCheckBox(tag, self)
            self.checkboxes[tag] = cb
            tag_layout.addWidget(cb)
        layout.addLayout(tag_layout)

        self.scrape_button = QPushButton("Scrape", self)
        self.scrape_button.clicked.connect(self.scrape_website)
        layout.addWidget(self.scrape_button)

        self.table = QTableWidget(self)
        layout.addWidget(self.table)

        self.save_button = QPushButton("Save to Excel", self)
        self.save_button.clicked.connect(self.save_to_excel)
        layout.addWidget(self.save_button)

        self.setLayout(layout)

    def scroll_page(self, driver):
        """ Scrolls the page dynamically until all products are loaded. """
        SCROLL_PAUSE_TIME = 2  # Adjust pause time based on site speed

        last_height = driver.execute_script("return document.body.scrollHeight")
        while True:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(SCROLL_PAUSE_TIME)  # Allow time for loading

            new_height = driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                print("[INFO] Scrolling completed - No more content loaded.")
                break  # Stop scrolling if no new content appears
            last_height = new_height

    def scrape_website(self):
        url = self.url_input.text().strip()
        if not url:
            return

        options = Options()
        options.add_argument("--headless")
        driver = webdriver.Chrome(options=options)
        driver.get(url)

        self.scroll_page(driver)  # Scroll dynamically until all products load

        soup = BeautifulSoup(driver.page_source, 'html.parser')
        articles = soup.find_all('article', class_='store store-item border-accent single-image')
        driver.quit()

        selected_tags = [tag for tag, cb in self.checkboxes.items() if cb.isChecked()]
        self.table.setColumnCount(len(selected_tags))
        self.table.setHorizontalHeaderLabels(selected_tags)

        data = []
        for article in articles:
            row = []
            for tag in selected_tags:
                element = article.select_one(self.tags[tag])
                text = element.text.strip() if element and tag != "Image" else (
                    element['src'] if element and element.has_attr('src') else "")
                row.append(text)
            data.append(row)

        self.table.setRowCount(len(data))
        for i, row in enumerate(data):
            for j, item in enumerate(row):
                self.table.setItem(i, j, QTableWidgetItem(item))

    def save_to_excel(self):
        path, _ = QFileDialog.getSaveFileName(self, "Save File", "scraped_data.xlsx", "Excel Files (*.xlsx)")
        if path:
            headers = [self.table.horizontalHeaderItem(i).text() for i in range(self.table.columnCount())]
            data = [[self.table.item(row, col).text() for col in range(self.table.columnCount())] for row in
                    range(self.table.rowCount())]
            df = pd.DataFrame(data, columns=headers)
            df.to_excel(path, index=False)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ScraperApp()
    window.show()
    sys.exit(app.exec())
