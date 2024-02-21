#!/usr/bin/env python
"""
===== main.py =====
This modules serves as the entrypoint of our program.

● usage :
    "python main.py"

● dependencies :
    > Python  (3.11.4)
    > PySide6 (6.6.1)
"""
# pylint: disable=[no-name-in-module, import-error]
import os
import sys
import json
from re import sub
from typing import Optional

from xlsxwriter import Workbook
from requests import post
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QTableWidgetItem,
                               QAbstractItemView, QLineEdit, QTableWidget, QMenuBar,
                               QMessageBox, QApplication, QStyleFactory)

class MainWindow(QMainWindow):
    """TO DO"""
    # pylint: disable=[too-few-public-methods, invalid-name]
    def __init__(self):
        """Class Constructor"""
        super().__init__()

        self.search_bar = QLineEdit()
        self.table_widget = QTableWidget(self)

        self.setMinimumSize(800, 600)
        self.setWindowTitle("Yugioh - Inventaire")

        self.setMenuBar(QMenuBar())
        self.menuBar().addAction("Mettre à jour les données ...")
        self.menuBar().addAction("A Propos")
        self.menuBar().actions()[0].triggered.connect(self.__updateDatas)
        self.menuBar().actions()[1].triggered.connect(self.__about)


        self.setCentralWidget(QWidget())
        self.centralWidget().setLayout(QVBoxLayout())


        self.__searchBox()
        self.__tableView()

        self.show()

    def __searchBox(self):
        """TO DO"""
        self.centralWidget().layout().addWidget(self.search_bar)

        # Connect the textChanged signal to the search function
        self.search_bar.textChanged.connect(self.__searchTable)

    def __searchTable(self, search_text):
        """TO DO"""
        # Filter the table based on the search text
        for row in range(self.table_widget.rowCount()):
            for col in range(self.table_widget.columnCount()):
                item = self.table_widget.item(row, col)
                if item and search_text.lower() in item.text().lower():
                    # Show the row if the search text is found in any cell
                    self.table_widget.setRowHidden(row, False)
                    break
                else:
                    # Hide the row if the search text is not found in any cell
                    self.table_widget.setRowHidden(row, True)

    def __updateDatas(self):
        """TODO"""
        download_data()
        self.__tableView()

    def __tableView(self):
        """TO DO"""
        self.centralWidget().layout().addWidget(self.table_widget)
        self.table_widget.clear()
        data = retrieve_data("sets")

        if not data:
            return

        banned_terms = ["Lost Art", "Shonen Jump", "promotional", "Prize", "Sneak Peek",
                        "collaboration", "participation", "Sweepstakes"]
        filtered_data = [x for x in data if not any(item in x["set_name"] for item in banned_terms)]

        self.table_widget.setRowCount(len(filtered_data))
        self.table_widget.setColumnCount(4)

        for count, card_set in enumerate(filtered_data):
            set_name = card_set["set_name"]
            set_code = card_set["set_code"]
            set_quantity = str(card_set["num_of_cards"])
            if "tcg_date" in card_set:
                set_date = card_set["tcg_date"]

            self.table_widget.setItem(count, 0, QTableWidgetItem(set_name))
            self.table_widget.setItem(count, 1, QTableWidgetItem(set_code))
            self.table_widget.setItem(count, 2, QTableWidgetItem(set_quantity))
            if "tcg_date" in card_set:
                self.table_widget.setItem(count, 3, QTableWidgetItem(set_date))

        for row in range(self.table_widget.rowCount()):
            for col in range(self.table_widget.columnCount()):
                item = self.table_widget.item(row, col)
                if item:
                    item.setFlags(item.flags() ^ Qt.ItemFlag.ItemIsEditable)

        self.table_widget.setHorizontalHeaderLabels(["Nom", "Code", "Cartes", "Date"])
        self.table_widget.verticalHeader().setHidden(True)

        self.table_widget.setSortingEnabled(True)
        self.table_widget.setColumnWidth(0, 520)
        self.table_widget.setColumnWidth(1, 80)
        self.table_widget.setColumnWidth(2, 65)
        self.table_widget.setColumnWidth(3, 100)
        self.table_widget.sortItems(3, order=Qt.SortOrder.DescendingOrder)
        self.table_widget.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)

        self.table_widget.itemDoubleClicked.connect(self.__tableDoubleClick)

    def __tableDoubleClick(self, item: QTableWidgetItem):
        #
        row = item.row()
        name = self.table_widget.item(row, 0).text()
        create_sheet(name)

    def __show_popup(self, message: str, icon_type: str, buttons: list[str],
               title: Optional[str] = None) -> str:
        """
        Display a popup dialog with a message, icon and buttons.

        Args:
            message {str} : the text displayed in the popup.
            icon_type {str} : the type of icon. Can be [Warning][Question][Error][Info].
            buttons {list[str]} : a list of str that will be use to populate the popup with buttons.
            title {Optional[str]} : the title for the popup.

        Returns:
            - str : the text of the button clicked.
        """
        def get_buttons(buttons: list) -> QMessageBox.StandardButton:
            """
            Get the standard buttons corresponding to the specified button types.

            Args:
                - buttons (list): The list of buttons to add.

            Returns:
                - QMessageBox.StandardButton: The standard buttons corresponding to the
                    specified button types.
            """
            button_mapping = {
                "OK": QMessageBox.StandardButton.Ok,
                "CANCEL": QMessageBox.StandardButton.Cancel,
                "YES": QMessageBox.StandardButton.Yes,
                "NO": QMessageBox.StandardButton.No,
                "SAVE": QMessageBox.StandardButton.Save,
                "DISCARD": QMessageBox.StandardButton.Discard,
                "CLOSE": QMessageBox.StandardButton.Close,
                "RETRY": QMessageBox.StandardButton.Retry,
                "IGNORE": QMessageBox.StandardButton.Ignore
            }

            standard_buttons = QMessageBox.StandardButton(QMessageBox.StandardButton.NoButton)

            for button in buttons:
                standard_buttons |= button_mapping.get(button, QMessageBox.StandardButton.NoButton)

            return standard_buttons

        icon_mapping = {
            "WARNING": QMessageBox.Icon.Warning,
            "QUESTION": QMessageBox.Icon.Question,
            "ERROR": QMessageBox.Icon.Critical,
            "INFO": QMessageBox.Icon.Information
        }

        popup = QMessageBox()
        popup.setIcon(icon_mapping.get(icon_type, QMessageBox.Icon.NoIcon))
        popup.setWindowFlags(Qt.WindowType.WindowStaysOnTopHint)
        popup.setStandardButtons(get_buttons(buttons))
        popup.setTextFormat(Qt.TextFormat.RichText)
        popup.setWindowTitle(title or icon_type)
        popup.setText(message)
        popup.exec()
        return popup.clickedButton().text()

    def __about(self):
        """display the about message"""
        self.__show_popup("""
            <h1>Yugioh - Inventaire</h1>
            <br>Author  : Costantin Hereiti
            <br>License : <a href="https://www.gnu.org/licenses/lgpl-3.0.en.html">LGPL v3</a>
            <br>Python  : 3.11.4 - 64Bit
            <br>PySide6 : 6.6.1
            """, "INFO", ["OK"], "About")


def download_data():
    """Retrieve data from the API and save them as local file"""
    api = {
        "cards_en": "https://db.ygoprodeck.com/api/v7/cardinfo.php",
        "cards_fr": "https://db.ygoprodeck.com/api/v7/cardinfo.php?language=fr",
        "sets": "https://db.ygoprodeck.com/api/v7/cardsets.php"
    }

    for key, link in api.items():
        response = post(link, timeout=1000)

        # Define the path to the configuration file's directory
        document_dir = os.path.join(os.path.expanduser("~"), "Documents")
        data_dir = os.path.join(document_dir, "Yugioh - Datas")

        # Create the directory if it doesn't exist
        if not os.path.exists(data_dir):
            os.makedirs(data_dir)

        if response.status_code == 200:
            with open(os.path.join(data_dir, f"{key}.json"), "w", encoding="utf-8") as file:
                file.write(json.dumps(response.json()))

        else:
            print(f"Error: {response.status_code}")
            print(response.text)

def retrieve_data(key):
    """Retrieve data from local files"""
    try:
        # Define the path to the configuration file's directory
        document_dir = os.path.join(os.path.expanduser("~"), "Documents")
        data_dir = os.path.join(document_dir, "Yugioh - Datas")

        data_path = os.path.join(data_dir, f"{key}.json")
        with open(data_path, "r", encoding="utf-8") as file:
            data = json.load(file)

        return data

    except FileNotFoundError:
        return None

def create_sheet(set_name: str):
    """
    Create a sheet from the corresponding extension.

    Parameters:
    filename (str) :: the name for the excel file.
    """
    # pylint: disable=too-many-locals

    # Create the excel file and page
    filename = sub(r'[\\/:\*\?"<>\| ]', '_', set_name)
    filename = sub(r'_+', '_', filename)

    sheet_name = sub(r'_', ' ', filename)
    if len(sheet_name) > 22:
        sheet_name = sheet_name[:22]
        sheet_name = sheet_name + '...'

    workbook = Workbook(f"{filename}.xlsx")
    worksheet = workbook.add_worksheet(sheet_name)

    # Initialize the format for the header
    header_format = workbook.add_format({
        "bold": True,
        "align": "center",
        "valign": "vcenter",
        "fg_color": "#D7E4BC",
        "border": 1,
    })

    # Fill the header and apply the format
    worksheet.write('A1', 'Nom EN', header_format)
    worksheet.write('B1', 'Nom FR', header_format)
    worksheet.write('C1', 'Code', header_format)
    worksheet.write('D1', 'Rareté', header_format)
    worksheet.write('E1', 'Quantité', header_format)

    # Retrieve cards from the downloaded datas
    data = retrieve_data("cards_en")
    data_fr = retrieve_data("cards_fr")

    # Check if the card is present in the set using the set name
    unsorted_cards_list = []
    for card in data["data"]:
        if "card_sets" in card:
            for card_set in card["card_sets"]:
                if set_name == card_set["set_name"]:
                    card_set: dict[str, str]

                    # Format the set code
                    prefix, suffix = card_set["set_code"].split("-")
                    suffix = suffix.replace("EN", "FR")
                    set_code = "-".join([prefix, suffix])

                    # Format the rarity code
                    if card_set["set_rarity"] == "Quarter Century Secret Rare":
                        rarity = "QCSR"
                    else:
                        rarity = card_set["set_rarity_code"].replace("(", "").replace(")", "")

                    # Add a tuple with the name, rarity and code
                    unsorted_cards_list.append((card["name"], set_code, rarity))

    # Sort the list by set code
    sorted_cards_list = sorted(unsorted_cards_list, key=lambda x: x[1])

    # Initialize the format for every other items
    center_format = workbook.add_format({"align": "center"})

    # Loop through the retrieved cards list and place them in the table
    for index, (name, code, rarity) in enumerate(sorted_cards_list):
        worksheet.write(f"A{index + 2}", name, center_format)
        for card in data_fr["data"]:
            if card["name_en"] == name:
                worksheet.write(f"B{index + 2}", card["name"], center_format)
        worksheet.write(f"C{index + 2}", code, center_format)
        worksheet.write(f"D{index + 2}", rarity, center_format)

    # Finalize sheet formatting
    worksheet.freeze_panes(1, 0)
    worksheet.autofit()
    workbook.close()

    # Inform the user that the operation has ended
    message_box = QMessageBox()
    message_box.setWindowTitle("Yugioh Datas")
    message_box.setText("Fichier Excel créé avec succès.")
    message_box.setIcon(QMessageBox.Icon.Information)
    message_box.exec()

if __name__ == "__main__":
    app = QApplication([])
    app.setStyle(QStyleFactory.create("fusion"))

    program = MainWindow()

    sys.exit(app.exec())
