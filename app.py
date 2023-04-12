import os
import re
import math
from numpy import array, around
from scipy.interpolate import RegularGridInterpolator, interp1d

from PyQt6.QtCore import QSize, Qt, QRegularExpression
from PyQt6.QtGui import QRegularExpressionValidator, QColor, QFont, QIcon, QRegularExpressionValidator
from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow,
    QMessageBox,
    QCheckBox,
    QPushButton,
    QLabel,
    QLineEdit,
    QGridLayout,
    QTableWidget,
    QVBoxLayout,
    QComboBox,
    QGroupBox,
    QRadioButton,
    QComboBox,
    QHBoxLayout,
    QWidget,
    QTabWidget,
    QTableWidgetItem,
)

from constants import CONSTANTS


basedir = os.path.dirname(__file__)


try:
    from ctypes import windll  # Only exists on Windows.
    myappid = 'akudjatechnology.air-system.1.0.0'
    windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except ImportError:
    pass


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(CONSTANTS.APP_TITLE)
        self.box_style = "QGroupBox::title {color: blue;}"
        self.last_floor = False

        self.tab_widget = QTabWidget(self)
        self.setCentralWidget(self.tab_widget)

        self.tab_widget.addTab(self.create_tab1_content(), CONSTANTS.TAB1_TITLE)
        self.tab_widget.addTab(self.create_tab2_content(), CONSTANTS.TAB2_TITLE)

        self.showMaximized()


    def create_tab1_content(self) -> object:
        _widget = QWidget()
        _layout = QVBoxLayout()

        hbox1 = QHBoxLayout()
        hbox1.addWidget(self.create_init_data_box())
        hbox1.addWidget(self.create_sputnik_table_box())

        hbox2 = QHBoxLayout()
        hbox2.addWidget(self.create_deflector_checkbox())
        hbox2.addWidget(self.create_buttons_box())
        
        hbox3 = QVBoxLayout()
        hbox3.addWidget(self.create_main_table_box())

        _layout.addLayout(hbox1)
        _layout.addLayout(hbox2)
        _layout.addLayout(hbox3)
        _widget.setLayout(_layout)

        return _widget


    def create_tab2_content(self) -> object:
        _widget = QWidget()
        _layout = QVBoxLayout()
        hbox1 = QHBoxLayout()
        hbox1.addWidget(self.create_deflector_calculation())

        _layout.addLayout(hbox1)
        _widget.setLayout(_layout)
        return _widget


    def create_init_data_box(self) -> object:
        _box = QGroupBox(CONSTANTS.INIT_DATA.TITLE)
        style = self.box_style
        _box.setStyleSheet(style)
        _box.setFixedWidth(445)

        self.init_data_layout = QGridLayout()
        self.init_data_layout.setVerticalSpacing(10)
        labels = CONSTANTS.INIT_DATA.LABELS
        for i in range(len(labels)):
            label_0 = QLabel(labels[i][0])
            label_0.setFixedHeight(CONSTANTS.INIT_DATA.LINE_HEIGHT)
            line_edit = QLineEdit()
            line_edit.setAlignment(Qt.AlignmentFlag.AlignCenter)
            line_edit.setStyleSheet('QLineEdit {background-color: %s}' % QColor(229, 255, 204).name())
            line_edit.setFixedWidth(CONSTANTS.INIT_DATA.INPUT_WIDTH)
            line_edit.setFixedHeight(CONSTANTS.INIT_DATA.LINE_HEIGHT)
            label_1 = QLabel(labels[i][1])
            label_1.setFixedHeight(CONSTANTS.INIT_DATA.LINE_HEIGHT)
            self.init_data_layout.addWidget(label_0, i, 0)
            self.init_data_layout.addWidget(line_edit, i, 1)
            self.init_data_layout.addWidget(label_1, i, 2)

        temperature_item = self.init_data_layout.itemAtPosition(0, 1)
        self.temperature_widget = temperature_item.widget()
        self.temperature_widget.setObjectName('temperature')
        temperature_regex = QRegularExpression("^(?:\d|[12]\d|30)(?:\.\d)?$")
        temperature_validator = QRegularExpressionValidator(temperature_regex)
        self.temperature_widget.setValidator(temperature_validator)
        self.temperature_widget.textChanged.connect(self.calculate_gravi_pressure)
        self.temperature_widget.textChanged.connect(self.calculate_dynamic)

        surface_item = self.init_data_layout.itemAtPosition(1, 1)
        self.surface_widget = surface_item.widget()
        self.surface_widget.setObjectName('surface')
        surface_regex = QRegularExpression("^(?:[0-9]|[1-9]\d|100)(?:\.\d{1,3})?$")
        surface_validator = QRegularExpressionValidator(surface_regex)
        self.surface_widget.setValidator(surface_validator)
        self.surface_widget.textChanged.connect(self.calculate_specific_pressure_loss)


        floor_height_item = self.init_data_layout.itemAtPosition(2, 1)
        self.floor_height_widget = floor_height_item.widget()
        floor_height_regex = QRegularExpression("^(?:[1-9]\d?|100)(?:\.\d{1,2})?$")
        floor_height_validator = QRegularExpressionValidator(floor_height_regex)
        self.floor_height_widget.setValidator(floor_height_validator)
        self.floor_height_widget.textChanged.connect(self.set_base_floor_height_in_main_table)

        self.channel_height_item = self.init_data_layout.itemAtPosition(3, 1)
        self.channel_height_widget = self.channel_height_item.widget()
        self.channel_height_widget.setObjectName('channel_height')
        channel_height_regex = QRegularExpression("^(?:[1-9]|[1-9]\d|100)(?:\.\d{1,2})?$")
        channel_height_validator = QRegularExpressionValidator(channel_height_regex)
        self.channel_height_widget.setValidator(channel_height_validator)
        self.channel_height_widget.textChanged.connect(self.calculate_height)

        klapan_label = QLabel(CONSTANTS.INIT_DATA.KLAPAN_LABEL)
        self.klapan_widget = QComboBox()
        self.klapan_widget.setStyleSheet('QLineEdit {background-color: %s}' % QColor(229, 255, 204).name())
        self.klapan_widget.setFixedHeight(CONSTANTS.INIT_DATA.LINE_HEIGHT)
        self.klapan_widget.addItems(CONSTANTS.INIT_DATA.KLAPAN_ITEMS.keys())
        self.klapan_widget.setFixedWidth(CONSTANTS.INIT_DATA.INPUT_WIDTH)

        self.klapan_widget_value = CONSTANTS.INIT_DATA.KLAPAN_ITEMS.get(self.klapan_widget.currentText())
        self.klapan_air_flow_label = QLabel(f'{self.klapan_widget_value} м<sup>3</sup>/ч')

        self.klapan_widget.currentTextChanged.connect(self.set_klapan_air_flow_in_label)
        self.klapan_widget.currentTextChanged.connect(self.calculate_klapan_pressure_loss)
        self.klapan_widget.currentTextChanged.connect(self.activate_klapan_input)

        klapan_input_label_1 = QLabel(CONSTANTS.INIT_DATA.KLAPAN_INPUT_LABEL_1)
        self.klapan_input = QLineEdit()
        klapan_input_label_2 = QLabel('м<sup>3</sup>/ч')
        self.klapan_input.setFixedWidth(CONSTANTS.INIT_DATA.INPUT_WIDTH)
        self.klapan_input.setFixedHeight(CONSTANTS.INIT_DATA.LINE_HEIGHT)
        self.klapan_input.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.klapan_input.setStyleSheet('QLineEdit {background-color: %s}' % QColor(224, 224, 224).name())
        self.klapan_input.setDisabled(True)
        klapan_input_regex = QRegularExpression("^(?:[1-9]|[1-9]\d|100)(?:)?$")
        klapan_input_validator = QRegularExpressionValidator(klapan_input_regex)
        self.klapan_input.setValidator(klapan_input_validator)
        self.klapan_input.setToolTip(CONSTANTS.INIT_DATA.KLAPAN_INPUT_TOOLTIP)
        self.klapan_input.textChanged.connect(self.calculate_klapan_pressure_loss)
        

        self.init_data_layout.addWidget(klapan_label, 4, 0)
        self.init_data_layout.addWidget(self.klapan_widget, 4, 1, 1, 2)
        self.init_data_layout.addWidget(self.klapan_air_flow_label, 4, 2)
        self.init_data_layout.addWidget(klapan_input_label_1, 5, 0)
        self.init_data_layout.addWidget(self.klapan_input, 5, 1)
        self.init_data_layout.addWidget(klapan_input_label_2, 5, 2)
        self.init_data_layout.setColumnStretch(1, 1)
        self.init_data_layout.setColumnStretch(2, 1)
        _box.setLayout(self.init_data_layout)
        return _box


    def create_deflector_checkbox(self) -> object:
        _widget = QWidget()
        _layout = QHBoxLayout()

        label = QLabel('Добавить дефлектор')
        self.activate_deflector = QCheckBox()
        self.activate_deflector.setDisabled(True)
        self.activate_deflector.setChecked(False)
        self.activate_deflector.stateChanged.connect(self.show_deflector_column_in_main_table)
        self.activate_deflector.stateChanged.connect(self.set_deflector_pressure_in_main_table)
        self.activate_deflector.stateChanged.connect(self.calculate_available_pressure)

        _layout.addWidget(label)
        _layout.addWidget(self.activate_deflector)
        _widget.setLayout(_layout)
        return _widget


    def create_buttons_box(self) -> object:
        _widget = QWidget()
        _layout = QHBoxLayout()
        _layout.addStretch()

        self.last_floor_row_button = QPushButton()
        self.last_floor_row_button.setText(CONSTANTS.BUTTONS.LAST_FLOOR_BUTTON_TITLE)
        _layout.addWidget(self.last_floor_row_button)
        self.last_floor_row_button.clicked.connect(self.add_last_floor)

        self.add_row_button = QPushButton()
        self.add_row_button.setText(CONSTANTS.BUTTONS.ADD_BUTTON_TITLE)
        self.add_row_button.setStyleSheet('QPushButton:hover {background-color: rgb(102, 204, 0)}')
        # add_icon = QIcon('add_row.png')
        # self.add_row_button.setIcon(add_icon)
        # self.add_row_button.setIconSize(QSize(25, 25))
        self.add_row_button.setToolTip(CONSTANTS.BUTTONS.ADD_BUTTON_TOOLTIP)
        _layout.addWidget(self.add_row_button)
        self.add_row_button.clicked.connect(self.add_row)
        self.add_row_button.clicked.connect(self.set_floor_number_in_main_table)
        self.add_row_button.clicked.connect(self._join_deflector_column_cells_in_main_table)
        # self.add_row_button.clicked.connect(self._set_deflector_pressure_in_main_table)

        self.delete_row_button = QPushButton()
        self.delete_row_button.setText(CONSTANTS.BUTTONS.DELETE_BUTTON_TITLE)
        self.delete_row_button.setStyleSheet('QPushButton:hover {background-color: rgb(255, 153, 153)}')
        # delete_icon = QIcon('delete_row.png')
        # self.delete_row_button.setIcon(delete_icon)
        # self.delete_row_button.setIconSize(QSize(25, 25))
        self.delete_row_button.setToolTip(CONSTANTS.BUTTONS.DELETE_BUTTON_TOOLTIP)
        _layout.addWidget(self.delete_row_button)
        self.delete_row_button.clicked.connect(self.delete_row)
        self.delete_row_button.clicked.connect(self.update_full_air_flow_in_deflector_after_delete_row)
        self.delete_row_button.clicked.connect(self.update_air_flow_column_in_main_table_after_delete_row)
        self.delete_row_button.clicked.connect(self.set_floor_number_in_main_table)

        _widget.setLayout(_layout)
        return _widget


    def create_sputnik_table_box(self) -> object:
        _box = QGroupBox(CONSTANTS.SPUTNIK_TABLE.TITLE)
        style = self.box_style
        _box.setStyleSheet(style)
        _box.setFixedHeight(298)

        _layout = QVBoxLayout()

        self.sputnik_table = QTableWidget(5, 15)
        self.sputnik_table.verticalHeader().setVisible(False)
        self.sputnik_table.horizontalHeader().setStretchLastSection(True)
        self.sputnik_table.horizontalHeader().setStyleSheet("QHeaderView::section { background-color: #F3F3F3 }")
        self.sputnik_table.setStyleSheet("QTableWidget { border: 2px solid grey; }")
        self.sputnik_table.setHorizontalHeaderLabels(CONSTANTS.SPUTNIK_TABLE.HEADER)
        self.sputnik_table.setObjectName(CONSTANTS.SPUTNIK_TABLE.NAME)

        num_rows = self.sputnik_table.rowCount()
        num_cols = self.sputnik_table.columnCount()

        # устанавливаем виджет QTableWidgetItem для всех ячеек таблицы и отключаем редактирование
        for row in range(num_rows):
            for col in range(num_cols):
                self.sputnik_table.setItem(row, col, QTableWidgetItem())
                self.sputnik_table.item(row, col).setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                editable = (
                    (0, 1),
                    (1, 1),
                    (1, 2),
                    (1, 3),
                    (1, 4),
                    (1, 11),
                    (3, 1),
                    (3, 2),
                    (3, 3),
                    (3, 4),
                    (3, 11),
                )
                if (row, col) not in editable:
                    self.sputnik_table.item(row, col).setFlags(Qt.ItemFlag.ItemIsEnabled)
                else:
                    self.sputnik_table.item(row, col).setBackground(QColor(229, 255, 204))

        for row in range(num_rows):
            self.sputnik_table.setRowHeight(row, 29)
            match row:
                case 2 | 4:
                    self.sputnik_table.setSpan(row, 0, 1, 13)
                    # self.sputnik_table.item(row, 0).setBackground(QColor(224, 224, 224))
                case 0:
                    self.sputnik_table.setSpan(row, 2, 1, 11)
                    # self.sputnik_table.item(row, 2).setBackground(QColor(224, 224, 224))
                case 1 | 3:
                    self.sputnik_table.setSpan(row, 14, 2, 1)

        for i in range(num_cols):
            match i:
                case 3 | 4:
                    self.sputnik_table.setColumnWidth(i, 110)
                case _:
                    self.sputnik_table.setColumnWidth(i, 72)
                    
            # if i in (1, 2, 3, 4):
            #     self.sputnik_table.item(1, i).setBackground(QColor(229, 255, 204))
            #     # self.sputnik_table.item(1, i).setFlags(~Qt.ItemFlag.NoItemFlags)
            #     self.sputnik_table.item(3, i).setBackground(QColor(229, 255, 204))
            #     # self.sputnik_table.item(3, i).setFlags(Qt.ItemFlag.NoItemFlags)

        # установка заливки для редактируемых столбцов
        self.sputnik_table.item(0, 1).setBackground(QColor(229, 255, 204))


        # заполняем таблицу
        self.sputnik_table.item(0, 0).setText(self.klapan_widget.currentText())
        self.sputnik_table.item(1, 0).setText('1-2')
        self.sputnik_table.item(3, 0).setText('1*-2*')

        # размещаем радиокнопки
        for row in (1, 3):
            widget = QWidget()
            layout = QHBoxLayout(widget)
            layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
            radio_button = QRadioButton(widget)
            layout.addWidget(radio_button)
            widget.setLayout(layout)
            self.sputnik_table.setCellWidget(row, 14, widget)

        self.radio_button1 = self.sputnik_table.cellWidget(1, 14).findChild(QRadioButton)
        self.radio_button2 = self.sputnik_table.cellWidget(3, 14).findChild(QRadioButton)
        self.radio_button1.setChecked(True)

        self.radio_button1.setToolTip(CONSTANTS.SPUTNIK_TABLE.RADIO_TOOLTIP_1)
        self.radio_button2.setToolTip(CONSTANTS.SPUTNIK_TABLE.RADIO_TOOLTIP_2)

        # обработчики
        self.radio_button1.clicked.connect(self.set_sputnik_airflow_in_main_table_by_radiobutton_1)
        self.radio_button2.clicked.connect(self.set_sputnik_airflow_in_main_table_by_radiobutton_2)

        self.radio_button2.clicked.connect(self.uncheck_radio_button_1)
        self.radio_button1.clicked.connect(self.uncheck_radio_button_2)

        self.radio_button1.clicked.connect(self.calculate_kms_by_radiobutton_1)
        self.radio_button2.clicked.connect(self.calculate_kms_by_radiobutton_2)
        self.radio_button1.clicked.connect(self.calculate_branch_pressure_by_radiobutton_1)
        self.radio_button2.clicked.connect(self.calculate_branch_pressure_by_radiobutton_2)
        self.radio_button1.clicked.connect(self.calculate_full_pressure_by_radiobutton_1)
        self.radio_button2.clicked.connect(self.calculate_full_pressure_by_radiobutton_2)

        self.sputnik_table.cellChanged.connect(self.set_sputnik_airflow_in_main_table)
        self.sputnik_table.cellChanged.connect(self.set_deflector_pressure_in_main_table)

        self.sputnik_table.cellChanged.connect(self.calculate_klapan_pressure_loss)
        self.sputnik_table.cellChanged.connect(self.calculate_air_velocity)
        self.sputnik_table.cellChanged.connect(self.calculate_diameter)
        self.sputnik_table.cellChanged.connect(self.calculate_dynamic)
        self.sputnik_table.cellChanged.connect(self.calculate_specific_pressure_loss)
        self.sputnik_table.cellChanged.connect(self.calculate_m)
        self.sputnik_table.cellChanged.connect(self.calculate_linear_pressure_loss)
        self.sputnik_table.cellChanged.connect(self.calculate_local_pressure_loss)
        self.sputnik_table.cellChanged.connect(self.calculate_sputnik_full_pressure_loss)
        self.sputnik_table.cellChanged.connect(self.calculate_kms)
        self.sputnik_table.cellChanged.connect(self.calculate_result_sputnik_pressure)
        self.sputnik_table.cellChanged.connect(self.calculate_branch_pressure)
        self.sputnik_table.cellChanged.connect(self.calculate_full_pressure)

        self.sputnik_table.itemChanged.connect(self.validate_input_data_in_tables)

        _layout.addWidget(self.sputnik_table)
        _box.setLayout(_layout)
        return _box


    def create_main_table_box(self) -> object:
        _box = QGroupBox(CONSTANTS.MAIN_TABLE.TITLE)
        style = self.box_style
        _box.setStyleSheet(style)

        _layout = QVBoxLayout()

        self.main_table = QTableWidget(5, 22)
        self.main_table.horizontalHeader().setStretchLastSection(True)
        self.main_table.setHorizontalHeaderLabels(CONSTANTS.MAIN_TABLE.HEADER)
        self.main_table.horizontalHeader().setStyleSheet("QHeaderView::section { background-color: #F3F3F3 }")
        self.main_table.setStyleSheet("QTableWidget { border: 2px solid grey; }")
        self.main_table.verticalHeader().setVisible(False)
        self.main_table.setObjectName(CONSTANTS.MAIN_TABLE.NAME)

        num_rows = self.main_table.rowCount()
        num_cols = self.main_table.columnCount()

        # устанавливаем виджет QTableWidgetItem для всех ячеек таблицы
        for row in range(num_rows):
            for col in range(num_cols):
                self.main_table.setItem(row, col, QTableWidgetItem())
                self.main_table.item(row, col).setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.main_table.item(row, 0).setText(str(row+1))

        # установка ширины столбцов
        for i in range(num_cols):
            match i:
                case 0:
                    self.main_table.setColumnWidth(i, 50)
                case 1 | 2 | 3 | 12 | 13 | 14 | 15 | 16 | 17 | 18 | 19 | 20:
                    self.main_table.setColumnWidth(i, 70)
                case 4 | 5 | 6 | 7 | 8 | 9:
                    self.main_table.setColumnWidth(i, 60)
                case 10 | 11:
                    self.main_table.setColumnWidth(i, 110)

            # установка заливки для редактируемых столбцов
            if i in (1, 2, 10, 11):
                for row in range(num_rows):
                    self.main_table.item(row, i).setBackground(QColor(229, 255, 204))

        # self.main_table.setSpan(0, 6, 0, 0)
        self.main_table.setColumnHidden(6, True)

        # обработчики
        self.main_table.cellChanged.connect(self.update_result)
        self.main_table.cellChanged.connect(self.update_height_column_in_main_table)

        self.main_table.cellChanged.connect(self.calculate_gravi_pressure)
        self.main_table.cellChanged.connect(self.calculate_air_velocity)
        self.main_table.cellChanged.connect(self.calculate_dynamic)
        self.main_table.cellChanged.connect(self.calculate_diameter)
        self.main_table.cellChanged.connect(self.calculate_specific_pressure_loss)
        self.main_table.cellChanged.connect(self.calculate_m)
        self.main_table.cellChanged.connect(self.calculate_linear_pressure_loss)
        self.main_table.cellChanged.connect(self.calculate_kms)
        self.main_table.cellChanged.connect(self.calculate_pass_pressure)
        self.main_table.cellChanged.connect(self.calculate_branch_pressure)
        self.main_table.cellChanged.connect(self.calculate_available_pressure)
        self.main_table.cellChanged.connect(self.calculate_height)
        self.main_table.cellChanged.connect(self.calculate_full_pressure)
        self.main_table.cellChanged.connect(self.set_full_air_flow_in_deflector)

        self.main_table.itemChanged.connect(self.validate_input_data_in_tables)

        _layout.addWidget(self.main_table)
        _box.setLayout(_layout)
        return _box


    def validate_input_data_in_tables(self, item) -> None:
        table = self.sender().objectName()
        column = item.column()
        # Allows values: whole numbers 0...2000
        dimensions_pattern = pattern = r'^([1-9]\d{0,2}|1\d{3}|2000)?$'
        # Allows values: 0...100 with or without one | two digit after separator
        length_height_kms_pattern = pattern = r'^(?:[0-9]|[1-9]\d|100)(?:\.\d{1,3})?$'
        match table:
            case 'sputnik':
                if column in (1, 2, 3, 4, 11):
                    match column:
                        
                        case 1:
                            # Allows values: 0...200 with or without one digit after separator
                            pattern = r'^(?:\d{1,2}|1\d{2}|2\d{2})(?:\.\d)?$'
                        case 2 | 11:
                            pattern = length_height_kms_pattern
                        case 3 | 4:
                            pattern = dimensions_pattern

                    if not re.match(pattern, item.text()):
                        item.setText('')
                    else:
                        item.setText(item.text())

            case 'main':
                if column in (1, 10, 11):
                    match column:
                        case 1:
                            length_height_kms_pattern
                        case 10 | 11:
                            pattern = dimensions_pattern

                    if not re.match(pattern, item.text()):
                        item.setText('')
                    else:
                        item.setText(item.text())


    def delete_row(self) -> None:
        num_rows = self.main_table.rowCount()
        if num_rows > 1:
            selected_row = self.main_table.currentRow()
            self.main_table.removeRow(selected_row)


    def add_row(self) -> None:
        if not self.last_floor:
            num_rows = self.main_table.rowCount()
            num_cols = self.main_table.columnCount()
            self.main_table.insertRow(num_rows)

            # вставка номера строки
            item = QTableWidgetItem(str(num_rows + 1))
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

            green_cols = (1, 2, 10, 11)
            # добавление виджета в расчетные ячейки
            for col in range(num_cols):
                if col not in green_cols or col not in (0, 3):
                    self.main_table.setItem(num_rows, col, QTableWidgetItem())

            # заливка зеленым ячеек для ввода
            for col in green_cols:
                self.main_table.setItem(num_rows, col, QTableWidgetItem())
                self.main_table.item(num_rows, col).setBackground(QColor(229, 255, 204))
                self.main_table.item(num_rows, col).setTextAlignment(Qt.AlignmentFlag.AlignCenter)

            # вставка высоты этажа
            if self.floor_height_widget.text() != '':
                self.main_table.item(num_rows, 1).setText("{:.2f}".format(float(self.floor_height_widget.text())))
                self.main_table.item(num_rows, 1).setTextAlignment(Qt.AlignmentFlag.AlignCenter)

            # установка расхода воздуха
            init_flow = self.main_table.item(0, 3)
            last_flow = self.main_table.item(num_rows-1, 3)
            if all([init_flow is not None, last_flow is not None]) and all([init_flow.text() != '', last_flow.text() != '']):
                flow = int(init_flow.text()) + int(last_flow.text())
                flow = str(flow)
                self.main_table.setItem(num_rows, 3, QTableWidgetItem())
                self.main_table.item(num_rows, 3).setText(flow)
                self.main_table.item(num_rows, 3).setTextAlignment(Qt.AlignmentFlag.AlignCenter)

            self.main_table.setItem(num_rows, 0, item)
        # self.set_row_columns_in_not_editable_mode(num_rows)
        else:
            pass


    def add_last_floor(self) -> None:
        num_rows = self.main_table.rowCount()
        self.main_table.insertRow(num_rows)
        self.main_table.setItem(num_rows, 0, QTableWidgetItem('Последний'))
        self.main_table.item(num_rows, 0).setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.last_floor = True


    def set_klapan_air_flow_in_label(self, value) -> None:
        klapan_flow = CONSTANTS.INIT_DATA.KLAPAN_ITEMS.get(value)
        self.klapan_air_flow_label.setText(f'{klapan_flow} м<sup>3</sup>/ч')
        self.sputnik_table.item(0, 0).setText(value)


    # def set_row_columns_in_not_editable_mode(self, row) -> None:
    #     columns = (0, 1, 2, 10, 11)
    #     num_cols = self.main_table.columnCount()
    #     for col in range(num_cols):
    #         if col not in columns:
    #             if self.main_table.item(row, col) is not None:
    #                 # self.main_table.setItem(row, col, QTableWidgetItem())
    #                 item = self.main_table.item(row, col)
    #                 item.setFlags(item.flags() ^ Qt.ItemFlag.ItemIsEditable)


    def update_result(self, row, column) -> None:
        if column in (7, 20):
            available_pressure = self.main_table.item(row, 7)
            full_pressure = self.main_table.item(row, 20)

            if all([available_pressure, full_pressure]) and all([available_pressure.text() != '', full_pressure.text() != '']):
                available_pressure = float(available_pressure.text())
                full_pressure = float(full_pressure.text())

                if available_pressure > full_pressure:
                    result = 'OK'
                    item = QTableWidgetItem(result)
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    self.main_table.setItem(row, 21, item)
                    self.main_table.item(row, 21).setBackground(QColor(229, 255, 204))
                else:
                    result = 'Тяги нет!'
                    item = QTableWidgetItem(result)
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    self.main_table.setItem(row, 21, item)
                    self.main_table.item(row, 21).setBackground(QColor(255, 153, 153))
            else:
                result = ''
                item = QTableWidgetItem(result)
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.main_table.setItem(row, 21, item)
                self.main_table.item(row, 21).setBackground(QColor(255, 255, 255))


    def calculate_gravi_pressure(self, row=False, column=False) -> None:
        sender = self.sender()
        if isinstance(sender, QTableWidget):
            # если изменения данных произошли в таблице
            if column == 4:
                temperature = self.temperature_widget.text()
                item_4 = self.main_table.item(row, 4)
                if (item_4.text() != '' and (temperature or temperature != '')):
                    value_4 = float(item_4.text())
                    temperature = float(temperature)
                    pressure = 0.9 * CONSTANTS.ACCELERATION_OF_GRAVITY * value_4 * ((353 / (273 + 5)) - (353 / (273 + temperature)))
                    pressure = "{:.2f}".format(round(pressure, 2))
                    self.main_table.item(row, 5).setText(pressure)
                    self.main_table.item(row, 5).setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                else:
                    self.main_table.setItem(row, 5, QTableWidgetItem())
                    self.main_table.item(row, 5).setText('')
        else:
            # если изменилась температура воздуха
            temperature = row
            rows = self.main_table.rowCount()
            for row in range(rows):
                item_4 = self.main_table.item(row, 4)
                if item_4.text() != '' and (temperature or temperature != ''):
                    value_4 = float(item_4.text())
                    temperature = float(temperature)
                    pressure = 0.9 * CONSTANTS.ACCELERATION_OF_GRAVITY * value_4 * ((353 / (273 + 5)) - (353 / (273 + temperature)))
                    pressure = "{:.2f}".format(round(pressure, 2))
                    self.main_table.item(row, 5).setText(pressure)
                    self.main_table.item(row, 5).setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                else:
                    self.main_table.setItem(row, 5, QTableWidgetItem())
                    self.main_table.item(row, 5).setText('')


    def calculate_klapan_pressure_loss(self, row=False, column=False) -> None:
        sender = self.sender()
        if isinstance(sender, QComboBox):
            klapan_widget_value = CONSTANTS.INIT_DATA.KLAPAN_ITEMS.get(row)
            hand_klapan_flow = self.klapan_input.text()
            current_klapan_flow = self.sputnik_table.item(0, 1).text()
            self._calculate_klapan_pressure_loss(klapan_widget_value, hand_klapan_flow, current_klapan_flow)

        elif isinstance(sender, QLineEdit):
            klapan_widget_value = CONSTANTS.INIT_DATA.KLAPAN_ITEMS.get(self.klapan_widget.currentText())
            hand_klapan_flow = row
            current_klapan_flow = self.sputnik_table.item(0, 1).text()
            self._calculate_klapan_pressure_loss(klapan_widget_value, hand_klapan_flow, current_klapan_flow)

        elif isinstance(sender, QTableWidget):
            if row == 0 and column == 1:
                klapan_widget_value = CONSTANTS.INIT_DATA.KLAPAN_ITEMS.get(self.klapan_widget.currentText())
                hand_klapan_flow = self.klapan_input.text()
                current_klapan_flow = self.sputnik_table.item(0, 1).text()
                self._calculate_klapan_pressure_loss(klapan_widget_value, hand_klapan_flow, current_klapan_flow)


    def _calculate_klapan_pressure_loss(self, klapan_widget_value, hand_klapan_flow, current_klapan_flow) -> None:
        if hand_klapan_flow or klapan_widget_value == '--':
            klapan_flow = hand_klapan_flow
        else:
            klapan_flow = klapan_widget_value

        if all([current_klapan_flow, klapan_flow]):
            current_klapan_flow = int(current_klapan_flow)
            klapan_flow = int(klapan_flow)
            pressure_loss = 10 * pow(current_klapan_flow / klapan_flow, 2)
            pressure_loss = "{:.2f}".format(round(pressure_loss, 2))
            self.sputnik_table.item(0, 13).setText(pressure_loss)
            self.sputnik_table.item(0, 13).setBackground(QColor(153, 255, 255))
        else:
            self.sputnik_table.item(0, 13).setText('')
            self.sputnik_table.item(0, 13).setBackground(QColor(255, 255, 255))


    def calculate_air_velocity(self, row, column) -> None:
        table = self.sender().objectName()
        match table:
            case CONSTANTS.SPUTNIK_TABLE.NAME:
                columns = (1, 3, 4)
                if column in columns:
                    air_flow = self.sputnik_table.item(row, 1).text()
                    a = self.sputnik_table.item(row, 3).text()
                    b = self.sputnik_table.item(row, 4).text()
                    if all([air_flow, a, b]):
                        velocity = float(air_flow) / (3_600 * ((float(a) * float(b) / 1_000_000)))
                        velocity = "{:.2f}".format(round(velocity, 2))
                        self.sputnik_table.item(row, 5).setText(velocity)
                    elif all([air_flow, a]):
                        velocity = float(air_flow) / (3_600 * (3.1415 * (float(a) / 1_000) * (float(a) / 1_000) / 4))
                        velocity = "{:.2f}".format(round(velocity, 2))
                        self.sputnik_table.item(row, 5).setText(velocity)
                    else:
                        self.sputnik_table.item(row, 5).setText('')
            case CONSTANTS.MAIN_TABLE.NAME:
                columns = (3, 10, 11)
                if column in columns:
                    air_flow = self.main_table.item(row, 3)
                    a = self.main_table.item(row, 10)
                    b = self.main_table.item(row, 11)
                    if all([air_flow is not None, a is not None, b is not None]) and all([air_flow.text() != '', a.text() != '', b.text() != '']):
                        velocity = float(air_flow.text()) / (3_600 * ((float(a.text()) * float(b.text()) / 1_000_000)))
                        velocity = "{:.2f}".format(round(velocity, 2))
                        self.main_table.item(row, 12).setText(velocity)
                        self.main_table.item(row, 12).setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    elif all([air_flow is not None, a is not None]) and all([air_flow.text() != '', a.text() != '']):
                        velocity = float(air_flow.text()) / (3_600 * (3.1415 * (float(a.text()) / 1_000) * (float(a.text()) / 1_000) / 4))
                        velocity = "{:.2f}".format(round(velocity, 2))
                        self.main_table.item(row, 12).setText(velocity)
                        self.main_table.item(row, 12).setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    else:
                        self.main_table.setItem(row, 12, QTableWidgetItem())
                        self.main_table.item(row, 12).setText('')


    def calculate_diameter(self, row, column) -> None:
        table = self.sender().objectName()
        match table:
            case CONSTANTS.SPUTNIK_TABLE.NAME:
                columns = (3, 4)
                if column in columns:
                    a = self.sputnik_table.item(row, 3).text()
                    b = self.sputnik_table.item(row, 4).text()
                    if all([a, b]):
                        diameter = (2 * float(a) * float(b) / (float(a) + float(b)) / 1_000)
                        diameter = "{:.3f}".format(round(diameter, 3))
                        self.sputnik_table.item(row, 6).setText(diameter)
                        self.sputnik_table.item(row, 6).setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    elif a:
                        diameter = float(a) / 1_000
                        diameter = "{:.3f}".format(round(diameter, 3))
                        self.sputnik_table.item(row, 6).setText(diameter)
                    else:
                        self.sputnik_table.item(row, 6).setText('')
            case CONSTANTS.MAIN_TABLE.NAME:
                columns = (10, 11)
                if column in columns:
                    a = self.main_table.item(row, 10)
                    b = self.main_table.item(row, 11)
                    if all([a is not None, b is not None]) and all([a.text() != '', b.text() != '']):
                        diameter = (2 * float(a.text()) * float(b.text()) / (float(a.text()) + float(b.text())) / 1_000)
                        diameter = "{:.3f}".format(round(diameter, 3))
                        self.main_table.item(row, 13).setText(diameter)
                        self.main_table.item(row, 13).setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    elif a is not None and a.text() != '':
                        diameter = float(a.text()) / 1_000
                        diameter = "{:.3f}".format(round(diameter, 3))
                        self.main_table.item(row, 13).setText(diameter)
                        self.main_table.item(row, 13).setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    else:
                        self.main_table.setItem(row, 13, QTableWidgetItem())
                        self.main_table.item(row, 13).setText('')


    def calculate_dynamic(self, row=False, column=False) -> None:
        sender = self.sender()
        table = sender.objectName()
        if isinstance(sender, QTableWidget):
            # если изменения данных произошли в таблице
            match table:
                case CONSTANTS.SPUTNIK_TABLE.NAME:
                    if column == 5:
                        temperature = self.temperature_widget.text()
                        velocity = self.sputnik_table.item(row, 5).text()
                        if all([temperature, velocity]):
                            temperature = float(temperature)
                            velocity = float(velocity)
                            density = 353 / (273.15 + temperature)
                            dynamic = velocity * velocity * density / 2
                            dynamic = "{:.2f}".format(round(dynamic, 2))
                            self.sputnik_table.item(row, 10).setText(dynamic)
                        else:
                            self.sputnik_table.item(row, 10).setText('')
                            self.sputnik_table.item(row, 10).setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                case CONSTANTS.MAIN_TABLE.NAME:
                    if column == 12:
                        temperature = self.temperature_widget.text()
                        velocity = self.main_table.item(row, 12).text()
                        if all([temperature, velocity]):
                            temperature = float(temperature)
                            velocity = float(velocity)
                            density = 353 / (273.15 + temperature)
                            dynamic = velocity * velocity * density / 2
                            dynamic = "{:.2f}".format(round(dynamic, 2))
                            self.main_table.item(row, 17).setText(dynamic)
                            self.main_table.item(row, 17).setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                        else:
                            self.main_table.setItem(row, 17, QTableWidgetItem())
                            self.main_table.item(row, 17).setText('')
                            self.main_table.item(row, 17).setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        else:
            # если изменилась температура воздуха
            temperature = row
            rows_sputnik = (1, 3)
            for row in rows_sputnik:
                velocity = self.sputnik_table.item(row, 5).text()
                if all([temperature, velocity]):
                    temperature = float(temperature)
                    velocity = float(velocity)
                    density = 353 / (273.15 + temperature)
                    dynamic = velocity * velocity * density / 2
                    dynamic = "{:.2f}".format(round(dynamic, 2))
                    self.sputnik_table.item(row, 10).setText(dynamic)
                    self.sputnik_table.item(row, 10).setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                else:
                    self.sputnik_table.item(row, 10).setText('')

            rows_main = self.main_table.rowCount()
            for row in range(rows_main):
                velocity = self.main_table.item(row, 12).text()
                if all([temperature, velocity]):
                    temperature = float(temperature)
                    velocity = float(velocity)
                    density = 353 / (273.15 + temperature)
                    dynamic = velocity * velocity * density / 2
                    dynamic = "{:.2f}".format(round(dynamic, 2))
                    self.main_table.item(row, 17).setText(dynamic)
                else:
                    self.main_table.item(row, 17).setText('')
                    self.main_table.item(row, 17).setTextAlignment(Qt.AlignmentFlag.AlignCenter)


    def calculate_specific_pressure_loss(self, row=False, column=False) -> None:
        sender = self.sender()
        name = sender.objectName()
        if isinstance(sender, QTableWidget):
            match name:
                case CONSTANTS.SPUTNIK_TABLE.NAME:
                    if column in (6, 10):
                        temperature = self.temperature_widget.text()
                        surface = self.surface_widget.text()
                        velocity = self.sputnik_table.item(row, 5).text()
                        diameter = self.sputnik_table.item(row, 6).text()
                        dynamic = self.sputnik_table.item(row, 10)
                        if all([temperature, diameter, velocity, surface, dynamic]):
                            try:
                                temperature = float(temperature)
                                velocity = float(velocity)
                                diameter = float(diameter)
                                surface = float(surface)
                                dynamic = float(dynamic.text())
                                mu = 1.458 * pow(10, -6) * pow((273.15 + temperature), 1.5) / ((273.15 + temperature) + 110.4)
                                density = 353 / (273.15 + temperature)
                                v = mu / density
                                re = velocity * diameter / v
                                lam = 0.11 * pow(((surface / 1_000) / diameter + 68 / re), 0.25)
                                resistivity = (lam / diameter) * dynamic
                                resistivity = "{:.4f}".format(round(resistivity, 4))
                                self.sputnik_table.item(row, 7).setText(resistivity)
                            except ZeroDivisionError:
                                self.sputnik_table.item(row, 7).setText('')
                        else:
                            self.sputnik_table.item(row, 7).setText('')
                case CONSTANTS.MAIN_TABLE.NAME:
                    if column in (13, 17):
                        temperature = self.temperature_widget.text()
                        surface = self.surface_widget.text()
                        velocity = self.main_table.item(row, 12)
                        diameter = self.main_table.item(row, 13)
                        dynamic = self.main_table.item(row, 17)
                        if all([temperature, diameter, velocity, surface, dynamic]) and all([temperature != '', diameter.text() != '', velocity.text() != '', surface != '', dynamic.text() != '']):
                            temperature = float(temperature)
                            surface = float(surface)
                            velocity = float(velocity.text())
                            diameter = float(diameter.text())
                            dynamic = float(dynamic.text())
                            mu = 1.458 * pow(10, -6) * pow((273.15 + temperature), 1.5) / ((273.15 + temperature) + 110.4)
                            density = 353 / (273.15 + temperature)
                            try:
                                v = mu / density
                                re = velocity * diameter / v
                                lam = 0.11 * pow(((surface / 1_000) / diameter + 68 / re), 0.25)
                                resistivity = (lam / diameter) * dynamic
                                resistivity = "{:.4f}".format(round(resistivity, 4))
                                self.main_table.item(row, 14).setText(resistivity)
                                self.main_table.item(row, 14).setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                            except ZeroDivisionError:
                                QMessageBox.warning(None, 'Внимание', 'This is a warning message')
                                self.main_table.setItem(row, 14, QTableWidgetItem())
                                self.main_table.item(row, 14).setText('')
                        else:
                            self.main_table.setItem(row, 14, QTableWidgetItem())
                            self.main_table.item(row, 14).setText('')
        elif isinstance(sender, QLineEdit) and name == 'surface':
            # если изменилась шероховатость
            surface = row
            rows_sputnik = (1, 3)
            for row in rows_sputnik:
                temperature = self.temperature_widget.text()
                velocity = self.sputnik_table.item(row, 5).text()
                diameter = self.sputnik_table.item(row, 6).text()
                dynamic = self.sputnik_table.item(row, 10).text()
                if all([temperature, diameter, velocity, surface, dynamic]):
                    temperature = float(temperature)
                    velocity = float(velocity)
                    diameter = float(diameter)
                    surface = float(surface)
                    dynamic = float(dynamic)
                    mu = 1.458 * pow(10, -6) * pow((273.15 + temperature), 1.5) / ((273.15 + temperature) + 110.4)
                    density = 353 / (273.15 + temperature)
                    v = mu / density
                    re = velocity * diameter / v
                    lam = 0.11 * pow(((surface / 1_000) / diameter + 68 / re), 0.25)
                    resistivity = (lam / diameter) * dynamic
                    resistivity = "{:.4f}".format(round(resistivity, 4))
                    self.sputnik_table.item(row, 7).setText(resistivity)
                else:
                    self.sputnik_table.item(row, 7).setText('')

            rows_main = self.main_table.rowCount()
            for row in range(rows_main):
                temperature = self.temperature_widget.text()
                velocity = self.main_table.item(row, 12).text()
                diameter = self.main_table.item(row, 13).text()
                dynamic = self.main_table.item(row, 17).text()
                if all([temperature, diameter, velocity, surface, dynamic]):
                    temperature = float(temperature)
                    velocity = float(velocity)
                    diameter = float(diameter)
                    surface = float(surface)
                    dynamic = float(dynamic)
                    mu = 1.458 * pow(10, -6) * pow((273.15 + temperature), 1.5) / ((273.15 + temperature) + 110.4)
                    density = 353 / (273.15 + temperature)
                    v = mu / density
                    re = velocity * diameter / v
                    lam = 0.11 * pow(((surface / 1_000) / diameter + 68 / re), 0.25)
                    resistivity = (lam / diameter) * dynamic
                    resistivity = "{:.4f}".format(round(resistivity, 4))
                    self.main_table.item(row, 14).setText(resistivity)
                    self.main_table.item(row, 14).setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                else:
                    self.main_table.item(row, 14).setText('')


    def uncheck_radio_button_1(self) -> None:
        self.radio_button2.setChecked(True)
        self.radio_button1.setChecked(False)


    def uncheck_radio_button_2(self) -> None:
        self.radio_button1.setChecked(True)
        self.radio_button2.setChecked(False)


    def set_sputnik_airflow_in_main_table_by_radiobutton_1(self, checked) -> None:
        if checked:
            self.radio_button2.setChecked(False)
            one_side_flow = self.sputnik_table.item(1, 1).text()

            if one_side_flow:
                self.main_table.item(0, 3).setText(str(int(one_side_flow)))
                self.main_table.item(0, 3).setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.main_table.item(0, 3).setBackground(QColor(255, 255, 255))
                self.fill_air_flow_column_in_main_table(one_side_flow)
            else:
                self.main_table.item(0, 3).setText('')
                self.clean_air_flow_column_in_main_table()


    def set_sputnik_airflow_in_main_table_by_radiobutton_2(self, checked) -> None:
        if checked:
            self.radio_button1.setChecked(False)
            one_side_flow = self.sputnik_table.item(1, 1).text()
            two_side_flow = self.sputnik_table.item(3, 1).text()

            if all([one_side_flow, two_side_flow]):
                flow = int(one_side_flow) + int(two_side_flow)
                self.main_table.item(0, 3).setText(str(flow))
                self.main_table.item(0, 3).setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.main_table.item(0, 3).setBackground(QColor(255, 255, 255))
                self.fill_air_flow_column_in_main_table(flow)
            else:
                self.main_table.item(0, 3).setText('')
                self.clean_air_flow_column_in_main_table()


    def set_sputnik_airflow_in_main_table(self, row, column) -> None:
        flows_cells = ((1, 1), (3, 1))
        one_side_flow = self.sputnik_table.item(1, 1).text()
        two_side_flow = self.sputnik_table.item(3, 1).text()

        if (row, column) in flows_cells:
            if all([self.radio_button1.isChecked(), one_side_flow]):
                self.main_table.item(0, 3).setText(str(int(one_side_flow)))
                self.main_table.item(0, 3).setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.main_table.item(0, 3).setBackground(QColor(255, 255, 255))
                self.fill_air_flow_column_in_main_table(one_side_flow)
            elif all([self.radio_button2.isChecked(), one_side_flow, two_side_flow]):
                flow = int(one_side_flow) + int(two_side_flow)
                self.main_table.item(0, 3).setText(str(flow))
                self.main_table.item(0, 3).setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.main_table.item(0, 3).setBackground(QColor(255, 255, 255))
                self.fill_air_flow_column_in_main_table(flow)
            else:
                self.main_table.item(0, 3).setText('')
                self.clean_air_flow_column_in_main_table()


    def fill_air_flow_column_in_main_table(self, flow) -> None:
        num_rows = self.main_table.rowCount()
        for row in range(1, num_rows):
            value = int(flow) + int(self.main_table.item(row-1, 3).text())
            value = str(value)
            self.main_table.item(row, 3).setText(value)
            self.main_table.item(row, 3).setTextAlignment(Qt.AlignmentFlag.AlignCenter)


    def update_air_flow_column_in_main_table_after_delete_row(self) -> None:
        init_flow = self.main_table.item(0, 3)
        if all([init_flow, init_flow.text() != '']):
            init_flow = int(init_flow.text())
            num_rows = self.main_table.rowCount()
            for row in range(1, num_rows):
                value = int(init_flow) + int(self.main_table.item(row-1, 3).text())
                value = str(value)
                self.main_table.item(row, 3).setText(value)
                self.main_table.item(row, 3).setTextAlignment(Qt.AlignmentFlag.AlignCenter)


    def clean_air_flow_column_in_main_table(self) -> None:
        num_rows = self.main_table.rowCount()
        for row in range(1, num_rows):
            self.main_table.item(row, 3).setText('')


    def calculate_m(self, row, column) -> None:
        table = self.sender().objectName()
        axis_x = CONSTANTS.REFERENCE_DATA.M.X
        axis_y = CONSTANTS.REFERENCE_DATA.M.Y
        z = array(CONSTANTS.REFERENCE_DATA.M.TABLE)
        interpolator = RegularGridInterpolator((axis_y, axis_x), z)
        match table:
            case CONSTANTS.SPUTNIK_TABLE.NAME:
                if column in (3, 4):
                    a = self.sputnik_table.item(row, 3).text()
                    b = self.sputnik_table.item(row, 4).text()
                    if all([a, b]) and all([100 <= int(a) <= 1_500, 100 <= int(b) <= 1_500]):
                        a, b = int(a), int(b)
                        m = interpolator((b, a))
                        m = "{:.3f}".format(around(m, 3))
                        self.sputnik_table.item(row, 8).setText(m)
                    elif a and 100 <= int(a) <= 1_500:
                        a = int(a)
                        m = interpolator((a, a))
                        m = "{:.3f}".format(around(m, 3))
                        self.sputnik_table.item(row, 8).setText(m)
                    else:
                        self.sputnik_table.item(row, 8).setText('')
            case CONSTANTS.MAIN_TABLE.NAME:
                if column in (10, 11):
                    a = self.main_table.item(row, 10)
                    b = self.main_table.item(row, 11)
                    if all([a, b]) and all([a.text() != '', b.text() != '']) and all([100 <= int(a.text()) <= 1_500, 100 <= int(b.text()) <= 1_500]):
                        a, b = int(a.text()), int(b.text())
                        m = interpolator((b, a))
                        m = "{:.3f}".format(around(m, 3))
                        self.main_table.setItem(row, 15, QTableWidgetItem())
                        self.main_table.item(row, 15).setText(m)
                        self.main_table.item(row, 15).setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    elif a and a.text() != '' and 100 <= int(a.text()) <= 1_500:
                        a = int(a.text())
                        m = interpolator((a, a))
                        m = "{:.3f}".format(around(m, 3))
                        self.main_table.setItem(row, 15, QTableWidgetItem())
                        self.main_table.item(row, 15).setText(m)
                        self.main_table.item(row, 15).setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    else:
                        self.main_table.setItem(row, 15, QTableWidgetItem())
                        self.main_table.item(row, 15).setText('')


    def calculate_linear_pressure_loss(self, row, column) -> None:
        table = self.sender().objectName()
        match table:
            case CONSTANTS.SPUTNIK_TABLE.NAME:
                if column in (2, 7, 8):
                    l = self.sputnik_table.item(row, 2).text()
                    r = self.sputnik_table.item(row, 7).text()
                    m = self.sputnik_table.item(row, 8).text()
                    if all([l, r, m]):
                        l = float(l)
                        r = float(r)
                        m = float(m)
                        result = "{:.4f}".format(round(l * r * m, 4))
                        self.sputnik_table.item(row, 9).setText(result)
                    else:
                        self.sputnik_table.item(row, 9).setText('')
            case CONSTANTS.MAIN_TABLE.NAME:
                if column in (1, 14, 15):
                    l = self.main_table.item(row, 1)
                    r = self.main_table.item(row, 14)
                    m = self.main_table.item(row, 15)
                    if all([l, r, m]) and all([l.text() != '', r.text() != '', m.text() != '']):
                        l = float(l.text())
                        r = float(r.text())
                        m = float(m.text())
                        result = "{:.4f}".format(round(l * r * m, 4))
                        self.main_table.setItem(row, 16, QTableWidgetItem())
                        self.main_table.item(row, 16).setText(result)
                        self.main_table.item(row, 16).setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    else:
                        self.main_table.setItem(row, 16, QTableWidgetItem())
                        self.main_table.item(row, 16).setText('')


    def calculate_local_pressure_loss(self, row, column) -> None:
        if column in (10, 11):
            dynamic = self.sputnik_table.item(row, 10).text()
            kms = self.sputnik_table.item(row, 11).text()
            if all([dynamic, kms]):
                dynamic = float(dynamic)
                kms = float(kms)
                result = "{:.2f}".format(round(dynamic * kms, 2))
                self.sputnik_table.item(row, 12).setText(result)
            else:
                self.sputnik_table.item(row, 12).setText('')


    def calculate_sputnik_full_pressure_loss(self, row, column) -> None:
        if column in (9, 12):
            linear_pressure_loss = self.sputnik_table.item(row, 9).text()
            local_pressure_loss = self.sputnik_table.item(row, 12).text()
            if all([linear_pressure_loss, local_pressure_loss]):
                linear_pressure_loss = float(linear_pressure_loss)
                local_pressure_loss = float(local_pressure_loss)
                result = "{:.2f}".format(round(linear_pressure_loss + local_pressure_loss, 2))
                self.sputnik_table.item(row, 13).setText(result)
            else:
                self.sputnik_table.item(row, 13).setText('')


    def calculate_kms_by_radiobutton_1(self, checked) -> None:
        if checked:
            self.radio_button2.setChecked(False)
            one_side_flow = self.sputnik_table.item(1, 1).text()
            sputnik_a = self.sputnik_table.item(1, 3).text()
            sputnik_b = self.sputnik_table.item(1, 4).text()
            if sputnik_b == '':
                sputnik_b = sputnik_a
            self._calculate_kms_by_radiobutton(one_side_flow, sputnik_a, sputnik_b)


    def calculate_kms_by_radiobutton_2(self, checked) -> None:
        if checked:
            self.radio_button1.setChecked(False)
            one_side_flow = self.sputnik_table.item(3, 1).text()
            sputnik_a = self.sputnik_table.item(3, 3).text()
            sputnik_b = self.sputnik_table.item(3, 4).text()
            if sputnik_b == '':
                sputnik_b = sputnik_a
            self._calculate_kms_by_radiobutton(one_side_flow, sputnik_a, sputnik_b)


    def calculate_kms(self, row, column) -> None:
        table = self.sender().objectName()
        if (table == CONSTANTS.SPUTNIK_TABLE.NAME and column in (1, 3, 4)) or (table == CONSTANTS.MAIN_TABLE.NAME and column in (3, 10, 11)):
            one_side_flow = self.sputnik_table.item(1, 1).text()
            two_side_flow = self.sputnik_table.item(3, 1).text()

            if all([self.radio_button1.isChecked(), one_side_flow]):
                sputnik_flow = one_side_flow
                sputnik_a = self.sputnik_table.item(1, 3).text()
                sputnik_b = self.sputnik_table.item(1, 4).text()
                if sputnik_b == '':
                    sputnik_b = sputnik_a
                self._calculate_kms_by_radiobutton(sputnik_flow, sputnik_a, sputnik_b)
            elif all([self.radio_button2.isChecked(), two_side_flow]):
                sputnik_flow = two_side_flow
                sputnik_a = self.sputnik_table.item(3, 3).text()
                sputnik_b = self.sputnik_table.item(3, 4).text()
                if sputnik_b == '':
                    sputnik_b = sputnik_a
                self._calculate_kms_by_radiobutton(sputnik_flow, sputnik_a, sputnik_b)


    def _calculate_kms_by_radiobutton(self, flow, a, b, num_rows=False) -> None:
        sputnik_a, sputnik_b = a, b
        if all([flow, sputnik_a, sputnik_b]):
            sputnik_flow = float(flow)
            sputnik_a, sputnik_b = int(sputnik_a), int(sputnik_b)
            num_rows = self.main_table.rowCount()
            for row in range(0, num_rows-1):
                floor_flow = self.main_table.item(row+1, 3)
                main_a = self.main_table.item(row, 10)
                main_b = self.main_table.item(row, 11)
                if not main_b or main_b.text() == '':
                    main_b = main_a

                if all([floor_flow, main_a, main_b]) and all([floor_flow.text() != '', main_a.text() != '', main_b.text() != '']):
                    floor_flow = float(floor_flow.text())
                    main_a, main_b = int(main_a.text()), int(main_b.text())

                    x_l = sputnik_flow / floor_flow
                    y_f = ((sputnik_a * sputnik_b) / 1_000_000) / ((main_a * main_b) / 1_000_000)

                    if x_l > 1:
                        x_l = 1
                    elif x_l < 0:
                        x_l = 0

                    if y_f > 1:
                        y_f = 1
                    elif y_f < 0.1:
                        y_f = 0.1

                    pass_kms = self._interpol_kms(x_l, y_f, type='pass_kms')
                    self.main_table.setItem(row, 8, QTableWidgetItem())
                    self.main_table.item(row, 8).setText(pass_kms)
                    self.main_table.item(row, 8).setTextAlignment(Qt.AlignmentFlag.AlignCenter)

                    branch_kms = self._interpol_kms(x_l, y_f, type='branch_kms')
                    self.main_table.setItem(row, 9, QTableWidgetItem())
                    self.main_table.item(row, 9).setText(branch_kms)
                    self.main_table.item(row, 9).setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                else:
                    self.main_table.setItem(row, 8, QTableWidgetItem())
                    self.main_table.setItem(row, 9, QTableWidgetItem())
                    self.main_table.item(row, 8).setText('')
                    self.main_table.item(row, 9).setText('')
        else:
            num_rows = self.main_table.rowCount()
            for row in range(0, num_rows-1):
                self.main_table.setItem(row, 8, QTableWidgetItem())
                self.main_table.setItem(row, 9, QTableWidgetItem())
                self.main_table.item(row, 8).setText('')
                self.main_table.item(row, 9).setText('')


    def _interpol_kms(self, x_l, y_f, type) -> str:
        match type:
            case 'pass_kms':
                pass_axis_x = CONSTANTS.REFERENCE_DATA.KMS_PASS.X_L
                pass_axis_y = CONSTANTS.REFERENCE_DATA.KMS_PASS.Y_F
                pass_z = array(CONSTANTS.REFERENCE_DATA.KMS_PASS.TABLE)
                pass_interpolator = RegularGridInterpolator((pass_axis_y, pass_axis_x), pass_z)
                pass_kms = pass_interpolator((y_f, x_l))
                pass_kms = "{:.2f}".format(around(pass_kms, 2))
                return pass_kms
            case 'branch_kms':
                branch_axis_x = CONSTANTS.REFERENCE_DATA.KMS_BRANCH.X_L
                branch_axis_y = CONSTANTS.REFERENCE_DATA.KMS_BRANCH.Y_F
                branch_z = array(CONSTANTS.REFERENCE_DATA.KMS_BRANCH.TABLE)
                branch_interpolator = RegularGridInterpolator((branch_axis_y, branch_axis_x), branch_z)
                branch_kms = branch_interpolator((y_f, x_l))
                branch_kms = "{:.2f}".format(around(branch_kms, 2))
                return branch_kms


    def calculate_pass_pressure(self, row, column) -> None:
        if column in (8, 17):
            kms = self.main_table.item(row, 8)
            dynamic = self.main_table.item(row, 17)
            if all([kms, dynamic]) and all([kms.text() != '', dynamic.text() != '']):
                kms = float(kms.text())
                dynamic = float(dynamic.text())
                result = kms * dynamic
                result = "{:.2f}".format(round(result, 2))
                self.main_table.setItem(row, 18, QTableWidgetItem())
                self.main_table.item(row, 18).setText(result)
                self.main_table.item(row, 18).setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            else:
                self.main_table.setItem(row, 18, QTableWidgetItem())
                self.main_table.item(row, 18).setText('')


    def calculate_branch_pressure_by_radiobutton_1(self, checked) -> None:
        if checked:
            self.radio_button2.setChecked(False)
            dynamic = self.sputnik_table.item(1, 10).text()
            self._calculate_branch_pressure(dynamic)


    def calculate_branch_pressure_by_radiobutton_2(self, checked) -> None:
        if checked:
            self.radio_button1.setChecked(False)
            dynamic = self.sputnik_table.item(3, 10).text()
            self._calculate_branch_pressure(dynamic)


    def calculate_branch_pressure(self, row, column) -> None:
        table = self.sender().objectName()
        if all([table == CONSTANTS.SPUTNIK_TABLE.NAME and column == 10]) or all([table == CONSTANTS.MAIN_TABLE.NAME and column == 9]):
            dynamic_one = self.sputnik_table.item(1, 10).text()
            dynamic_two = self.sputnik_table.item(3, 10).text()
            if all([self.radio_button1.isChecked(), dynamic_one]):
                self._calculate_branch_pressure(dynamic_one)
            elif all([self.radio_button1.isChecked(), dynamic_two]):
                self._calculate_branch_pressure(dynamic_two)


    def _calculate_branch_pressure(self, dynamic) -> None:
        num_rows = self.main_table.rowCount()
        if dynamic:
            dynamic = float(dynamic)
            for row in range(num_rows):
                kms = self.main_table.item(row, 9)
                if kms and kms.text() != '':
                    kms = float(kms.text())
                    result = kms * dynamic
                    result = "{:.2f}".format(round(result, 2))
                    self.main_table.setItem(row, 19, QTableWidgetItem())
                    self.main_table.item(row, 19).setText(result)
                    self.main_table.item(row, 19).setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                else:
                    self.main_table.setItem(row, 19, QTableWidgetItem())
                    self.main_table.item(row, 19).setText('')
        else:
            for row in range(num_rows):
                self.main_table.setItem(row, 19, QTableWidgetItem())
                self.main_table.item(row, 19).setText('')


    def calculate_result_sputnik_pressure(self, row, column) -> None:
        pressure_cells = ((0, 13), (1, 13), (3, 13))
        if (row, column) in pressure_cells:
            klapan_pressure = self.sputnik_table.item(0, 13).text()
            one_side_pressure = self.sputnik_table.item(1, 13).text()
            two_side_pressure = self.sputnik_table.item(3, 13).text()

            if all([klapan_pressure, one_side_pressure]):
                klapan_pressure = float(klapan_pressure)
                one_side_pressure = float(one_side_pressure)
                first_result = "{:.2f}".format(round((klapan_pressure + one_side_pressure), 2))
                self.sputnik_table.item(2, 13).setText(first_result)
                self.sputnik_table.item(2, 13).setBackground(QColor(153, 255, 255))
            else:
                self.sputnik_table.item(2, 13).setText('')
                self.sputnik_table.item(2, 13).setBackground(QColor(255, 255, 255))

            if all([klapan_pressure, two_side_pressure]):
                klapan_pressure = float(klapan_pressure)
                two_side_pressure = float(two_side_pressure)
                second_result = "{:.2f}".format(round((klapan_pressure + two_side_pressure), 2))
                self.sputnik_table.item(4, 13).setText(second_result)
                self.sputnik_table.item(4, 13).setBackground(QColor(153, 255, 255))
            else:
                self.sputnik_table.item(4, 13).setText('')
                self.sputnik_table.item(4, 13).setBackground(QColor(255, 255, 255))


    def set_floor_number_in_main_table(self) -> None:
        if not self.last_floor:
            num_rows = self.main_table.rowCount()
            for row in range(num_rows):
                self.main_table.item(row, 0).setText(str(row+1))
                self.main_table.item(row, 0).setTextAlignment(Qt.AlignmentFlag.AlignCenter)


    # def set_channel_height_in_main_table(self) -> None:
    #     height = self.channel_height_widget.text()
    #     if height != '':
    #         height = str(float(height))
    #         self.main_table.setItem(0, 4, QTableWidgetItem())
    #         self.main_table.item(0, 4).setText(height)
    #         self.main_table.item(0, 4).setTextAlignment(Qt.AlignmentFlag.AlignCenter)
    #         # self.fill_height_column_in_main_table()
    #     else:
    #         self.main_table.setItem(0, 4, QTableWidgetItem())
    #         self.main_table.item(0, 4).setText('')


    def set_base_floor_height_in_main_table(self) -> None:
        num_rows = self.main_table.rowCount()
        base_floor_height = self.floor_height_widget.text()
        if base_floor_height:
            all_values = []
            # собираем все значения высоты этажей
            for row in range(num_rows):
                all_values.append(self.main_table.item(row, 1).text())

            # если значения по этажам не менялись - меняем на новое
            if len(set(all_values)) == 1:
                for row in range(num_rows):
                    self.main_table.item(row, 1).setText("{:.2f}".format(float(base_floor_height)))
                    self.main_table.item(row, 1).setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            else:
                # определяем максимально встречающееся значение, т.е. ранее установленное значение которое не было изменено
                current_height = max(set(all_values), key=all_values.count)
                # заменяем все базовые значения на новое, кастомные значения сохраняются
                for row in range(num_rows):
                    if self.main_table.item(row, 1).text() == current_height:
                        self.main_table.item(row, 1).setText("{:.2f}".format(float(self.floor_height_widget.text())))
                        self.main_table.item(row, 1).setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        else:
            for row in range(num_rows):
                self.main_table.item(row, 1).setText('')


    def calculate_height(self, row=False, column=False):
        sender = self.sender().objectName()
        match sender:
            case 'main':
                if column == 1:
                    channel_height = self.channel_height_widget.text()
                    base_floor_height = self.floor_height_widget.text()
                    if channel_height and base_floor_height:
                        self.main_table.item(0, 4).setText("{:.2f}".format(float(channel_height)))
                        self.main_table.item(0, 4).setTextAlignment(Qt.AlignmentFlag.AlignCenter)

                        floor_height = self.main_table.item(row, 1)
                        previous_height = self.main_table.item(row-1, 4)
                        if all([floor_height, previous_height]) and all([floor_height.text() != '', previous_height.text() != '']):
                            height = float(previous_height.text()) - float(floor_height.text())
                            if (height > 0 and previous_height.text() != '') or (height == 0 and previous_height.text() != ''):
                                height = "{:.2f}".format(round(height, 2))
                                self.main_table.item(row, 4).setText(height)
                                self.main_table.item(row, 4).setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                            else:
                                self.main_table.item(row, 4).setText('')
                                self._clean_height_column_in_main_table(row)
                    else:
                        self._clean_height_column_in_main_table(0)
            case 'channel_height':
                channel_height = row
                if channel_height:
                    self.main_table.item(0, 4).setText("{:.2f}".format(float(channel_height)))
                    self.main_table.item(0, 4).setTextAlignment(Qt.AlignmentFlag.AlignCenter)

                    num_rows = self.main_table.rowCount()
                    self._fill_height_column_in_main_table(num_rows)
                else:
                    self._clean_height_column_in_main_table(0)


    def _clean_height_column_in_main_table(self, row) -> None:
        """Clean height column in main table from `row`."""
        num_rows = self.main_table.rowCount()
        for row in range(row, num_rows):
            if self.main_table.item(row, 4):
                self.main_table.item(row, 4).setText('')


    def update_height_column_in_main_table(self, row, column) -> None:
        base_floor_height = self.floor_height_widget.text()
        current_floor_height = self.main_table.item(row, 1)
        if current_floor_height and current_floor_height.text() != '':
            current_floor_height = current_floor_height.text()
            if column == 1 and current_floor_height != base_floor_height:
                num_rows = self.main_table.rowCount()
                self._fill_height_column_in_main_table(num_rows, row)


    def _fill_height_column_in_main_table(self, num_rows, row=False) -> None:
        start = row if row else 1
        for line in range(start, num_rows):
            floor_height = self.main_table.item(line, 1)
            previous_height = self.main_table.item(line-1, 4)
            if all([floor_height, previous_height]) and all([floor_height.text() != '', previous_height.text() != '']):
                height = float(previous_height.text()) - float(floor_height.text())
                if (height > 0 and previous_height.text() != '') or (height == 0 and previous_height.text() != ''):
                    height = "{:.2f}".format(round(height, 2))
                    self.main_table.item(line, 4).setText(height)
                    self.main_table.item(line, 4).setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                else:
                    self.main_table.item(line, 4).setText('')
                    self._clean_height_column_in_main_table(line)
            else:
                self._clean_height_column_in_main_table(line)


    def calculate_full_pressure(self, row=False, column=False) -> None:
        sender = self.sender()
        table = sender.objectName()
        if (table == CONSTANTS.SPUTNIK_TABLE.NAME and column == 13) or (table == CONSTANTS.MAIN_TABLE.NAME and column in (16, 18, 19)) or isinstance(sender, QComboBox):
            klapan_full_pressure_one = self.sputnik_table.item(2, 13).text()
            klapan_full_pressure_two = self.sputnik_table.item(4, 13).text()
            
            if all([self.radio_button1.isChecked(), klapan_full_pressure_one]):
                klapan_full_pressure = klapan_full_pressure_one
                self._calculate_full_pressure_by_radiobutton(klapan_full_pressure)
            elif all([self.radio_button2.isChecked(), klapan_full_pressure_two]):
                klapan_full_pressure = klapan_full_pressure_two
                self._calculate_full_pressure_by_radiobutton(klapan_full_pressure)


    def calculate_full_pressure_by_radiobutton_1(self, checked) -> None:
        if checked:
            self.radio_button2.setChecked(False)
            klapan_full_pressure = self.sputnik_table.item(2, 13).text()
            self._calculate_full_pressure_by_radiobutton(klapan_full_pressure)


    def calculate_full_pressure_by_radiobutton_2(self, checked) -> None:
        if checked:
            self.radio_button1.setChecked(False)
            klapan_full_pressure_1 = self.sputnik_table.item(2, 13).text()
            klapan_full_pressure_2 = self.sputnik_table.item(4, 13).text()
            if all([klapan_full_pressure_1, klapan_full_pressure_2]):
                klapan_full_pressure = str(max(float(klapan_full_pressure_1), float(klapan_full_pressure_2)))
                self._calculate_full_pressure_by_radiobutton(klapan_full_pressure)


    def _calculate_full_pressure_by_radiobutton(self, klapan_full_pressure) -> None:
        num_rows = self.main_table.rowCount()
        for row in range(num_rows):
            branch_pressure = self.main_table.item(row, 19)

            if klapan_full_pressure and branch_pressure and branch_pressure.text() != '':
                branch_pressure = float(branch_pressure.text())
                klapan_full_pressure = float(klapan_full_pressure)

                all_pass_pressure = []
                for line in range(row, num_rows):
                    pass_pressure = self.main_table.item(line, 16)
                    if pass_pressure and pass_pressure.text() != '':
                        all_pass_pressure.append(pass_pressure.text())
                sum_pass_pressure = sum(map(float, all_pass_pressure))

                all_linear_pressure = []
                for line in range(row, num_rows):
                    linear_pressure = self.main_table.item(line, 18)
                    if linear_pressure and linear_pressure.text() != '':
                        all_linear_pressure.append(linear_pressure.text())
                sum_linear_pressure = sum(map(float, all_linear_pressure))

                result = klapan_full_pressure + branch_pressure + sum_pass_pressure + sum_linear_pressure
                result = "{:.2f}".format(round(result, 2))
                self.main_table.setItem(row, 20, QTableWidgetItem())
                self.main_table.item(row, 20).setText(result)
                self.main_table.item(row, 20).setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            else:
                self.main_table.setItem(row, 20, QTableWidgetItem())
                self.main_table.item(row, 20).setText('')


    def update_last_row_after_delete_row(self) -> None:
        num_rows = self.main_table.rowCount()
        kms_cols = (8, 9)
        clean_cols = (18, 19, 20, 21)
        
        kms_pass = self.main_table.item(num_rows, 8)
        kms_branch = self.main_table.item(num_rows, 9)

        for col in kms_cols:
            if self.main_table.item(num_rows, col) and self.sputnik_table.item(num_rows, col).text != '':
                self.main_table.item(num_rows, col).setText('')
                self.main_table.item(num_rows, col).setBackground(QColor(229, 255, 204))
        
        if all([kms_pass, kms_branch]) or all([kms_pass.text() != '', kms_branch.text() != '']):
            for col in kms_cols:
                self.main_table.item(num_rows, col).setText('')
                self.main_table.item(num_rows, col).setBackground(QColor(229, 255, 204))
            for col in clean_cols:
                self.main_table.item(num_rows, col).setText('')
                if col == 21:
                    self.main_table.item(num_rows, col).setBackground(QColor(255, 255, 255))


    def create_deflector_calculation(self) -> object:
        _box = QGroupBox(CONSTANTS.DEFLECTOR.TITLE)
        style = self.box_style
        _box.setStyleSheet(style)
        _box.setFixedHeight(316)
        _box.setFixedWidth(540)

        _layout = QHBoxLayout()

        self.deflector = QTableWidget(9, 1)
        self.deflector.horizontalHeader().setVisible(False)
        self.deflector.verticalHeader().setStyleSheet("QHeaderView::section { background-color: #F3F3F3 }")
        self.deflector.verticalHeader().setStyleSheet("QHeaderView::section { border-top: 0px solid gray; border-bottom: 0px solid gray; }")
        self.deflector.setStyleSheet("QHeaderView {padding: 5px; vertical-align: top} QTableWidget { border: 2px solid grey; }")
        self.deflector.setVerticalHeaderLabels(CONSTANTS.DEFLECTOR.HEADER)
        self.deflector.setObjectName(CONSTANTS.DEFLECTOR.NAME)

        num_rows = self.deflector.rowCount()
        for row in range(num_rows):
            self.deflector.setItem(row, 0, QTableWidgetItem())
            self.deflector.item(row, 0).setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            if row == 0:
                self.deflector.item(row, 0).setBackground(QColor(229, 255, 204))
            else:
                self.deflector.item(row, 0).setFlags(Qt.ItemFlag.ItemIsEnabled)

        self.deflector.cellChanged.connect(self.calculate_recommended_deflector_velocity)
        self.deflector.cellChanged.connect(self.calculate_required_deflector_square)
        self.deflector.cellChanged.connect(self.calculate_deflector_diameter)
        self.deflector.cellChanged.connect(self.calculate_real_deflector_velocity)
        self.deflector.cellChanged.connect(self.calculate_velocity_relation)
        self.deflector.cellChanged.connect(self.calculate_pressure_relation)
        self.deflector.cellChanged.connect(self.calculate_deflector_pressure)
        self.deflector.cellChanged.connect(self.activate_deflector_checkbox)
        self.deflector.cellChanged.connect(self.calculate_available_pressure)

        self.deflector.cellChanged.connect(self.set_deflector_pressure_in_main_table)

        self.deflector.itemChanged.connect(self.update_deflector_pressure_in_main_table)
        self.deflector.itemChanged.connect(self.validate_deflector_input)


        _layout.addWidget(self.deflector)
        _box.setLayout(_layout)
        # _box.adjustSize()
        return _box


    def validate_deflector_input(self, item) -> None:
        if item.column() == 0 and item.row() == 0:
            pattern = r'^(?:[0-9]|[0-9]\d|50)(?:\.\d{1,2})?$'
            if not re.match(pattern, item.text()):
                item.setText('')
                item.setBackground(QColor(255, 153, 153))
            else:
                item.setBackground(QColor(229, 255, 204))
                item.setText(item.text())


    def calculate_recommended_deflector_velocity(self, row, column) -> None:
        if row == 0 and column == 0:
            window_velocity = self.deflector.item(0, 0).text()
            if window_velocity:
                window_velocity = float(window_velocity)
                result = "{:.2f}".format(round(window_velocity * 0.3, 2))
                self.deflector.item(1, 0).setText(result)
            else:
                self.deflector.item(1, 0).setText('')


    def set_full_air_flow_in_deflector(self, row, column) -> None:
        last_row_main_table = self.main_table.rowCount() - 1
        if row == last_row_main_table and column == 3:
            air_flow = self.main_table.item(last_row_main_table, 3)
            if air_flow and air_flow.text() != '':
                self.deflector.item(2, 0).setText(air_flow.text())
            else:
                self.deflector.item(2, 0).setText('')


    def update_full_air_flow_in_deflector_after_delete_row(self) -> None:
        num_rows = self.main_table.rowCount()
        last_row_main_table = num_rows - 1
        air_flow = self.main_table.item(last_row_main_table, 3)
        if air_flow and air_flow.text() != '':
            self.deflector.item(2, 0).setText(air_flow.text())
        else:
            self.deflector.item(2, 0).setText('')


    def calculate_required_deflector_square(self, row, column) -> None:
        if column == 0 and row in (1, 2):
            recommended_velocity = self.deflector.item(1, 0).text()
            air_flow = self.deflector.item(2, 0).text()
            if all([recommended_velocity, air_flow]):
                recommended_velocity = float(recommended_velocity)
                air_flow = float(air_flow)
                result = "{:.3f}".format(round(air_flow / (3_600 * recommended_velocity), 3))
                self.deflector.item(3, 0).setText(result)
            else:
                self.deflector.item(3, 0).setText('')


    def calculate_deflector_diameter(self, row, column) -> None:
        if column == 0 and row == 3:
            required_deflector_square = self.deflector.item(3, 0).text()
            if required_deflector_square:
                required_deflector_square = float(required_deflector_square)

                for k in CONSTANTS.DEFLECTOR.DIAMETERS.keys():
                    if required_deflector_square <= k:
                        diameter = CONSTANTS.DEFLECTOR.DIAMETERS.get(k)
                        self.deflector.item(4, 0).setText(diameter)
                        self.deflector.item(4, 0).setBackground(QColor(255, 255, 255))
                        break
                    else:
                        self.deflector.item(4, 0).setBackground(QColor(229, 255, 204))
                        self.deflector.item(4, 0).setText('Нет значения!')
            else:
                self.deflector.item(4, 0).setText('')


    def calculate_real_deflector_velocity(self, row, column) -> None:
        if (row, column) in ((2, 0), (4, 0)):
            air_flow = self.deflector.item(2, 0).text()
            diameter = self.deflector.item(4, 0).text()
            if all([air_flow, diameter]) and diameter != 'Нет значения!':
                air_flow = float(air_flow)
                diameter = float(diameter)
                result = "{:.2f}".format(round(air_flow / (3_600 * math.pi * (pow(diameter / 1_000, 2) / 4)), 2))
                self.deflector.item(5, 0).setText(result)
            else:
                self.deflector.item(5, 0).setText('')


    def calculate_velocity_relation(self, row, column) -> None:
        if (row, column) in ((0, 0), (5, 0)):
            wind_velocity = self.deflector.item(0, 0).text()
            real_velocity = self.deflector.item(5, 0).text()
            if all([wind_velocity, real_velocity]):
                wind_velocity = float(wind_velocity)
                real_velocity = float(real_velocity)
                result = "{:.2f}".format(round(real_velocity / wind_velocity, 2))
                self.deflector.item(6, 0).setText(result)
            else:
                self.deflector.item(6, 0).setText('')


    def calculate_pressure_relation(self, row, column) -> None:
        if row == 6 and column == 0:
            velocity_relation = self.deflector.item(6, 0).text()
            if velocity_relation:
                try:
                    x = float(velocity_relation)
                    f = interp1d(
                        CONSTANTS.REFERENCE_DATA.DEFLECTOR_PRESSURE_RELATION.X,
                        CONSTANTS.REFERENCE_DATA.DEFLECTOR_PRESSURE_RELATION.TABLE
                    )
                    result = "{:.2f}".format(around(f(x), 2))
                    self.deflector.item(7, 0).setText(result)
                except ValueError:
                    self.deflector.item(7, 0).setText('')
            else:
                self.deflector.item(7, 0).setText('')


    def calculate_deflector_pressure(self, row, column) -> None:
        if (row, column) in ((0, 0), (7, 0)):
            wind_velocity = self.deflector.item(0, 0).text()
            pressure_relation = self.deflector.item(7, 0).text()
            if all([wind_velocity, pressure_relation]):
                wind_velocity = float(wind_velocity)
                pressure_relation = float(pressure_relation)
                result = pressure_relation * ((353 / (273.15 + 5)) * pow(wind_velocity, 2) / 2)
                result = "{:.2f}".format(round(result, 2))
                self.deflector.item(8, 0).setText(result)
                self.deflector.item(8, 0).setBackground(QColor(153, 255, 255))
            else:
                self.deflector.item(8, 0).setText('')
                self.deflector.item(8, 0).setBackground(QColor(255, 255, 255))


    def _join_deflector_column_cells_in_main_table(self) -> None:
        num_rows = self.main_table.rowCount()
        self.main_table.setSpan(0, 6, num_rows, 1)


    def activate_deflector_checkbox(self, row, column) -> None:
        if row == 8 and column == 0:
            deflector_pressure = self.deflector.item(8, 0).text()
            if deflector_pressure:
                self.activate_deflector.setDisabled(False)


    def show_deflector_column_in_main_table(self, state) -> None:
        if state == 2:
            self.main_table.setColumnHidden(6, False)
            self._join_deflector_column_cells_in_main_table()
        else:
            self.main_table.setColumnHidden(6, True)


    def set_deflector_pressure_in_main_table(self, row=False, column=False) -> None:
        sender = self.sender()
        deflector_pressure = self.deflector.item(8, 0).text()

        if isinstance(sender, QCheckBox):
            if deflector_pressure and row == 2:
                self.main_table.item(0, 6).setText(deflector_pressure)
                self.main_table.item(0, 6).setBackground(QColor(255, 255, 255))
        elif isinstance(sender, QTableWidget):
            if deflector_pressure and row == 8 and column == 0:
                self.main_table.item(0, 6).setText(deflector_pressure)
                self.main_table.item(0, 6).setBackground(QColor(255, 255, 255))
        else:
            self.main_table.item(0, 6).setText('')


    def update_deflector_pressure_in_main_table(self, item) -> None:
        if item.row() == 8 and item.column() == 0:
            deflector_pressure = item.text()
            if self.activate_deflector.isChecked() and deflector_pressure:
                self.main_table.item(0, 6).setText(deflector_pressure)
                self.main_table.item(0, 6).setBackground(QColor(255, 255, 255))
            else:
                self.main_table.item(0, 6).setText('')


    def calculate_available_pressure(self, row=False, column=False) -> None:
        sender = self.sender()
        name = sender.objectName()

        if isinstance(sender, QTableWidget):
            match name:
                case 'deflector':
                    deflector_is_checked = self.activate_deflector.isChecked()
                    if deflector_is_checked:
                        num_rows = self.main_table.rowCount()
                        for row in range(num_rows):
                            deflector_pressure = self.main_table.item(0, 6).text()
                            gravi_pressure = self.main_table.item(row, 5)
                            if all([deflector_pressure, gravi_pressure]) and gravi_pressure.text() != '':
                                deflector_pressure = float(deflector_pressure)
                                gravi_pressure = float(gravi_pressure.text())
                                result = gravi_pressure + deflector_pressure
                                result = "{:.2f}".format(round(result, 2))
                                self.main_table.setItem(row, 7, QTableWidgetItem())
                                self.main_table.item(row, 7).setText(result)
                                self.main_table.item(row, 7).setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                            else:
                                self.main_table.setItem(row, 7, QTableWidgetItem())
                                self.main_table.item(row, 7).setText('')
                    else:
                        num_rows = self.main_table.rowCount()
                        for row in range(num_rows):
                            gravi_pressure = self.main_table.item(row, 5)
                            if gravi_pressure and gravi_pressure.text() != '':
                                result = gravi_pressure.text()
                                self.main_table.setItem(row, 7, QTableWidgetItem())
                                self.main_table.item(row, 7).setText(result)
                                self.main_table.item(row, 7).setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                            else:
                                self.main_table.setItem(row, 7, QTableWidgetItem())
                                self.main_table.item(row, 7).setText('')
                case 'main':
                    if column == 5:
                        deflector_is_checked = self.activate_deflector.isChecked()
                        if deflector_is_checked:
                            deflector_pressure = self.main_table.item(0, 6).text()
                            gravi_pressure = self.main_table.item(row, 5)
                            if all([deflector_pressure, gravi_pressure]) and gravi_pressure.text() != '':
                                deflector_pressure = float(deflector_pressure)
                                gravi_pressure = float(gravi_pressure.text())
                                result = gravi_pressure + deflector_pressure
                                result = "{:.2f}".format(round(result, 2))
                                self.main_table.setItem(row, 7, QTableWidgetItem())
                                self.main_table.item(row, 7).setText(result)
                                self.main_table.item(row, 7).setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                            else:
                                self.main_table.setItem(row, 7, QTableWidgetItem())
                                self.main_table.item(row, 7).setText('')
        elif isinstance(sender, QCheckBox):
            if row == 2:
                num_rows = self.main_table.rowCount()
                for row in range(num_rows):
                    deflector_pressure = self.main_table.item(0, 6).text()
                    gravi_pressure = self.main_table.item(row, 5)
                    if all([deflector_pressure, gravi_pressure]) and gravi_pressure.text() != '':
                        deflector_pressure = float(deflector_pressure)
                        gravi_pressure = float(gravi_pressure.text())
                        result = gravi_pressure + deflector_pressure
                        result = "{:.2f}".format(round(result, 2))
                        self.main_table.setItem(row, 7, QTableWidgetItem())
                        self.main_table.item(row, 7).setText(result)
                        self.main_table.item(row, 7).setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    else:
                        self.main_table.setItem(row, 7, QTableWidgetItem())
                        self.main_table.item(row, 7).setText('')
            else:
                num_rows = self.main_table.rowCount()
                for row in range(num_rows):
                    gravi_pressure = self.main_table.item(row, 5)
                    if gravi_pressure and gravi_pressure.text() != '':
                        result = gravi_pressure.text()
                        self.main_table.setItem(row, 7, QTableWidgetItem())
                        self.main_table.item(row, 7).setText(result)
                        self.main_table.item(row, 7).setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    else:
                        self.main_table.setItem(row, 7, QTableWidgetItem())
                        self.main_table.item(row, 7).setText('')


    def activate_klapan_input(self, value) -> None:
        if value == 'Другой':
            self.klapan_input.setDisabled(False)
            self.klapan_input.setStyleSheet('QLineEdit {background-color: %s}' % QColor(229, 255, 204).name())
        else:
            self.klapan_input.setText('')
            self.klapan_input.setDisabled(True)
            self.klapan_input.setStyleSheet('QLineEdit {background-color: %s}' % QColor(224, 224, 224).name())


















if __name__ == '__main__':
    import sys
    app = QApplication(sys.argv)
    app.setFont(QFont('Consolas', 10))
    window = MainWindow()
    window.setWindowIcon(QIcon(os.path.join(basedir, 'app.ico')))
    window.setIconSize(QSize(15, 15))
    window.show()
    sys.exit(app.exec())
