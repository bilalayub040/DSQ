import sys, re
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QApplication, QWidget, QTabWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QComboBox, QPushButton, QTextEdit, QGroupBox,
    QScrollArea, QTableWidget, QTableWidgetItem
)

class PlanWidget(QGroupBox):
    def __init__(self, remove_callback):
        super().__init__("Plan Section")
        self.remove_callback = remove_callback
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(8, 6, 8, 6)
        layout.setSpacing(5)

        # Top row: Reset + Remove buttons
        top_row = QHBoxLayout()
        self.reset_btn = QPushButton("Reset")
        self.remove_btn = QPushButton("Remove")
        self.reset_btn.setFixedHeight(30)
        self.remove_btn.setFixedHeight(30)
        self.reset_btn.clicked.connect(self.reset_fields)
        self.remove_btn.clicked.connect(self.remove_self)
        top_row.addWidget(self.reset_btn)
        top_row.addWidget(self.remove_btn)
        top_row.addStretch()
        layout.addLayout(top_row)

        # Plan / Addons / Promo / Discount row
        row1 = QHBoxLayout()
        self.plan = QLineEdit(); self.plan.setFixedHeight(30); self.plan.setFixedWidth(150)
        self.addons = QLineEdit(); self.addons.setFixedHeight(30); self.addons.setFixedWidth(150)
        self.promo = QLineEdit(); self.promo.setFixedHeight(30); self.promo.setFixedWidth(150)
        self.discount = QComboBox(); self.discount.setFixedHeight(30); self.discount.setFixedWidth(150)
        self.discount.addItems(["5%", "10%", "15%", "20%", "5%ontop", "10%ontop", "15%ontop", "20%ontop"])
        for lbl_text, widget in [("Plan", self.plan), ("Addons", self.addons), ("Promo", self.promo), ("Discount", self.discount)]:
            lbl = QLabel(lbl_text)
            lbl.setAlignment(Qt.AlignCenter)
            row1.addWidget(lbl)
            row1.addWidget(widget)
        layout.addLayout(row1)

        # Msisdns | Simserial row
        row2 = QHBoxLayout()
        lbl_ms = QLabel("Msisdns"); lbl_ms.setAlignment(Qt.AlignTop)
        lbl_sim = QLabel("Simserials"); lbl_sim.setAlignment(Qt.AlignTop)
        self.msisdns = QTextEdit(); self.msisdns.setFixedHeight(80); self.msisdns.setFixedWidth(180)
        self.simserials = QTextEdit(); self.simserials.setFixedHeight(80); self.simserials.setFixedWidth(180)
        self.msisdns.textChanged.connect(self.format_msisdns)
        self.simserials.textChanged.connect(self.format_simserials)
        col1 = QVBoxLayout(); col1.addWidget(lbl_ms); col1.addWidget(self.msisdns)
        col2 = QVBoxLayout(); col2.addWidget(lbl_sim); col2.addWidget(self.simserials)
        row2.addLayout(col1); row2.addLayout(col2)
        row2.addStretch()
        layout.addLayout(row2)

        self.setLayout(layout)

    def reset_fields(self):
        self.plan.clear()
        self.addons.clear()
        self.promo.clear()
        self.msisdns.clear()
        self.simserials.clear()
        self.discount.setCurrentIndex(0)

    def remove_self(self):
        self.setParent(None)
        self.deleteLater()
        self.remove_callback(self)

    def format_msisdns(self):
        txt = re.sub(r"\D", "", self.msisdns.toPlainText())
        self.msisdns.blockSignals(True)
        self.msisdns.setPlainText(txt)
        self.msisdns.blockSignals(False)

    def format_simserials(self):
        txt = re.sub(r"\D", "", self.simserials.toPlainText())
        self.simserials.blockSignals(True)
        self.simserials.setPlainText(txt)
        self.simserials.blockSignals(False)

class SubmissionTab(QWidget):
    LEFT_WIDTH = 400
    RIGHT_WIDTH = 400

    def __init__(self):
        super().__init__()
        self.plan_widgets = []
        self.initUI()

    def initUI(self):
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setSpacing(10)
        self.main_layout.setAlignment(Qt.AlignTop)

        # Top: Left and Right Columns
        top_row = QHBoxLayout()
        top_row.setSpacing(50)

        # Left
        self.left_column = QVBoxLayout(); self.left_column.setAlignment(Qt.AlignTop)
        self.left_widgets = {}
        left_labels = ["Account Name", "New/Existing", "CR", "QID", "Email", "Dynamic ID", "Type", "Account Number", "Agent ID"]
        for lbl in left_labels: self.add_left_row(lbl)
        self.b_select.currentIndexChanged.connect(self.toggle_B_fields)
        left_widget = QWidget(); left_widget.setLayout(self.left_column); left_widget.setFixedWidth(self.LEFT_WIDTH)
        top_row.addWidget(left_widget)

        # Right
        self.right_column = QVBoxLayout(); self.right_column.setAlignment(Qt.AlignTop)
        self.right_widgets = {}
        right_labels = ["Require Movement", "Parent Account No", "Dynamic ID", "CR", "Existing Revenue"]
        for lbl in right_labels: self.add_right_row(lbl)
        self.require_movement.currentIndexChanged.connect(self.toggle_movement_fields)
        right_widget = QWidget(); right_widget.setLayout(self.right_column); right_widget.setFixedWidth(self.RIGHT_WIDTH)
        top_row.addWidget(right_widget)

        self.main_layout.addLayout(top_row)

        # Plan section
        self.add_plan_btn = QPushButton("Add Plan")
        self.add_plan_btn.setFixedHeight(30); self.add_plan_btn.setFixedWidth(100)
        self.add_plan_btn.clicked.connect(self.add_plan)
        self.main_layout.addWidget(self.add_plan_btn)

        self.plans_area = QScrollArea(); self.plans_area.setWidgetResizable(True)
        self.plans_area.setFixedHeight(250)
        self.plans_holder = QWidget()
        self.plans_layout = QVBoxLayout(self.plans_holder)
        self.plans_layout.setSpacing(5); self.plans_layout.setAlignment(Qt.AlignTop)
        self.plans_area.setWidget(self.plans_holder)
        self.main_layout.addWidget(self.plans_area)

        # Preview Table
        self.preview_table = QTableWidget(); self.preview_table.setColumnCount(7)
        self.preview_table.setHorizontalHeaderLabels(["Account Number", "Account Name", "Agent ID", "Msisdns", "Simserials", "Plan/Addon/Promo/Disc", "Spendlimit"])
        self.preview_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.main_layout.addWidget(self.preview_table)

        self.toggle_B_fields(); self.toggle_movement_fields()

    def add_left_row(self, label_text):
        container = QWidget()
        lbl = QLabel(label_text); lbl.setFixedWidth(140); lbl.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        if label_text == "New/Existing":
            input_widget = QComboBox(); input_widget.addItems(["New", "Existing"]); self.b_select = input_widget
        elif label_text == "Type":
            input_widget = QComboBox(); input_widget.addItems(["Company paid", "Reimbursement"])
        else: input_widget = QLineEdit()
        input_widget.setFixedWidth(240)
        layout = QHBoxLayout(container); layout.setContentsMargins(0,0,0,0)
        layout.addWidget(lbl); layout.addWidget(input_widget); layout.addStretch()
        container.setLayout(layout); self.left_widgets[label_text] = container
        self.left_column.addWidget(container)

    def add_right_row(self, label_text):
        container = QWidget()
        lbl = QLabel(label_text); lbl.setFixedWidth(140); lbl.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        if label_text == "Require Movement": input_widget = QComboBox(); input_widget.addItems(["Yes","No","Not Now"]); self.require_movement = input_widget
        else: input_widget = QLineEdit(); input_widget.setFixedWidth(240)
        layout = QHBoxLayout(container); layout.setContentsMargins(0,0,0,0)
        layout.addWidget(lbl); layout.addWidget(input_widget); layout.addStretch()
        container.setLayout(layout); self.right_widgets[label_text] = container
        self.right_column.addWidget(container)

    def toggle_B_fields(self):
        if self.b_select.currentText() == "New":
            for k in ["CR","QID","Email","Dynamic ID","Type"]: self.left_widgets[k].setVisible(True)
            self.left_widgets["Account Number"].setVisible(False); self.left_widgets["Agent ID"].setVisible(False)
        else:
            for k in ["CR","QID","Email","Dynamic ID","Type"]: self.left_widgets[k].setVisible(False)
            self.left_widgets["Account Number"].setVisible(True); self.left_widgets["Agent ID"].setVisible(True)

    def toggle_movement_fields(self):
        if self.require_movement.currentText() == "Yes":
            for k in ["Parent Account No","Dynamic ID","CR","Existing Revenue"]: self.right_widgets[k].setVisible(True)
        else:
            for k in ["Parent Account No","Dynamic ID","CR","Existing Revenue"]: self.right_widgets[k].setVisible(False)

    def add_plan(self):
        plan = PlanWidget(self.remove_plan)
        self.plans_layout.addWidget(plan)
        self.plan_widgets.append(plan)
        self.update_preview()

    def remove_plan(self, plan):
        if plan in self.plan_widgets: self.plan_widgets.remove(plan)
        self.update_preview()

    def update_preview(self):
        self.preview_table.setRowCount(0)
        # row 1: Account details
        acc_num = self.left_widgets.get("Account Number", QLineEdit()).layout().itemAt(1).widget().text()
        acc_name = self.left_widgets.get("Account Name", QLineEdit()).layout().itemAt(1).widget().text()
        agent_id = self.left_widgets.get("Agent ID", QLineEdit()).layout().itemAt(1).widget().text()
        self.preview_table.setRowCount(2+len(self.plan_widgets))
        self.preview_table.setItem(0,0,QTableWidgetItem("Account Number")); self.preview_table.setItem(0,1,QTableWidgetItem("Account Name")); self.preview_table.setItem(0,2,QTableWidgetItem("Agent ID"))
        self.preview_table.setItem(1,0,QTableWidgetItem(acc_num)); self.preview_table.setItem(1,1,QTableWidgetItem(acc_name)); self.preview_table.setItem(1,2,QTableWidgetItem(agent_id))
        # row 2: plan headers
        self.preview_table.setItem(2,0,QTableWidgetItem("Msisdns")); self.preview_table.setItem(2,1,QTableWidgetItem("Simserials")); self.preview_table.setItem(2,2,QTableWidgetItem("Plan")); self.preview_table.setItem(2,3,QTableWidgetItem("Addons"))
        self.preview_table.setItem(2,4,QTableWidgetItem("Promo")); self.preview_table.setItem(2,5,QTableWidgetItem("Discount")); self.preview_table.setItem(2,6,QTableWidgetItem("Spendlimit"))
        # rows 3+: plan values
        for i, plan in enumerate(self.plan_widgets):
            row = i+3
            self.preview_table.setItem(row,0,QTableWidgetItem(plan.msisdns.toPlainText()))
            self.preview_table.setItem(row,1,QTableWidgetItem(plan.simserials.toPlainText()))
            self.preview_table.setItem(row,2,QTableWidgetItem(plan.plan.text()))
            self.preview_table.setItem(row,3,QTableWidgetItem(plan.addons.text()))
            self.preview_table.setItem(row,4,QTableWidgetItem(plan.promo.text()))
            self.preview_table.setItem(row,5,QTableWidgetItem(plan.discount.currentText()))
            self.preview_table.setItem(row,6,QTableWidgetItem("0.01"))

class MainApp(QTabWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Business App")
        self.sub_tab = SubmissionTab()
        self.addTab(self.sub_tab, "Submission Mobility")
        self.addTab(QWidget(), "Modification")
        self.addTab(QWidget(), "Discount")
        self.addTab(QWidget(), "APS")

def main():
    app = QApplication(sys.argv)
    window = MainApp()
    window.resize(1100,800)
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
