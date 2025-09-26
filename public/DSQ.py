import sys, re, os, requests, win32com.client
from datetime import datetime
from PyQt5.QtCore import Qt, QSize, QThread, pyqtSignal, QTimer
from PyQt5.QtGui import QFont, QColor, QBrush
from PyQt5.QtWidgets import (
    QApplication, QWidget, QTabWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QComboBox, QPushButton, QTextEdit, QGroupBox,
    QScrollArea, QTableWidget, QTableWidgetItem, QFileDialog, QSizePolicy,
    QListWidget, QListWidgetItem
)

# ----------------- THREADS -----------------
class EmailLoaderThread(QThread):
    emails_loaded = pyqtSignal(list, list)
    def run(self):
        to_list, cc_list = [], []
        try:
            to_txt = requests.get("https://dsq-beta.vercel.app/BO_emails.txt", timeout=3).text
            to_list = [x.strip() for x in to_txt.splitlines() if x.strip()]
        except: pass
        try:
            cc_txt = requests.get("https://dsq-beta.vercel.app/CC_emails.txt", timeout=3).text
            cc_list = [x.strip() for x in cc_txt.splitlines() if x.strip()]
        except: pass
        self.emails_loaded.emit(to_list, cc_list)

class OutlookLoaderThread(QThread):
    users_loaded = pyqtSignal(list)
    def run(self):
        accounts = []
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            ns = outlook.GetNamespace("MAPI")
            accounts = [ns.Accounts.Item(i).SmtpAddress for i in range(1, ns.Accounts.Count+1)]
        except: pass
        self.users_loaded.emit(accounts)

# ----------------- PLAN WIDGET -----------------
class PlanWidget(QGroupBox):
    def __init__(self, remove_callback):
        super().__init__()
        self.remove_callback = remove_callback
        self.initUI()
        # Debounce timers
        self.ms_timer = QTimer()
        self.ms_timer.setSingleShot(True)
        self.ms_timer.timeout.connect(self._format_msisdns)
        self.sim_timer = QTimer()
        self.sim_timer.setSingleShot(True)
        self.sim_timer.timeout.connect(self._format_simserials)

    def reset_fields(self):
        try:
            self.plan.clear()
            self.addons.clear()
            self.promo.clear()
            self.msisdns.clear()
            self.simserials.clear()
            self.discount.setCurrentIndex(0)
        except Exception:
            pass

    def remove_self(self):
        self.setParent(None)
        self.deleteLater()
        self.remove_callback(self)

    def initUI(self):
        self.setTitle("Plan Section")
        layout = QVBoxLayout()
        layout.setContentsMargins(3, 3, 3, 3)
        layout.setSpacing(5)

        # Top buttons
        top_row = QHBoxLayout()
        self.reset_btn = QPushButton("Reset"); self.remove_btn = QPushButton("Remove")
        self.reset_btn.setFixedHeight(20); self.remove_btn.setFixedHeight(20)
        self.reset_btn.clicked.connect(self.reset_fields)
        self.remove_btn.clicked.connect(self.remove_self)
        top_row.addWidget(self.reset_btn)
        top_row.addWidget(self.remove_btn)
        top_row.addStretch()
        layout.addLayout(top_row)

        # Main horizontal row
        row = QHBoxLayout(); row.setSpacing(5); row.setContentsMargins(0,0,0,0)

        # Plan / Addon / Promo / Discount
        self.plan = QLineEdit(); self.plan.setFixedWidth(225)
        self.addons = QLineEdit(); self.addons.setFixedWidth(225)
        self.promo = QLineEdit(); self.promo.setFixedWidth(150)
        self.discount = QComboBox(); self.discount.setFixedWidth(80)
        self.discount.addItems(["0%", "5%", "10%", "20%", "5%ontop", "10%ontop", "15%ontop", "20%ontop"])

        for lbl_text, widget in [("Plan", self.plan), ("Addon", self.addons),
                                 ("Promo", self.promo), ("Discount", self.discount)]:
            col = QVBoxLayout(); col.setSpacing(4); col.setAlignment(Qt.AlignTop)
            lbl = QLabel(lbl_text); lbl.setAlignment(Qt.AlignLeft)
            col.addWidget(lbl); col.addWidget(widget)
            row.addLayout(col, 0)

        # Msisdns / Simserials
        self.msisdns = QTextEdit(); self.msisdns.setFixedWidth(120); self.msisdns.setFixedHeight(15*18)
        self.simserials = QTextEdit(); self.simserials.setFixedWidth(220); self.simserials.setFixedHeight(15*18)
        self.msisdns.textChanged.connect(lambda: self.ms_timer.start(200))
        self.simserials.textChanged.connect(lambda: self.sim_timer.start(200))

        for lbl_text, widget in [("Msisdns", self.msisdns), ("Simserials", self.simserials)]:
            col = QVBoxLayout(); col.setSpacing(4); col.setAlignment(Qt.AlignTop)
            lbl = QLabel(lbl_text); lbl.setAlignment(Qt.AlignLeft)
            col.addWidget(lbl); col.addWidget(widget)
            row.addLayout(col, 0)

        row.addStretch(1)
        container = QWidget(); container.setLayout(row)
        container.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        layout.addWidget(container)
        self.setLayout(layout)

    # ----------------- FORMATTING -----------------
    def _format_msisdns(self):
        raw = self.msisdns.toPlainText()
        digits = re.sub(r"\D", "", raw)
        blocks = []
        while digits:
            if digits.startswith("9742900"): n=13
            elif digits.startswith("974"): n=11
            elif digits.startswith("2900"): n=10
            else: n=8
            blocks.append(digits[:n])
            digits = digits[n:]
        newtxt = "\n".join(blocks)
        if newtxt != raw:
            self.msisdns.blockSignals(True)
            self.msisdns.setPlainText(newtxt)
            self.msisdns.blockSignals(False)

    def _format_simserials(self):
        raw = self.simserials.toPlainText()
        digits = re.sub(r"\D", "", raw)
        blocks = []
        while digits:
            blocks.append(digits[:19])
            digits = digits[19:]
        newtxt = "\n".join(blocks)
        if newtxt != raw:
            self.simserials.blockSignals(True)
            self.simserials.setPlainText(newtxt)
            self.simserials.blockSignals(False)

# ----------------- SUBMISSION TAB -----------------
class SubmissionTab(QWidget):
    def __init__(self):
        super().__init__()
        self.plan_widgets = []
        self.attachments = []
        self.logged_in_users = []
        self.initUI()
        self.update_preview()

        # Start background threads
        self.email_thread = EmailLoaderThread()
        self.email_thread.emails_loaded.connect(self.populate_email_lists)
        self.email_thread.start()

        self.outlook_thread = OutlookLoaderThread()
        self.outlook_thread.users_loaded.connect(self.populate_outlook_users)
        self.outlook_thread.start()

    def initUI(self):
        self.setUpdatesEnabled(False)  # prevent slow repaint during setup
        scroll = QScrollArea(); scroll.setWidgetResizable(True)
        outer_container = QWidget(); scroll.setWidget(outer_container)
        self.outer_layout = QVBoxLayout(outer_container); self.outer_layout.setSpacing(10); self.outer_layout.setContentsMargins(8,8,8,8)
        main_layout = QVBoxLayout(self); main_layout.addWidget(scroll)

        # --- TOP USER SELECTION ---
        top_user_row = QHBoxLayout(); top_user_row.setSpacing(10)
        user_lbl = QLabel("User:"); self.user_combo = QComboBox(); self.user_display = QLabel("")
        top_user_row.addWidget(user_lbl); top_user_row.addWidget(self.user_combo); top_user_row.addWidget(self.user_display)
        top_user_row.addStretch(); self.outer_layout.addLayout(top_user_row)
        self.user_combo.currentTextChanged.connect(self.update_user_display)

        # --- LEFT / RIGHT COLUMNS ---
        self.LEFT_WIDTH, self.RIGHT_WIDTH = 360, 360
        self.left_column, self.right_column = QVBoxLayout(), QVBoxLayout()
        self.left_widgets, self.right_widgets = {}, {}

        for lbl in ["Account Name","New/Existing","CR","QID","Email","Dynamic ID","Type","Account Number","Agent ID"]:
            self._add_left_row(lbl)
        for lbl in ["Require Movement","Parent Account No","Dynamic ID","CR","Existing Revenue"]:
            self._add_right_row(lbl)

        left_container, right_container = QWidget(), QWidget()
        left_container.setLayout(self.left_column); left_container.setFixedWidth(self.LEFT_WIDTH)
        right_container.setLayout(self.right_column); right_container.setFixedWidth(self.RIGHT_WIDTH)

        top_row = QHBoxLayout(); top_row.addWidget(left_container); top_row.addWidget(right_container); top_row.addStretch()
        self.outer_layout.addLayout(top_row)

        # Connect visibility logic
        self.left_widgets["New/Existing"].findChild(QComboBox).currentTextChanged.connect(self.update_left_visibility)
        self.right_widgets["Require Movement"].findChild(QComboBox).currentTextChanged.connect(self.update_right_visibility)
        self.update_left_visibility(); self.update_right_visibility()

        # --- PLAN SECTION ---
        self.add_plan_btn = QPushButton("Add Plan"); self.add_plan_btn.setFixedHeight(30); self.add_plan_btn.clicked.connect(self.add_plan)
        plans_header = QHBoxLayout(); plans_header.addWidget(QLabel("Plan section")); plans_header.addStretch(); plans_header.addWidget(self.add_plan_btn)
        self.outer_layout.addLayout(plans_header)

        self.plans_area = QScrollArea(); self.plans_area.setWidgetResizable(True)
        self.plans_holder = QWidget(); self.plans_layout = QVBoxLayout(self.plans_holder); self.plans_layout.setContentsMargins(0,0,0,0); self.plans_layout.setSpacing(10)
        self.plans_holder.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed); self.plans_area.setWidget(self.plans_holder)
        self.plans_area.setFixedHeight(300); self.outer_layout.addWidget(self.plans_area)

        # --- EMAIL & ATTACHMENTS ---
        email_attach_container = QHBoxLayout(); email_attach_container.setSpacing(10)
        email_inputs_layout = QVBoxLayout(); email_inputs_layout.setSpacing(5)

        for lbl, widget_type in [("Submit to:", QComboBox), ("CC:", QComboBox), ("Others:", QLineEdit)]:
            row = QHBoxLayout(); row.addWidget(QLabel(lbl), 0, Qt.AlignRight)
            w = widget_type(); w.setFixedWidth(220)
            if lbl=="Submit to:": self.to_input=w
            elif lbl=="CC:": self.cc_input=w
            else: self.others_input=w
            row.addWidget(w); email_inputs_layout.addLayout(row)
        email_inputs_layout.addStretch(); email_attach_container.addLayout(email_inputs_layout,0)

        attach_container = QVBoxLayout(); attach_container.setSpacing(5)
        attach_btn_row = QHBoxLayout(); self.attach_btn = QPushButton("Add Attachments"); self.attach_btn.setFixedHeight(28)
        self.attach_btn.clicked.connect(self.select_attachments)
        self.clear_btn = QPushButton("Clear Attachments"); self.clear_btn.setFixedHeight(28); self.clear_btn.clicked.connect(self.clear_attachments)
        attach_btn_row.addWidget(self.attach_btn); attach_btn_row.addWidget(self.clear_btn)
        attach_container.addLayout(attach_btn_row)
        self.attachments_list = QListWidget(); self.attachments_list.setViewMode(QListWidget.IconMode)
        self.attachments_list.setIconSize(QSize(64,64)); self.attachments_list.setResizeMode(QListWidget.Adjust)
        self.attachments_list.setSpacing(8); self.attachments_list.setMinimumHeight(60)
        self.attachments_list.itemDoubleClicked.connect(self.remove_attachment_item)
        attach_container.addWidget(self.attachments_list)
        email_attach_container.addLayout(attach_container,1)
        self.outer_layout.addLayout(email_attach_container)

        # --- GREETING & PREVIEW ---
        self.greeting_lbl = QLabel("Hi team,\n\nPlease action the below:"); self.outer_layout.addWidget(self.greeting_lbl)
        self.preview_btn = QPushButton("Preview"); self.preview_btn.clicked.connect(self.update_preview); main_layout.addWidget(self.preview_btn)

        self.preview_table = QTableWidget(); self.preview_table.setColumnCount(7)
        self.preview_table.setHorizontalHeaderLabels(["Account Number","Account Name","Plan","Addons","Promo","Discount","Spendlimit"])
        self.preview_table.setEditTriggers(QTableWidget.NoEditTriggers); self.preview_table.setFixedHeight(300)
        self.outer_layout.addWidget(self.preview_table)

        # --- SUBMIT BUTTON ---
        submit_row = QHBoxLayout(); self.submit_btn = QPushButton("Send"); self.submit_btn.setFixedHeight(36); self.submit_btn.setFixedWidth(120)
        self.submit_btn.clicked.connect(self.send_email); submit_row.addWidget(self.submit_btn)
        self.status_lbl = QLabel("Ready"); submit_row.addWidget(self.status_lbl); submit_row.addStretch()
        self.outer_layout.addLayout(submit_row)
        self.setUpdatesEnabled(True)

    # ----------------- HELPER METHODS -----------------
    def _add_left_row(self,label_text):
        container = QWidget(); lbl = QLabel(label_text); lbl.setFixedWidth(120); lbl.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        if label_text=="New/Existing": input_widget=QComboBox(); input_widget.addItems(["New","Existing"])
        elif label_text=="Type": input_widget=QComboBox(); input_widget.addItems(["Company paid","Reimbursement"])
        else: input_widget=QLineEdit(); input_widget.setFixedWidth(200)
        layout = QHBoxLayout(container); layout.setContentsMargins(0,0,0,0)
        layout.addWidget(lbl); layout.addWidget(input_widget); layout.addStretch(); container.setLayout(layout)
        self.left_column.addWidget(container); self.left_widgets[label_text]=container

    def _add_right_row(self,label_text):
        container = QWidget(); lbl = QLabel(label_text); lbl.setFixedWidth(120); lbl.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        if label_text=="Require Movement": input_widget=QComboBox(); input_widget.addItems(["Yes","No","Not Now"])
        else: input_widget=QLineEdit(); input_widget.setFixedWidth(200)
        layout = QHBoxLayout(container); layout.setContentsMargins(0,0,0,0)
        layout.addWidget(lbl); layout.addWidget(input_widget); layout.addStretch(); container.setLayout(layout)
        self.right_column.addWidget(container); self.right_widgets[label_text]=container

    def update_left_visibility(self):
        val=self.left_widgets["New/Existing"].findChild(QComboBox).currentText()
        show_labels=["CR","QID","Email","Dynamic ID","Type"] if val=="New" else []
        for lbl in ["CR","QID","Email","Dynamic ID","Type"]: self.left_widgets[lbl].setVisible(lbl in show_labels)

    def update_right_visibility(self):
        val=self.right_widgets["Require Movement"].findChild(QComboBox).currentText()
        show_labels=["Parent Account No","Dynamic ID","CR","Existing Revenue"] if val=="Yes" else []
        for lbl in ["Parent Account No","Dynamic ID","CR","Existing Revenue"]: self.right_widgets[lbl].setVisible(lbl in show_labels)

    # ----------------- PLANS -----------------
    def add_plan(self):
        self.setUpdatesEnabled(False)
        plan = PlanWidget(self.remove_plan); plan.setSizePolicy(QSizePolicy.Expanding,QSizePolicy.Fixed)
        self.plans_layout.addWidget(plan); self.plan_widgets.append(plan)
        self.setUpdatesEnabled(True)

    def remove_plan(self, plan):
        if plan in self.plan_widgets:
            self.plan_widgets.remove(plan); self.plans_layout.removeWidget(plan); plan.deleteLater()

    # ----------------- ATTACHMENTS -----------------
    def select_attachments(self):
        files,_=QFileDialog.getOpenFileNames(self,"Select Files")
        for f in files:
            if f not in self.attachments:
                self.attachments.append(f)
                item = QListWidgetItem(os.path.basename(f)); item.setData(Qt.UserRole,f)
                self.attachments_list.addItem(item)

    def clear_attachments(self):
        self.attachments=[]; self.attachments_list.clear()

    def remove_attachment_item(self,item):
        path=item.data(Qt.UserRole)
        if path in self.attachments: self.attachments.remove(path)
        self.attachments_list.takeItem(self.attachments_list.row(item))

    # ----------------- EMAIL LOADING -----------------
    def populate_email_lists(self,to_list,cc_list):
        self.to_input.addItems(to_list); self.cc_input.addItems(cc_list)

    def populate_outlook_users(self, accounts):
        self.user_combo.clear()
        if accounts: self.user_combo.addItems(accounts); self.update_user_display(accounts[0])
        else: self.user_combo.addItem("No Outlook Account"); self.user_display.setText("")

    def update_user_display(self,email):
        if "@" in email: parts=email.split("@")[0].split("."); name=" ".join([p.capitalize() for p in parts]); self.user_display.setText(name)
        else: self.user_display.setText(email)

    # ----------------- PREVIEW -----------------
    def _get_left_input_text(self,label):
        cont=self.left_widgets.get(label)
        if not cont: return ""
        le=cont.findChild(QLineEdit)
        return le.text() if le else ""

    def update_preview(self):
        self.preview_table.setUpdatesEnabled(False)
        self.preview_table.clearContents()
        self.preview_table.setRowCount(3)
        bold_font=QFont(); bold_font.setBold(True); yellow_brush=QBrush(QColor("yellow")); black_brush=QBrush(QColor("black"))

        for col,text in enumerate(["Account Number","Account Name","Agent ID"]):
            item=QTableWidgetItem(text); item.setFont(bold_font); item.setBackground(yellow_brush); item.setForeground(black_brush)
            self.preview_table.setItem(0,col,item)

        acc_num=self._get_left_input_text("Account Number"); acc_name=self._get_left_input_text("Account Name"); agent_id=self._get_left_input_text("Agent ID")
        self.preview_table.setItem(1,0,QTableWidgetItem(acc_num)); self.preview_table.setItem(1,1,QTableWidgetItem(acc_name)); self.preview_table.setItem(1,2,QTableWidgetItem(agent_id))
        self.preview_table.resizeColumnsToContents(); self.preview_table.setUpdatesEnabled(True)

    # ----------------- SEND EMAIL -----------------
    def send_email(self):
        self.status_lbl.setText("Sending...")
        # Implement Outlook sending logic here
        self.status_lbl.setText("Email sent!")

# ----------------- MAIN APP -----------------
if __name__=="__main__":
    app = QApplication(sys.argv)
    window = SubmissionTab(); window.resize(900,800); window.show()
    sys.exit(app.exec_())
