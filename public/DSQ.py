import sys, re, os
import win32com.client  # for Outlook desktop
import pythoncom          # required when using COM from a background thread
import requests
from PyQt5.QtCore import Qt, QSize, QThread, pyqtSignal, QTimer
from datetime import datetime  # Add at the top of the file
from PyQt5.QtGui import QFont, QColor, QBrush
from PyQt5.QtWidgets import (
    QApplication, QWidget, QTabWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QComboBox, QPushButton, QTextEdit, QGroupBox,
    QScrollArea, QTableWidget, QTableWidgetItem, QFileDialog, QSizePolicy,
    QListWidget, QListWidgetItem
)

# ----------------- Background Worker -----------------
class DeferredInitWorker(QThread):
    accounts_loaded = pyqtSignal(list)
    email_lists_loaded = pyqtSignal(list, list)
    finished_all = pyqtSignal()

    def run(self):
        # Initialize COM for this thread before using win32com
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass

        # 1) Try to load Outlook accounts (same logic as original populate_outlook_users but not touching UI)
        accounts = []
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            ns = outlook.GetNamespace("MAPI")
            accounts = [ns.Accounts.Item(i).SmtpAddress for i in range(1, ns.Accounts.Count+1)]
        except Exception as ex:
            # keep accounts empty; main thread will handle UI fallback
            accounts = []
            # optional: print to console for diagnostics
            print("Outlook background error:", ex)

        # emit accounts even if empty
        self.accounts_loaded.emit(accounts)

        # 2) Try to load remote email lists (same logic as load_email_lists)
        to_list = []
        cc_list = []
        try:
            to_txt = requests.get("https://dsq-beta.vercel.app/BO_emails.txt", timeout=6).text
            to_list = [x.strip() for x in to_txt.splitlines() if x.strip()]
        except Exception:
            to_list = []
        try:
            cc_txt = requests.get("https://dsq-beta.vercel.app/CC_emails.txt", timeout=6).text
            cc_list = [x.strip() for x in cc_txt.splitlines() if x.strip()]
        except Exception:
            cc_list = []

        self.email_lists_loaded.emit(to_list, cc_list)

        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

        # signal finish
        self.finished_all.emit()


# ----------------- PlanWidget -----------------
class PlanWidget(QGroupBox):
    def __init__(self, remove_callback):
        super().__init__()
        self.remove_callback = remove_callback
        self.initUI()

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
        self.reset_btn = QPushButton("Reset")
        self.remove_btn = QPushButton("Remove")
        self.reset_btn.setFixedHeight(20)
        self.remove_btn.setFixedHeight(20)
        self.reset_btn.clicked.connect(self.reset_fields)
        self.remove_btn.clicked.connect(self.remove_self)
        top_row.addWidget(self.reset_btn)
        top_row.addWidget(self.remove_btn)
        top_row.addStretch()
        layout.addLayout(top_row)

        # Main horizontal row
        row = QHBoxLayout()
        row.setSpacing(5)
        row.setContentsMargins(0, 0, 0, 0)

        # Plan / Addon / Promo / Discount
        self.plan = QLineEdit(); self.plan.setFixedWidth(225)
        self.addons = QLineEdit(); self.addons.setFixedWidth(225)
        self.promo = QLineEdit(); self.promo.setFixedWidth(150)
        self.discount = QComboBox(); self.discount.setFixedWidth(80)
        self.discount.addItems(["0%", "5%", "10%", "20%", "5%ontop", "10%ontop", "15%ontop", "20%ontop"])

        for lbl_text, widget in [("Plan", self.plan), ("Addon", self.addons), ("Promo", self.promo), ("Discount", self.discount)]:
            col = QVBoxLayout()
            col.setSpacing(4)
            col.setAlignment(Qt.AlignTop)
            lbl = QLabel(lbl_text)
            lbl.setAlignment(Qt.AlignLeft)
            col.addWidget(lbl)
            col.addWidget(widget)
            row.addLayout(col, 0)

        # Msisdns / Simserials
        self.msisdns = QTextEdit(); self.msisdns.setFixedWidth(120); self.msisdns.setFixedHeight(15*18)
        self.simserials = QTextEdit(); self.simserials.setFixedWidth(220); self.simserials.setFixedHeight(15*18)
        self.msisdns.textChanged.connect(self.format_msisdns)
        self.simserials.textChanged.connect(self.format_simserials)

        for lbl_text, widget in [("Msisdns", self.msisdns), ("Simserials", self.simserials)]:
            col = QVBoxLayout()
            col.setSpacing(4)
            col.setAlignment(Qt.AlignTop)
            lbl = QLabel(lbl_text)
            lbl.setAlignment(Qt.AlignLeft)
            col.addWidget(lbl)
            col.addWidget(widget)
            row.addLayout(col, 0)

        row.addStretch(1)

        row_container = QWidget()
        row_container.setLayout(row)
        row_container.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        layout.addWidget(row_container)

        self.setLayout(layout)

    def format_msisdns(self):
        raw = self.msisdns.toPlainText()
        digits = re.sub(r"\D", "", raw)
        blocks = []
        while digits:
            if digits.startswith("9742900"):
                n = 13
            elif digits.startswith("974"):
                n = 11
            elif digits.startswith("2900"):
                n = 10
            else:
                n = 8
            blocks.append(digits[:n])
            digits = digits[n:]
        newtxt = "\n".join(blocks)
        if newtxt != raw:
            self.msisdns.blockSignals(True)
            self.msisdns.setPlainText(newtxt)
            self.msisdns.blockSignals(False)

    def format_simserials(self):
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


# ----------------- SubmissionTab -----------------
class SubmissionTab(QWidget):
    def __init__(self):
        super().__init__()
        self.plan_widgets = []
        self.attachments = []
        self.logged_in_users = []
        self._accounts_loaded_flag = False
        self._emails_loaded_flag = False
        self.initUI()

        # ensure UI is visible immediately, then start deferred init after event loop starts
        QTimer.singleShot(0, self.start_deferred_init)

        # original call preserved: update preview at init (keeps logic same)
        self.update_preview()

    def initUI(self):
        # --- SCROLLABLE OUTER AREA ---
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        outer_container = QWidget()
        scroll.setWidget(outer_container)

        self.outer_layout = QVBoxLayout(outer_container)
        self.outer_layout.setSpacing(10)
        self.outer_layout.setContentsMargins(8,8,8,8)

        main_layout = QVBoxLayout(self)
        main_layout.addWidget(scroll)

        # --- TOP USER SELECTION ---
        top_user_row = QHBoxLayout()
        top_user_row.setSpacing(10)
        user_lbl = QLabel("User:")
        self.user_combo = QComboBox()
        self.user_display = QLabel("")
        top_user_row.addWidget(user_lbl)
        top_user_row.addWidget(self.user_combo)
        top_user_row.addWidget(self.user_display)
        top_user_row.addStretch()
        self.outer_layout.addLayout(top_user_row)

        # NOTE: populate_outlook_users is now deferred to background worker
        self.user_combo.currentTextChanged.connect(self.update_user_display)

        # --- LEFT / RIGHT columns ---
        top_row = QHBoxLayout()
        self.LEFT_WIDTH = 360
        self.left_column = QVBoxLayout()
        self.left_widgets = {}
        for lbl in ["Account Name", "New/Existing", "CR", "QID", "Email", "Dynamic ID", "Type", "Account Number", "Agent ID"]:
            self._add_left_row(lbl)
        left_container = QWidget()
        left_container.setLayout(self.left_column)
        left_container.setFixedWidth(self.LEFT_WIDTH)
        top_row.addWidget(left_container)

        self.RIGHT_WIDTH = 360
        self.right_column = QVBoxLayout()
        self.right_widgets = {}
        for lbl in ["Require Movement", "Parent Account No", "Dynamic ID", "CR", "Existing Revenue"]:
            self._add_right_row(lbl)
        right_container = QWidget()
        right_container.setLayout(self.right_column)
        right_container.setFixedWidth(self.RIGHT_WIDTH)
        top_row.addWidget(right_container)

        top_row.addStretch()
        self.outer_layout.addLayout(top_row)

        # --- PLAN SECTION ---
        self.add_plan_btn = QPushButton("Add Plan")
        self.add_plan_btn.setFixedHeight(30)
        self.add_plan_btn.clicked.connect(self.add_plan)

        plans_header = QHBoxLayout()
        plans_header.addWidget(QLabel("Plan section"))
        plans_header.addStretch()
        plans_header.addWidget(self.add_plan_btn)
        self.outer_layout.addLayout(plans_header)

        self.plans_area = QScrollArea()
        self.plans_area.setWidgetResizable(True)
        self.plans_holder = QWidget()
        self.plans_layout = QVBoxLayout(self.plans_holder)
        self.plans_layout.setContentsMargins(0,0,0,0)
        self.plans_layout.setSpacing(10)
        self.plans_holder.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.plans_area.setWidget(self.plans_holder)
        self.plans_area.setFixedHeight(300)  # visible height
        self.outer_layout.addWidget(self.plans_area)

        # --- EMAIL & ATTACHMENTS SECTION ---
        email_attach_container = QHBoxLayout()
        email_attach_container.setSpacing(10)

        # ----- Left: Email Inputs -----
        email_inputs_layout = QVBoxLayout()
        email_inputs_layout.setSpacing(5)
        for lbl, widget_type in [("Submit to:", QComboBox), ("CC:", QComboBox), ("Others:", QLineEdit)]:
            row = QHBoxLayout()
            row.addWidget(QLabel(lbl), 0, Qt.AlignRight)
            w = widget_type()
            w.setFixedWidth(220)
            if lbl=="Submit to:": self.to_input = w
            elif lbl=="CC:": self.cc_input = w
            else: self.others_input = w
            row.addWidget(w)
            email_inputs_layout.addLayout(row)
        email_inputs_layout.addStretch()
        email_attach_container.addLayout(email_inputs_layout, 0)

        # ----- Right: Attachments -----
        attach_container = QVBoxLayout()
        attach_container.setSpacing(5)
        attach_btn_row = QHBoxLayout()
        self.attach_btn = QPushButton("Add Attachments"); self.attach_btn.setFixedHeight(28)
        self.attach_btn.clicked.connect(self.select_attachments)
        self.clear_btn = QPushButton("Clear Attachments"); self.clear_btn.setFixedHeight(28)
        self.clear_btn.clicked.connect(self.clear_attachments)
        attach_btn_row.addWidget(self.attach_btn)
        attach_btn_row.addWidget(self.clear_btn)
        attach_container.addLayout(attach_btn_row)

        self.attachments_list = QListWidget()
        self.attachments_list.setViewMode(QListWidget.IconMode)
        self.attachments_list.setIconSize(QSize(64,64))
        self.attachments_list.setResizeMode(QListWidget.Adjust)
        self.attachments_list.setSpacing(8)
        self.attachments_list.setMinimumHeight(60)
        self.attachments_list.itemDoubleClicked.connect(self.remove_attachment_item)
        attach_container.addWidget(self.attachments_list)
        email_attach_container.addLayout(attach_container, 1)
        self.outer_layout.addLayout(email_attach_container)

        # NOTE: load_email_lists is deferred to background worker

        # --- GREETING LABEL ---
        self.greeting_lbl = QLabel("Hi team,\n\nPlease action the below:")
        self.outer_layout.addWidget(self.greeting_lbl)
        self.preview_btn = QPushButton("Preview")
        self.preview_btn.clicked.connect(self.update_preview)
        main_layout.addWidget(self.preview_btn)
        # --- PREVIEW TABLE ---
        self.preview_area = QScrollArea()
        self.preview_area.setWidgetResizable(True)  # allows inner widget to expand vertically
        self.preview_table = QTableWidget()
        self.preview_table.setColumnCount(7)
        self.preview_table.setHorizontalHeaderLabels(["Account Number", "Account Name", "Plan","Addons"])
        self.preview_table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)  # fixed width
        self.preview_table.setFixedHeight(300)  # your fixed width
        container = QWidget()
        container_layout = QVBoxLayout(container)
        container_layout.addWidget(self.preview_table)
        container_layout.setContentsMargins(0, 0, 0, 0)
        self.preview_area.setWidget(container)
        self.preview_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.outer_layout.addWidget(self.preview_table)

        # --- SUBMIT BUTTON ROW ---
        submit_row = QHBoxLayout()
        self.submit_btn = QPushButton("Send"); self.submit_btn.setFixedHeight(36); self.submit_btn.setFixedWidth(120)
        self.submit_btn.clicked.connect(self.send_email)
        submit_row.addWidget(self.submit_btn)
        self.status_lbl = QLabel("Ready")
        submit_row.addWidget(self.status_lbl)
        submit_row.addStretch()
        self.outer_layout.addLayout(submit_row)

        # Connect visibility logic
        self.left_widgets["New/Existing"].findChild(QComboBox).currentTextChanged.connect(self.update_left_visibility)
        self.right_widgets["Require Movement"].findChild(QComboBox).currentTextChanged.connect(self.update_right_visibility)
        self.update_left_visibility()
        self.update_right_visibility()

        # --- Loading overlay (shown while background tasks run) ---
        self.loading_overlay = QWidget(self)
        self.loading_overlay.setStyleSheet("background-color: rgba(0,0,0,0.35);")
        overlay_layout = QVBoxLayout(self.loading_overlay)
        overlay_layout.setContentsMargins(0,0,0,0)
        overlay_layout.setAlignment(Qt.AlignCenter)
        self.loading_label = QLabel("Loading..."); self.loading_label.setStyleSheet("color: white;")
        font = QFont(); font.setPointSize(14); font.setBold(True)
        self.loading_label.setFont(font)
        overlay_layout.addWidget(self.loading_label)
        self.loading_overlay.raise_()
        self.loading_overlay.setVisible(True)

    def resizeEvent(self, ev):
        # keep overlay covering entire widget
        try:
            self.loading_overlay.setGeometry(0, 0, self.width(), self.height())
        except Exception:
            pass
        return super().resizeEvent(ev)

    # ----------------- DEFERRED INIT -----------------
    def start_deferred_init(self):
        # start worker that will run populate_outlook_users and load_email_lists without blocking UI
        self._worker = DeferredInitWorker()
        self._worker.accounts_loaded.connect(self.on_accounts_loaded)
        self._worker.email_lists_loaded.connect(self.on_email_lists_loaded)
        self._worker.finished_all.connect(self.on_deferred_finished)
        self._worker.start()

    def on_accounts_loaded(self, accounts):
        # Mirror original populate_outlook_users behavior but executed in main thread (UI update)
        try:
            self.logged_in_users = accounts
            if accounts:
                # clear possible placeholder and add accounts
                self.user_combo.clear()
                self.user_combo.addItems(accounts)
                # show first account display (same behavior as original)
                self.update_user_display(accounts[0])
            else:
                # if no accounts found, present original fallback
                self.user_combo.clear()
                self.user_combo.addItem("No Outlook Account")
                self.user_display.setText("")
        except Exception as ex:
            print("Error updating accounts in UI:", ex)
        finally:
            self._accounts_loaded_flag = True

    def on_email_lists_loaded(self, to_list, cc_list):
        try:
            # update the to and cc comboboxes similar to original load_email_lists
            if to_list:
                # clear existing items then add
                try:
                    self.to_input.clear()
                except Exception:
                    pass
                self.to_input.addItems(to_list)
            if cc_list:
                try:
                    self.cc_input.clear()
                except Exception:
                    pass
                self.cc_input.addItems(cc_list)
        except Exception as ex:
            print("Error updating email lists in UI:", ex)
        finally:
            self._emails_loaded_flag = True

    def on_deferred_finished(self):
        # both tasks finished (worker finished) - hide overlay
        self.loading_overlay.setVisible(False)
        # preserve behavior: status label remains "Ready" unless changed elsewhere
        # nothing else changes in logic - background initializations have populated UI elements
        # mark worker for GC
        try:
            self._worker.quit()
            self._worker.wait(100)
        except Exception:
            pass

    # ----------------- OUTLOOK USER (original function removed from init but logic preserved) -----------------
    # (populate_outlook_users originally existed; moved to background worker but UI update is preserved above)

    def update_user_display(self, email):
        if "@" in email:
            local_part = email.split("@")[0]
            parts = local_part.split(".")
            name = " ".join([p.capitalize() for p in parts])
            self.user_display.setText(name)
        else:
            self.user_display.setText(email)

    # ----------------- LEFT / RIGHT HELPERS -----------------
    def _add_left_row(self, label_text):
        container = QWidget()
        lbl = QLabel(label_text)
        lbl.setFixedWidth(120)
        lbl.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        if label_text == "New/Existing":
            input_widget = QComboBox(); input_widget.addItems(["New", "Existing"])
        elif label_text == "Type":
            input_widget = QComboBox(); input_widget.addItems(["Company paid", "Reimbursement"])
        else:
            input_widget = QLineEdit(); input_widget.setFixedWidth(200)
        layout = QHBoxLayout(container); layout.setContentsMargins(0,0,0,0)
        layout.addWidget(lbl); layout.addWidget(input_widget); layout.addStretch()
        container.setLayout(layout)
        self.left_column.addWidget(container)
        self.left_widgets[label_text] = container

    def _add_right_row(self, label_text):
        container = QWidget()
        lbl = QLabel(label_text)
        lbl.setFixedWidth(120)
        lbl.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        if label_text == "Require Movement":
            input_widget = QComboBox(); input_widget.addItems(["Yes","No","Not Now"])
        else:
            input_widget = QLineEdit(); input_widget.setFixedWidth(200)
        layout = QHBoxLayout(container); layout.setContentsMargins(0,0,0,0)
        layout.addWidget(lbl); layout.addWidget(input_widget); layout.addStretch()
        container.setLayout(layout)
        self.right_column.addWidget(container)
        self.right_widgets[label_text] = container

    def update_left_visibility(self):
        val = self.left_widgets["New/Existing"].findChild(QComboBox).currentText()
        show_labels = ["CR","QID","Email","Dynamic ID","Type"] if val=="New" else []
        for lbl in ["CR","QID","Email","Dynamic ID","Type"]:
            self.left_widgets[lbl].setVisible(lbl in show_labels)

    def update_right_visibility(self):
        val = self.right_widgets["Require Movement"].findChild(QComboBox).currentText()
        show_labels = ["Parent Account No","Dynamic ID","CR","Existing Revenue"] if val=="Yes" else []
        for lbl in ["Parent Account No","Dynamic ID","CR","Existing Revenue"]:
            self.right_widgets[lbl].setVisible(lbl in show_labels)

    # ----------------- PLAN -----------------
    def add_plan(self):
        plan = PlanWidget(self.remove_plan)
        plan.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.plans_layout.addWidget(plan)
        self.plan_widgets.append(plan)

    def remove_plan(self, plan):
        if plan in self.plan_widgets:
            self.plan_widgets.remove(plan)
            self.plans_layout.removeWidget(plan)
            plan.deleteLater()

    # ----------------- ATTACHMENTS -----------------
    def select_attachments(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Select Files")
        for f in files:
            if f not in self.attachments:
                self.attachments.append(f)
                item = QListWidgetItem(os.path.basename(f))
                item.setData(Qt.UserRole, f)
                self.attachments_list.addItem(item)

    def clear_attachments(self):
        self.attachments = []
        self.attachments_list.clear()

    def remove_attachment_item(self, item):
        path = item.data(Qt.UserRole)
        if path in self.attachments:
            self.attachments.remove(path)
        self.attachments_list.takeItem(self.attachments_list.row(item))

    # ----------------- PREVIEW -----------------
    def _get_left_input_text(self, label):
        cont = self.left_widgets.get(label)
        if not cont: return ""
        le = cont.findChild(QLineEdit)
        return le.text() if le else ""

    def update_preview(self):
        self.preview_table.setRowCount(0)
        self.preview_table.horizontalHeader().setVisible(False)
        self.preview_table.verticalHeader().setVisible(False)

        acc_num = self._get_left_input_text("Account Number")
        acc_name = self._get_left_input_text("Account Name")
        agent_id = self._get_left_input_text("Agent ID")

        self.preview_table.resizeRowsToContents()

        bold_font = QFont(); bold_font.setBold(True)
        yellow_brush = QBrush(QColor("yellow"))
        black_brush = QBrush(QColor("black"))

        # Default brushes (white background + black text)
        default_fg = QBrush(QColor("black"))
        default_bg = QBrush(QColor("white"))

        # --- Row 0: titles ---
        self.preview_table.setRowCount(3)
        for col, text in enumerate(["Account Number", "Account Name", "Agent ID"]):
            item = QTableWidgetItem(text)
            item.setFont(bold_font)
            item.setBackground(yellow_brush)
            item.setForeground(black_brush)
            self.preview_table.setItem(0, col, item)

        # --- Row 1: account values (force default white/black) ---
        self.preview_table.setItem(1, 0, QTableWidgetItem(acc_num))
        self.preview_table.item(1, 0).setForeground(default_fg)
        self.preview_table.item(1, 0).setBackground(default_bg)

        self.preview_table.setItem(1, 1, QTableWidgetItem(acc_name))
        self.preview_table.item(1, 1).setForeground(default_fg)
        self.preview_table.item(1, 1).setBackground(default_bg)

        self.preview_table.setItem(1, 2, QTableWidgetItem(agent_id))
        self.preview_table.item(1, 2).setForeground(default_fg)
        self.preview_table.item(1, 2).setBackground(default_bg)

        # --- Row 2: plan headers ---
        headers = ["Msisdns", "Simserials", "Plan", "Addon", "Promo", "Discount", "Spendlimit"]
        for c, h in enumerate(headers):
            item = QTableWidgetItem(h)
            item.setFont(bold_font)
            item.setBackground(yellow_brush)
            item.setForeground(black_brush)
            self.preview_table.setItem(2, c, item)

        # --- Row 3+: plan data (force default white/black) ---
        row_idx = 3
        for plan in self.plan_widgets:
            ms_list = [s for s in plan.msisdns.toPlainText().splitlines() if s.strip()]
            sim_list = [s for s in plan.simserials.toPlainText().splitlines() if s.strip()]
            max_lines = max(len(ms_list), len(sim_list), 1)
            for i in range(max_lines):
                self.preview_table.insertRow(row_idx)
                self.preview_table.setItem(row_idx, 0, QTableWidgetItem(ms_list[i] if i < len(ms_list) else ""))
                self.preview_table.setItem(row_idx, 1, QTableWidgetItem(sim_list[i] if i < len(sim_list) else ""))
                self.preview_table.setItem(row_idx, 2, QTableWidgetItem(plan.plan.text()))
                self.preview_table.setItem(row_idx, 3, QTableWidgetItem(plan.addons.text()))
                self.preview_table.setItem(row_idx, 4, QTableWidgetItem(plan.promo.text()))
                self.preview_table.setItem(row_idx, 5, QTableWidgetItem(plan.discount.currentText()))
                self.preview_table.setItem(row_idx, 6, QTableWidgetItem("0.01"))

                # Force default style (white background, black text)
                for c in range(7):
                    cell = self.preview_table.item(row_idx, c)
                    if cell:
                        cell.setForeground(default_fg)
                        cell.setBackground(default_bg)

                self.preview_table.setRowHeight(row_idx, 20)
                row_idx += 1

        self.preview_table.resizeColumnsToContents()

    # ----------------- SEND EMAIL -----------------

    def send_email(self):
        self.update_preview()
        to_email = self.to_input.currentText().strip()
        cc_email = self.cc_input.currentText().strip()

        sender_email = self.user_combo.currentText()
        if sender_email:
            cc_email = f"{cc_email}; {sender_email}" if cc_email else sender_email

        others_raw = self.others_input.text().strip()
        others_emails = re.split(r"[ ,;]+", others_raw) if others_raw else []

        all_to = [to_email] + others_emails

        # --- Generate Subject ---
        account_name = self._get_left_input_text("Account Name") or "Unknown Account"
        today_str = datetime.today().strftime("%Y-%m-%d")
        subject_line = f"{account_name} - Mobility Submission - {today_str}"

        # --- Generate HTML with formatting ---
        html = "<p>Hi team,<br><br>Please action the below:</p>"
        html += "<table border='1' cellspacing='0' cellpadding='4' style='border-collapse: collapse;'>"

        for r in range(self.preview_table.rowCount()):
            html += "<tr>"
            for c in range(self.preview_table.columnCount()):
                item = self.preview_table.item(r, c)
                txt = item.text() if item else ""
                style = ""

                if item:
                    # Bold text
                    if item.font().bold():
                        style += "font-weight:bold;"
                    # Background color
                    bg = item.background().color()
                    style += f"background-color: rgb({bg.red()},{bg.green()},{bg.blue()});"
                    # Optional: text color
                    fg = item.foreground().color()
                    style += f"color: rgb({fg.red()},{fg.green()},{fg.blue()});"

                html += f"<td style='{style}'>{txt}</td>"
            html += "</tr>"

        html += "</table>"

        try:
            self.status_lbl.setText("Sending...")
            QApplication.processEvents()
            outlook = win32com.client.Dispatch("Outlook.Application")
            for to in all_to:
                if not to.strip(): continue
                mail = outlook.CreateItem(0)
                mail.To = to.strip()
                mail.CC = cc_email
                mail.Subject = subject_line
                mail.HTMLBody = html
                for f in self.attachments:
                    try:
                        mail.Attachments.Add(f)
                    except:
                        pass
                mail.Send()
            self.status_lbl.setText("Sent successfully ✓")
        except Exception as ex:
            self.status_lbl.setText("Send failed: " + str(ex))


# ----------------- MainApp -----------------
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
    window.resize(1200, 820)
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
