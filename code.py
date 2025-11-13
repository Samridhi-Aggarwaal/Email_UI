import sys
import win32com.client as win32
from PyQt5.QtGui import QFont, QColor, QIcon, QPalette
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton, QScrollArea,
    QVBoxLayout, QHBoxLayout, QMessageBox, QCheckBox, QComboBox, QFrame
)


RECIPIENT_GROUPS = {
  "CDT Block": ["nikhila.bharath@fractal.ai", "kashmira.magarde@fractal.ai"],
  "DS Block": ["virag.jhaveri@fractal.ai", "aniket.shetty@fractal.ai"]
}
CC_LIST = [
  "rahul.roychowdhury@fractal.ai","ankit.gupta@fractal.ai",
  "nayana.ck@fractal.ai"
]
ROLES = [
  "Custom", "Power BI", "Fullstack", "Devops", "Scrum", "TPM", "UI/UX designer",
  "Motion designer", "RPA developer", "QA tester", "DS", "gen AI DS",
  "BA-Supply chain", "Azure Data Engineer"
]
COMPANIES = ["Custom", "Kenvue", "Lowes", "Jim Beam"]


class OutlookEmailApp(QWidget):
  def __init__(self):
    super().__init__()
    self.setWindowTitle("Outlook Email Sender")
    self.setMinimumSize(900, 700)
    self.setup_ui()
    self.set_styles()
    self.show()

  def set_styles(self):
    self.setStyleSheet("""
          QWidget {
              background-color: #f0f0f0;
              font-family: Arial;
              font-size: 14px;
          }
          QLabel {
              font-weight: bold;
          }
          QLineEdit, QComboBox, QCheckBox {
              padding: 4px;
              background-color: #ffffff;
              border: 1px solid #c5c5c5;
          }
          QCheckBox {
              spacing: 5px;
          }
          QPushButton {
              background-color: #0078D7;
              color: white;
              font-weight: bold;
              border-radius: 6px;
              padding: 6px 12px;
          }
          QPushButton#deleteButton {
              background-color: #e57373;
          }
          QPushButton:hover {
              background-color: #005bb5;
          }
          QPushButton:pressed {
              background-color: #004a94;
          }
      """)
    
  def setup_ui(self):
    layout = QVBoxLayout()
    self.setLayout(layout)
    # Initialize each section of the UI
    self.setup_recipient_section(layout)
    self.setup_cc_section(layout)
    self.setup_sender_section(layout)
    self.setup_employee_section(layout)
    self.setup_action_buttons(layout)

  def setup_recipient_section(self, layout):
    self.add_section_header(layout, "Recipients")
    to_row = QHBoxLayout()
    to_label = QLabel("To:")
    self.custom_to_entry = self.create_line_edit("Enter custom IDs, semicolon-separated")
    self.to_dropdown = QComboBox()
    self.to_dropdown.addItems(["CDT Block", "DS Block", "Custom"])
    self.to_dropdown.setCurrentIndex(2)
    self.to_dropdown.currentTextChanged.connect(self.handle_to_dropdown)
    to_row.addWidget(to_label)
    to_row.addWidget(self.to_dropdown)
    to_row.addWidget(self.custom_to_entry)
    layout.addLayout(to_row)
    self.to_popup_shown = False
    self.custom_to_entry.textEdited.connect(self.show_to_instruction)

  def setup_cc_section(self, layout):
    self.add_section_header(layout, "CC")
    self.cc_value = "; ".join(self.CC_LIST)
    self.cc_entry = self.add_row(layout, "Cc:", self.cc_value)
    self.cc_entry.setReadOnly(False)
    self.cc_entry.setToolTip("Enter CC emails separated by semicolons (;)")
    self.cc_popup_shown = False
    self.cc_entry.textEdited.connect(self.show_cc_instruction)

  def setup_sender_section(self, layout):
    self.add_section_header(layout, "Sender Details")
    sender_row = QHBoxLayout()
    self.sender_entry = QLineEdit("Nishi")
    self.sender_entry.setToolTip("Sender name for signature")
    sender_label = QLabel("Sender Name (Signature):")
    sender_row.addWidget(sender_label)
    sender_row.addWidget(self.sender_entry)
    self.fixed_sender_checkbox = QCheckBox("Use Fixed Sender Name")
    self.fixed_sender_checkbox.setChecked(True)
    self.sender_entry.setReadOnly(True)
    self.fixed_sender_checkbox.stateChanged.connect(self.toggle_sender_fixed)
    sender_row.addWidget(self.fixed_sender_checkbox)
    layout.addLayout(sender_row)

  def setup_employee_section(self, layout):
      self.add_section_header(layout, "Employee Details")
      self.people_entries = []
      self.people_layout = QVBoxLayout()
      scroll = QScrollArea()
      scroll.setWidgetResizable(True)
      people_widget = QWidget()
      people_widget.setLayout(self.people_layout)
      scroll.setWidget(people_widget)
      scroll.setMinimumHeight(120)
      layout.addWidget(scroll)
      self.add_person_fields()
      add_person_btn = QPushButton("Add Another Person")
      add_person_btn.setToolTip("Add fields for another employee")
      add_person_btn.clicked.connect(self.add_person_fields)
      layout.addWidget(add_person_btn)

  def setup_action_buttons(self, layout):
      button_row = QHBoxLayout()
      button_row.addStretch(1)

      preview_button = QPushButton("Preview and Send")
      preview_button.setToolTip("Preview email in Outlook before sending")
      preview_button.clicked.connect(self.preview_email)

      send_button = QPushButton("Send")
      send_button.setToolTip("Send email directly via Outlook")
      send_button.clicked.connect(self.send_email)

      button_row.addWidget(preview_button)
      button_row.addWidget(send_button)
      button_row.addStretch(1)

      layout.addLayout(button_row)

  def add_section_header(self, layout, text):
      label = QLabel(text)
      label.setFont(QFont("Arial", 12, QFont.Bold))
      label.setStyleSheet("margin-top:12px; margin-bottom:4px;")
      line = QFrame()
      line.setFrameShape(QFrame.HLine)
      line.setFrameShadow(QFrame.Sunken)
      layout.addWidget(label)
      layout.addWidget(line)

  def add_row(self, parent_layout, label_text, default="", readonly=False):
    row = QHBoxLayout()
    label = QLabel(label_text)
    entry = QLineEdit(default)
    entry.setReadOnly(readonly)
    row.addWidget(label)
    row.addWidget(entry)
    parent_layout.addLayout(row)
    return entry

  def create_line_edit(self, placeholder):
    entry = QLineEdit()
    entry.setPlaceholderText(placeholder)
    return entry

  def toggle_sender_fixed(self, state):
    self.sender_entry.setReadOnly(state == 2)

  def handle_to_dropdown(self, selection):
    if selection == "Custom":
        self.custom_to_entry.clear()
        self.custom_to_entry.setPlaceholderText("Enter recipient emails, semicolon-separated")
    else:
        emails = self.RECIPIENT_GROUPS.get(selection, [])
        self.custom_to_entry.setText("; ".join(emails))

  def show_to_instruction(self, _):
    if not self.to_popup_shown:
        QMessageBox.information(
            self,
            "Input Format",
            "If adding any recipient, use a semicolon (;) to separate."
        )
        self.to_popup_shown = True

  def show_cc_instruction(self, _):
     if not self.cc_popup_shown:
        QMessageBox.information(
            self,
            "Input Format",
            "If adding any recipient, use a semicolon (;) to separate."
        )
        self.cc_popup_shown = True

  def add_person_fields(self):
    row = QHBoxLayout()
    row.setSpacing(5)
    row.setContentsMargins(0, 0, 0, 0)

    sr_no_entry = self.create_line_edit("Sr. No.")
    name_entry = self.create_line_edit("Employee Name")
    code_entry = self.create_line_edit("Employee Code")

    role_dropdown, role_entry = self.create_dropdown_with_custom_entry("Role", self.ROLES)
    company_dropdown, company_entry = self.create_dropdown_with_custom_entry("Company", self.COMPANIES)

    delete_btn = QPushButton("Delete", objectName="deleteButton")                                                                                                                                                   # pyright: ignore[reportCallIssue]
    delete_btn.setToolTip("Remove this person")
    row.addWidget(sr_no_entry)
    row.addWidget(QLabel("Name:"))
    row.addWidget(name_entry)
    row.addWidget(QLabel("Code:"))
    row.addWidget(code_entry)
    row.addWidget(role_dropdown)
    row.addWidget(role_entry)
    row.addWidget(company_dropdown)
    row.addWidget(company_entry)
    row.addWidget(delete_btn)

    frame = QFrame()
    frame.setLayout(row)
    self.people_layout.addWidget(frame)

    person_tuple = (sr_no_entry, name_entry, code_entry, role_dropdown, company_dropdown, role_entry, company_entry, frame, delete_btn)
    self.people_entries.append(person_tuple)

    delete_btn.clicked.connect(lambda: self.delete_person(person_tuple))
    self.update_delete_buttons()

    new_height = 500 + 80 * len(self.people_entries)
    self.resize(self.width(), new_height)

  def create_dropdown_with_custom_entry(self, placeholder_text, items):
    dropdown = QComboBox()
    dropdown.addItem(placeholder_text)
    dropdown.addItems(items)
    dropdown.setCurrentIndex(0)
    
    entry = QLineEdit()
    entry.setPlaceholderText(placeholder_text)
    entry.setVisible(False)

    dropdown.currentTextChanged.connect(lambda: self.handle_custom_dropdown(dropdown, entry))
    return dropdown, entry

  def delete_person(self, person_tuple):
    if len(self.people_entries) <= 1:
      return
    _, _, _, _, _, _, _, frame, _ = person_tuple
    self.people_entries.remove(person_tuple)
    frame.deleteLater()
    self.update_delete_buttons()
