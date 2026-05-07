from __future__ import annotations

import sys
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional

from PySide6.QtCore import QDate, QFile, Qt
from PySide6.QtUiTools import QUiLoader
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QDateEdit,
    QFileDialog,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QScrollArea,
    QSpinBox,
    QSplitter,
    QStatusBar,
    QTableWidget,
    QTableWidgetItem,
    QWidget,
)

from excel_repo import ExcelRepository
from models import ContractData
from planner import PlanResult
from extractor import ContractExtractor


UI_FILE = Path(__file__).with_name("contract_planner_upgraded.ui")
DEFAULT_STATUSES = ["Planned", "Running", "Due Soon", "Completed", "On Hold"]


@dataclass
class ProjectRecord:
    contract: ContractData
    plan: PlanResult
    status: str = "Planned"
    notes: str = ""
    raw_text: str = ""
    source_file: str = ""
    created_at: datetime = field(default_factory=datetime.now)

    @property
    def contract_no(self) -> str:
        return self.contract.contract_no

    @property
    def customer_name(self) -> str:
        return self.contract.customer_name

    @property
    def address(self) -> str:
        return self.contract.installation_address or self.contract.source_file

    @property
    def start_date(self) -> datetime:
        return self.plan.start_date

    @property
    def product_end(self) -> datetime:
        return self.plan.product_end

    @property
    def installation_end(self) -> datetime:
        return self.plan.installation_end


class ContractPlannerApp(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.extractor = ContractExtractor()
        self.projects: list[ProjectRecord] = []
        self.current_extracted_data: Optional[ContractData] = None
        self.current_extracted_plan: Optional[PlanResult] = None
        self.current_extracted_pdf: str = ""
        self.filtered_project_indices: list[int] = []

        self._load_ui()
        self._cache_widgets()
        self._setup_initial_state()
        self._connect_signals()
        self._configure_layout_behavior()
        self.refresh_all_views()

    # ---------- UI setup ----------
    def _load_ui(self) -> None:
        if not UI_FILE.exists():
            raise FileNotFoundError(f"UI file not found: {UI_FILE}")

        loader = QUiLoader()
        ui_file = QFile(str(UI_FILE))
        if not ui_file.open(QFile.ReadOnly):
            raise RuntimeError(f"Could not open UI file: {UI_FILE}")

        loaded_window = loader.load(ui_file)
        ui_file.close()

        if loaded_window is None:
            raise RuntimeError(f"Failed to load UI file: {UI_FILE}")

        self.ui = loaded_window.centralWidget()
        self.setCentralWidget(self.ui)
        self.setWindowTitle(loaded_window.windowTitle())
        self.resize(loaded_window.size())

    def _cache_widgets(self) -> None:
        # Path widgets
        self.inputWorkbookPath: QLineEdit = self._w("inputWorkbookPath", QLineEdit)
        self.inputPdfPath: QLineEdit = self._w("inputPdfPath", QLineEdit)

        # Form widgets
        self.inputContractNo: QLineEdit = self._w("inputContractNo", QLineEdit)
        self.inputCustomerName: QLineEdit = self._w("inputCustomerName", QLineEdit)
        self.inputCustomerAddress: QPlainTextEdit = self._w("inputCustomerAddress", QPlainTextEdit)
        self.inputInstallationAddress: QPlainTextEdit = self._w("inputInstallationAddress", QPlainTextEdit)
        self.inputProductDays: QSpinBox = self._w("inputProductDays", QSpinBox)
        self.inputInstallationDays: QSpinBox = self._w("inputInstallationDays", QSpinBox)
        self.inputStartDate: QDateEdit = self._w("inputStartDate", QDateEdit)
        self.comboProjectStatus: QComboBox = self._w("comboProjectStatus", QComboBox)
        self.inputSessionNotes: QPlainTextEdit = self._w("inputSessionNotes", QPlainTextEdit)
        self.textProjectNotes: QPlainTextEdit = self._w("textProjectNotes", QPlainTextEdit)
        self.textRawExtractedText: QPlainTextEdit = self._w("textRawExtractedText", QPlainTextEdit)
        self.textExtractionSummary: QPlainTextEdit = self._w("textExtractionSummary", QPlainTextEdit)

        self.inputLoadCapacityKg: QSpinBox = self._w("inputLoadCapacityKg", QSpinBox)
        self.inputSpeedMpm: QSpinBox = self._w("inputSpeedMpm", QSpinBox)
        self.inputMotorBrand: QLineEdit = self._w("inputMotorBrand", QLineEdit)
        self.inputMotorPower: QLineEdit = self._w("inputMotorPower", QLineEdit)
        self.inputControlSystemBrand: QLineEdit = self._w("inputControlSystemBrand", QLineEdit)
        self.inputNumberOfStops: QSpinBox = self._w("inputNumberOfStops", QSpinBox)
        self.inputNumberOfFloorsServed: QSpinBox = self._w("inputNumberOfFloorsServed", QSpinBox)
        self.inputElectricalBoxBrand: QLineEdit = self._w("inputElectricalBoxBrand", QLineEdit)

        self.splitterMain: QSplitter = self._w("splitterMain", QSplitter)
        self.leftPanelFrame: QWidget = self._w("leftPanelFrame", QWidget)
        self.splitterDetails: QSplitter = self._w("splitterDetails", QSplitter)

        # Tables and filters
        self.tableContracts: QTableWidget = self._w("tableContracts", QTableWidget)
        self.tableTimeline: QTableWidget = self._w("tableTimeline", QTableWidget)
        self.tableTasksForDate: QTableWidget = self._w("tableTasksForDate", QTableWidget)
        self.comboStatusFilter: QComboBox = self._w("comboStatusFilter", QComboBox)
        self.inputFilterMonth: QDateEdit = self._w("inputFilterMonth", QDateEdit)
        self.inputSearchContracts: QLineEdit = self._w("inputSearchContracts", QLineEdit)
        self.comboTimelineScale: QComboBox = self._w("comboTimelineScale", QComboBox)
        self.calendarWidget = self._w("calendarWidget", QWidget)
        self.listImportQueue = self._w("listImportQueue", QWidget)

        # Selected project fields
        self.inputSelectedContractNo: QLineEdit = self._w("inputSelectedContractNo", QLineEdit)
        self.inputSelectedCustomer: QLineEdit = self._w("inputSelectedCustomer", QLineEdit)
        self.inputSelectedStatus: QLineEdit = self._w("inputSelectedStatus", QLineEdit)
        self.inputSelectedDateRange: QLineEdit = self._w("inputSelectedDateRange", QLineEdit)

        # Summary labels
        self.labelCardTotalContractsValue = self._w("labelCardTotalContractsValue", QWidget)
        self.labelCardRunningProjectsValue = self._w("labelCardRunningProjectsValue", QWidget)
        self.labelCardDueSoonValue = self._w("labelCardDueSoonValue", QWidget)
        self.labelCardOverdueValue = self._w("labelCardOverdueValue", QWidget)

        # Buttons
        self.btnBrowseWorkbook: QPushButton = self._w("btnBrowseWorkbook", QPushButton)
        self.btnUploadPdf: QPushButton = self._w("btnUploadPdf", QPushButton)
        self.btnExtract: QPushButton = self._w("btnExtract", QPushButton)
        self.btnAddToProjects: QPushButton = self._w("btnAddToProjects", QPushButton)
        self.btnReplaceSelectedProject: QPushButton = self._w("btnReplaceSelectedProject", QPushButton)
        self.btnGeneratePlan: QPushButton = self._w("btnGeneratePlan", QPushButton)
        self.btnSaveExcel: QPushButton = self._w("btnSaveExcel", QPushButton)
        self.btnReloadWorkbook: QPushButton = self._w("btnReloadWorkbook", QPushButton)
        self.btnRemoveSelectedProject: QPushButton = self._w("btnRemoveSelectedProject", QPushButton)
        self.btnClearForm: QPushButton = self._w("btnClearForm", QPushButton)
        self.btnExportSelection: QPushButton = self._w("btnExportSelection", QPushButton)

        self.statusbar = self.findChild(QStatusBar, "statusbar")
        if self.statusbar is None:
            self.statusbar = QStatusBar(self)
            self.setStatusBar(self.statusbar)

    def _setup_initial_state(self) -> None:
        self.comboProjectStatus.clear()
        self.comboProjectStatus.addItems(DEFAULT_STATUSES)

        self.comboStatusFilter.clear()
        self.comboStatusFilter.addItems(["All"] + DEFAULT_STATUSES)

        if self.comboTimelineScale.count() == 0:
            self.comboTimelineScale.addItems(["Days", "Weeks"])

        today = QDate.currentDate()
        self.inputStartDate.setDate(today)
        self.inputFilterMonth.setDate(today)

        self.inputProductDays.setRange(0, 9999)
        self.inputInstallationDays.setRange(0, 9999)

        self.tableContracts.setSelectionBehavior(QTableWidget.SelectRows)
        self.tableContracts.setSelectionMode(QTableWidget.SingleSelection)
        self.tableContracts.setColumnCount(9)
        self.tableContracts.setHorizontalHeaderLabels(
            [
                "Contract No",
                "Customer",
                "Address",
                "Product Days",
                "Installation Days",
                "Start Date",
                "Product End",
                "Installation End",
                "Status",
            ]
        )
        self.tableContracts.setSortingEnabled(False)
        self.tableContracts.verticalHeader().setVisible(False)

        self.tableTimeline.setSelectionBehavior(QTableWidget.SelectRows)
        self.tableTimeline.setSelectionMode(QTableWidget.SingleSelection)
        self.tableTimeline.verticalHeader().setVisible(False)

        self.tableTasksForDate.setSelectionBehavior(QTableWidget.SelectRows)
        self.tableTasksForDate.setSelectionMode(QTableWidget.SingleSelection)
        self.tableTasksForDate.setColumnCount(4)
        self.tableTasksForDate.setHorizontalHeaderLabels(["Contract No", "Customer", "Task", "Date"])
        self.tableTasksForDate.verticalHeader().setVisible(False)

        for field in [
            self.inputSelectedContractNo,
            self.inputSelectedCustomer,
            self.inputSelectedStatus,
            self.inputSelectedDateRange,
        ]:
            field.setReadOnly(True)

        self.textRawExtractedText.setReadOnly(True)
        self.textExtractionSummary.setReadOnly(True)
        self.textRawExtractedText.hide()
        self.inputLoadCapacityKg.setRange(0, 100000)
        self.inputSpeedMpm.setRange(0, 10000)
        self.inputNumberOfStops.setRange(0, 1000)
        self.inputNumberOfFloorsServed.setRange(0, 1000)
        self.show_status("Ready.")

    def _configure_layout_behavior(self) -> None:
        left_scroll = self.findChild(QScrollArea, "leftScrollArea")
        if left_scroll is None:
            left_scroll = QScrollArea(self.splitterMain)
            left_scroll.setObjectName("leftScrollArea")
            left_scroll.setWidgetResizable(True)
            left_scroll.setFrameShape(QScrollArea.NoFrame)
            self.leftPanelFrame.setParent(None)
            left_scroll.setWidget(self.leftPanelFrame)
            self.splitterMain.insertWidget(0, left_scroll)

        self.splitterMain.setStretchFactor(0, 0)
        self.splitterMain.setStretchFactor(1, 1)
        self.splitterMain.setSizes([520, 860])
        self.splitterDetails.setSizes([700, 0])
        self.textRawExtractedText.hide()

    def _connect_signals(self) -> None:
        self.btnBrowseWorkbook.clicked.connect(self.choose_workbook)
        self.btnUploadPdf.clicked.connect(self.choose_pdf)
        self.btnExtract.clicked.connect(self.extract_from_pdf)
        self.btnGeneratePlan.clicked.connect(self.generate_plan_from_form)
        self.btnAddToProjects.clicked.connect(self.add_project)
        self.btnReplaceSelectedProject.clicked.connect(self.replace_selected_project)
        self.btnSaveExcel.clicked.connect(self.save_all_projects_to_excel)
        self.btnReloadWorkbook.clicked.connect(self.refresh_all_views)
        self.btnRemoveSelectedProject.clicked.connect(self.remove_selected_project)
        self.btnClearForm.clicked.connect(self.clear_form)
        self.btnExportSelection.clicked.connect(self.save_selected_project_to_excel)

        self.tableContracts.itemSelectionChanged.connect(self.on_contract_selection_changed)
        self.comboStatusFilter.currentIndexChanged.connect(self.refresh_contracts_table)
        self.inputFilterMonth.dateChanged.connect(self.refresh_contracts_table)
        self.inputSearchContracts.textChanged.connect(self.refresh_contracts_table)
        self.comboTimelineScale.currentIndexChanged.connect(self.refresh_timeline_table)

        # Calendar widget is declared as QWidget in the generated object tree, but it still exposes selectionChanged at runtime.
        if hasattr(self.calendarWidget, "selectionChanged"):
            self.calendarWidget.selectionChanged.connect(self.refresh_tasks_for_selected_date)

    def _w(self, name: str, expected_type):
        widget = self.findChild(expected_type, name)
        if widget is None:
            raise AttributeError(f"Widget '{name}' not found in UI.")
        return widget


    def choose_workbook(self) -> None:
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Choose Excel Workbook",
            self.inputWorkbookPath.text().strip() or str(Path.cwd() / "contracts.xlsx"),
            "Excel Workbook (*.xlsx)",
        )
        if file_path:
            if not file_path.lower().endswith(".xlsx"):
                file_path += ".xlsx"
            self.inputWorkbookPath.setText(file_path)
            self.show_status(f"Workbook selected: {file_path}")

    def choose_pdf(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Choose Contract PDF",
            self.inputPdfPath.text().strip() or str(Path.cwd()),
            "PDF Files (*.pdf *.PDF)",
        )
        if file_path:
            self.inputPdfPath.setText(file_path)
            if hasattr(self.listImportQueue, "addItem"):
                self.listImportQueue.addItem(file_path)
            self.show_status(f"PDF selected: {Path(file_path).name}")

    def extract_from_pdf(self) -> None:
        pdf_path = self.inputPdfPath.text().strip()
        if not pdf_path:
            self.warn("Please choose a PDF file first.")
            return

        try:
            data = self.extractor.extract(pdf_path)
            self.current_extracted_data = data
            self.current_extracted_pdf = pdf_path
            self._populate_form_from_contract(data)
            self.generate_plan_from_form(show_message=False)
            self.textRawExtractedText.setPlainText(data.raw_text)
            self.textExtractionSummary.setPlainText(self._build_summary(data, self.current_extracted_plan, self.comboProjectStatus.currentText()))
            self.show_status(f"Extracted contract: {data.contract_no or Path(pdf_path).name}")
        except Exception as exc:
            self.error(f"Failed to extract contract from PDF.\n\n{exc}")

    def generate_plan_from_form(self, show_message: bool = True) -> None:
        contract = self._build_contract_from_form()
        errors = self.validate_contract(contract)
        if errors:
            if show_message:
                self.warn("Cannot generate plan:\n- " + "\n- ".join(errors))
            return

        start_dt = self.qdate_to_datetime(self.inputStartDate.date())
        plan = self.build_plan(contract.product_days, contract.installation_days, start_dt)
        self.current_extracted_data = contract
        self.current_extracted_plan = plan
        self._update_preview_date_fields(contract, plan)

        if show_message:
            self.show_status("Plan generated from current form values.")

    def add_project(self) -> None:
        project = self._build_project_from_form()
        if project is None:
            return

        existing_index = self.find_project_index(project.contract_no)
        if existing_index is not None:
            self.warn(
                "A project with the same contract number already exists.\n"
                "Use 'Replace Selected Project' if you want to update it."
            )
            return

        self.projects.append(project)
        self.refresh_all_views()
        self.select_project_by_index(len(self.projects) - 1)
        self.show_status(f"Project added: {project.contract_no}")

    def replace_selected_project(self) -> None:
        selected_index = self.get_selected_project_index()
        if selected_index is None:
            self.warn("Please select a project row to replace.")
            return

        project = self._build_project_from_form()
        if project is None:
            return

        duplicate_index = self.find_project_index(project.contract_no)
        if duplicate_index is not None and duplicate_index != selected_index:
            self.warn("Another project already uses this contract number.")
            return

        self.projects[selected_index] = project
        self.refresh_all_views()
        self.select_project_by_index(selected_index)
        self.show_status(f"Project updated: {project.contract_no}")

    def remove_selected_project(self) -> None:
        selected_index = self.get_selected_project_index()
        if selected_index is None:
            self.warn("Please select a project to remove.")
            return

        removed = self.projects.pop(selected_index)
        self.refresh_all_views()
        self.show_status(f"Removed project: {removed.contract_no}")

    def clear_form(self) -> None:
        self.current_extracted_data = None
        self.current_extracted_plan = None
        self.current_extracted_pdf = ""

        self.inputPdfPath.clear()
        self.inputContractNo.clear()
        self.inputCustomerName.clear()
        self.inputCustomerAddress.clear()
        self.inputInstallationAddress.clear()
        self.inputProductDays.setValue(0)
        self.inputInstallationDays.setValue(0)
        self.inputStartDate.setDate(QDate.currentDate())
        self.comboProjectStatus.setCurrentText("Planned")
        self.inputLoadCapacityKg.setValue(0)
        self.inputSpeedMpm.setValue(0)
        self.inputMotorBrand.clear()
        self.inputMotorPower.clear()
        self.inputControlSystemBrand.clear()
        self.inputNumberOfStops.setValue(0)
        self.inputNumberOfFloorsServed.setValue(0)
        self.inputElectricalBoxBrand.clear()
        self.textProjectNotes.clear()
        self.textExtractionSummary.clear()
        self.textRawExtractedText.clear()
        self.show_status("Form cleared.")

    def save_selected_project_to_excel(self) -> None:
        project = self.get_selected_project()
        if project is None:
            self.warn("Please select a project to export.")
            return

        workbook_path = self.inputWorkbookPath.text().strip()
        if not workbook_path:
            self.warn("Please choose an Excel workbook path first.")
            return

        try:
            repo = ExcelRepository(workbook_path)
            repo.save_contract_and_plan(project.contract, project.plan)
            self.show_status(f"Exported selected project to Excel: {project.contract_no}")
        except Exception as exc:
            self.error(f"Failed to export selected project.\n\n{exc}")

    def save_all_projects_to_excel(self) -> None:
        workbook_path = self.inputWorkbookPath.text().strip()
        if not workbook_path:
            self.warn("Please choose an Excel workbook path first.")
            return

        if not self.projects:
            self.warn("There are no projects to save.")
            return

        try:
            repo = ExcelRepository(workbook_path)
            for project in self.projects:
                repo.save_contract_and_plan(project.contract, project.plan)
            self.show_status(f"Saved {len(self.projects)} project(s) to Excel.")
        except Exception as exc:
            self.error(f"Failed to save projects to Excel.\n\n{exc}")

    #Data helpers 
    def _populate_form_from_contract(self, data: ContractData) -> None:
        self.inputContractNo.setText(data.contract_no)
        self.inputCustomerName.setText(data.customer_name)
        self.inputCustomerAddress.setPlainText(data.customer_address)
        self.inputInstallationAddress.setPlainText(data.installation_address)
        self.inputProductDays.setValue(data.product_days)
        self.inputInstallationDays.setValue(data.installation_days)
        self.inputLoadCapacityKg.setValue(data.load_capacity_kg)
        self.inputSpeedMpm.setValue(data.speed_mpm)
        self.inputMotorBrand.setText(data.motor_brand)
        self.inputMotorPower.setText(data.motor_power)
        self.inputControlSystemBrand.setText(data.control_system_brand)
        self.inputNumberOfStops.setValue(data.number_of_stops)
        self.inputNumberOfFloorsServed.setValue(data.number_of_floors_served)
        self.inputElectricalBoxBrand.setText(data.electrical_box_brand)

    def _build_contract_from_form(self) -> ContractData:
        source_file = self.current_extracted_pdf or self.inputPdfPath.text().strip()
        raw_text = self.textRawExtractedText.toPlainText().strip()
        return ContractData(
            source_file=source_file,
            contract_no=self.inputContractNo.text().strip(),
            customer_name=self.inputCustomerName.text().strip(),
            customer_address=self.inputCustomerAddress.toPlainText().strip(),
            installation_address=self.inputInstallationAddress.toPlainText().strip(),
            product_days=int(self.inputProductDays.value()),
            installation_days=int(self.inputInstallationDays.value()),
            load_capacity_kg=int(self.inputLoadCapacityKg.value()),
            speed_mpm=int(self.inputSpeedMpm.value()),
            motor_brand=self.inputMotorBrand.text().strip(),
            motor_power=self.inputMotorPower.text().strip(),
            control_system_brand=self.inputControlSystemBrand.text().strip(),
            number_of_stops=int(self.inputNumberOfStops.value()),
            number_of_floors_served=int(self.inputNumberOfFloorsServed.value()),
            electrical_box_brand=self.inputElectricalBoxBrand.text().strip(),
            raw_text=raw_text,
        )

    def _build_project_from_form(self) -> Optional[ProjectRecord]:
        contract = self._build_contract_from_form()
        errors = self.validate_contract(contract)
        if errors:
            self.warn("Cannot add project:\n- " + "\n- ".join(errors))
            return None

        start_dt = self.qdate_to_datetime(self.inputStartDate.date())
        plan = self.build_plan(contract.product_days, contract.installation_days, start_dt)

        return ProjectRecord(
            contract=contract,
            plan=plan,
            status=self.comboProjectStatus.currentText().strip() or "Planned",
            notes=self.textProjectNotes.toPlainText().strip(),
            raw_text=contract.raw_text,
            source_file=contract.source_file,
        )

    def validate_contract(self, contract: ContractData) -> list[str]:
        errors: list[str] = []
        if not contract.contract_no:
            errors.append("Missing contract number")
        if not contract.customer_name:
            errors.append("Missing customer name")
        if contract.product_days <= 0:
            errors.append("Product days must be greater than 0")
        if contract.installation_days <= 0:
            errors.append("Installation days must be greater than 0")
        return errors

    @staticmethod
    def build_plan(product_days: int, installation_days: int, start_date: datetime) -> PlanResult:
        product_start = start_date
        product_end = product_start + timedelta(days=product_days)
        installation_start = product_end
        installation_end = installation_start + timedelta(days=installation_days)
        return PlanResult(
            start_date=start_date,
            product_start=product_start,
            product_end=product_end,
            installation_start=installation_start,
            installation_end=installation_end,
        )

    @staticmethod
    def qdate_to_datetime(date_value: QDate) -> datetime:
        return datetime(date_value.year(), date_value.month(), date_value.day())

    def find_project_index(self, contract_no: str) -> Optional[int]:
        contract_no = (contract_no or "").strip().lower()
        for idx, project in enumerate(self.projects):
            if project.contract_no.strip().lower() == contract_no:
                return idx
        return None

    # Table refresh
    def refresh_all_views(self) -> None:
        self.refresh_summary_cards()
        self.refresh_contracts_table()
        self.refresh_timeline_table()
        self.refresh_tasks_for_selected_date()
        self.refresh_details_panel(None)

    def refresh_summary_cards(self) -> None:
        total = len(self.projects)
        running = sum(1 for p in self.projects if p.status in {"Running", "Planned"})
        due_soon = sum(1 for p in self.projects if 0 <= (p.installation_end.date() - datetime.today().date()).days <= 7)
        overdue = sum(1 for p in self.projects if p.installation_end.date() < datetime.today().date() and p.status != "Completed")

        self.labelCardTotalContractsValue.setText(str(total))
        self.labelCardRunningProjectsValue.setText(str(running))
        self.labelCardDueSoonValue.setText(str(due_soon))
        self.labelCardOverdueValue.setText(str(overdue))

    def get_filtered_projects(self) -> list[tuple[int, ProjectRecord]]:
        status_filter = self.comboStatusFilter.currentText().strip()
        search_text = self.inputSearchContracts.text().strip().lower()
        month_date = self.inputFilterMonth.date()
        month_year = (month_date.year(), month_date.month())

        filtered: list[tuple[int, ProjectRecord]] = []
        for idx, project in enumerate(self.projects):
            if status_filter != "All" and project.status != status_filter:
                continue

            if (project.start_date.year, project.start_date.month) != month_year and (
                project.product_end.year,
                project.product_end.month,
            ) != month_year and (
                project.installation_end.year,
                project.installation_end.month,
            ) != month_year:
                continue

            haystack = " | ".join(
                [
                    project.contract_no,
                    project.customer_name,
                    project.contract.installation_address,
                    project.source_file,
                ]
            ).lower()
            if search_text and search_text not in haystack:
                continue

            filtered.append((idx, project))
        return filtered

    def refresh_contracts_table(self) -> None:
        filtered = self.get_filtered_projects()
        self.filtered_project_indices = [idx for idx, _ in filtered]

        self.tableContracts.setRowCount(len(filtered))
        for row, (_, project) in enumerate(filtered):
            values = [
                project.contract_no,
                project.customer_name,
                project.contract.installation_address,
                str(project.contract.product_days),
                str(project.contract.installation_days),
                project.start_date.strftime("%Y-%m-%d"),
                project.product_end.strftime("%Y-%m-%d"),
                project.installation_end.strftime("%Y-%m-%d"),
                project.status,
            ]
            for col, value in enumerate(values):
                item = QTableWidgetItem(value)
                if col in {3, 4}:
                    item.setTextAlignment(Qt.AlignCenter)
                self.tableContracts.setItem(row, col, item)

        self.tableContracts.resizeColumnsToContents()
        self.refresh_timeline_table()

    def refresh_timeline_table(self) -> None:
        filtered = self.get_filtered_projects()
        if not filtered:
            self.tableTimeline.setRowCount(0)
            self.tableTimeline.setColumnCount(0)
            return

        all_dates = []
        for _, project in filtered:
            all_dates.extend([project.start_date.date(), project.product_end.date(), project.installation_end.date()])
        start = min(all_dates)
        end = max(all_dates)
        total_days = (end - start).days + 1

        self.tableTimeline.setRowCount(len(filtered))
        self.tableTimeline.setColumnCount(total_days + 1)

        headers = ["Contract"] + [(start + timedelta(days=i)).strftime("%m-%d") for i in range(total_days)]
        self.tableTimeline.setHorizontalHeaderLabels(headers)

        for row, (_, project) in enumerate(filtered):
            self.tableTimeline.setItem(row, 0, QTableWidgetItem(project.contract_no))
            for day_offset in range(total_days):
                current = start + timedelta(days=day_offset)
                text = ""
                if project.plan.product_start.date() <= current <= project.plan.product_end.date():
                    text = "P"
                if project.plan.installation_start.date() <= current <= project.plan.installation_end.date():
                    text = "I"
                self.tableTimeline.setItem(row, day_offset + 1, QTableWidgetItem(text))

        self.tableTimeline.resizeColumnsToContents()

    def refresh_tasks_for_selected_date(self) -> None:
        if hasattr(self.calendarWidget, "selectedDate"):
            selected_qdate = self.calendarWidget.selectedDate()
        else:
            selected_qdate = QDate.currentDate()
        selected_date = self.qdate_to_datetime(selected_qdate).date()

        rows: list[tuple[str, str, str, str]] = []
        for project in self.projects:
            if project.plan.product_start.date() == selected_date:
                rows.append((project.contract_no, project.customer_name, "Product Start", selected_date.isoformat()))
            if project.plan.product_end.date() == selected_date:
                rows.append((project.contract_no, project.customer_name, "Product End", selected_date.isoformat()))
            if project.plan.installation_start.date() == selected_date:
                rows.append((project.contract_no, project.customer_name, "Installation Start", selected_date.isoformat()))
            if project.plan.installation_end.date() == selected_date:
                rows.append((project.contract_no, project.customer_name, "Installation End", selected_date.isoformat()))

        self.tableTasksForDate.setRowCount(len(rows))
        for row, values in enumerate(rows):
            for col, value in enumerate(values):
                self.tableTasksForDate.setItem(row, col, QTableWidgetItem(value))
        self.tableTasksForDate.resizeColumnsToContents()

    #  Selection and details 
    def on_contract_selection_changed(self) -> None:
        project = self.get_selected_project()
        self.refresh_details_panel(project)

    def refresh_details_panel(self, project: Optional[ProjectRecord]) -> None:
        if project is None:
            self.inputSelectedContractNo.clear()
            self.inputSelectedCustomer.clear()
            self.inputSelectedStatus.clear()
            self.inputSelectedDateRange.clear()
            if not self.tableContracts.selectedItems():
                self.textProjectNotes.clear()
                if not self.current_extracted_data:
                    self.textExtractionSummary.clear()
                    self.textRawExtractedText.clear()
            return

        self.inputSelectedContractNo.setText(project.contract_no)
        self.inputSelectedCustomer.setText(project.customer_name)
        self.inputSelectedStatus.setText(project.status)
        self.inputSelectedDateRange.setText(
            f"{project.start_date:%Y-%m-%d} → {project.installation_end:%Y-%m-%d}"
        )
        self.textProjectNotes.setPlainText(project.notes)
        self.textExtractionSummary.setPlainText(self._build_summary(project.contract, project.plan, project.status))
        self.textRawExtractedText.setPlainText(project.raw_text or project.contract.raw_text)

    def get_selected_project_index(self) -> Optional[int]:
        row = self.tableContracts.currentRow()
        if row < 0 or row >= len(self.filtered_project_indices):
            return None
        return self.filtered_project_indices[row]

    def get_selected_project(self) -> Optional[ProjectRecord]:
        index = self.get_selected_project_index()
        if index is None:
            return None
        if 0 <= index < len(self.projects):
            return self.projects[index]
        return None

    def select_project_by_index(self, index: int) -> None:
        if index not in self.filtered_project_indices:
            self.refresh_contracts_table()
        try:
            row = self.filtered_project_indices.index(index)
        except ValueError:
            return
        self.tableContracts.selectRow(row)
        self.tableContracts.setCurrentCell(row, 0)

    def _update_preview_date_fields(self, contract: ContractData, plan: PlanResult) -> None:
        status = self.comboProjectStatus.currentText()
        self.inputSelectedContractNo.setText(contract.contract_no)
        self.inputSelectedCustomer.setText(contract.customer_name)
        self.inputSelectedStatus.setText(status)
        self.inputSelectedDateRange.setText(f"{plan.start_date:%Y-%m-%d} → {plan.installation_end:%Y-%m-%d}")
        self.textExtractionSummary.setPlainText(self._build_summary(contract, plan, status))

    def _build_summary(self, contract: ContractData, plan: Optional[PlanResult], status: str = "") -> str:
        lines: list[str] = []

        def add(label: str, value: str) -> None:
            value = (value or "").strip()
            if value:
                lines.append(f"{label}: {value}")

        add("Contract", contract.contract_no)
        add("Customer", contract.customer_name)
        add("Customer Address", contract.customer_address)
        add("Install Address", contract.installation_address)

        timeline_parts = []
        if contract.product_days:
            timeline_parts.append(f"Production {contract.product_days} days")
        if contract.installation_days:
            timeline_parts.append(f"Installation {contract.installation_days} days")
        if timeline_parts:
            lines.append("Timeline: " + " | ".join(timeline_parts))

        spec_parts = []
        if contract.load_capacity_kg:
            spec_parts.append(f"Load {contract.load_capacity_kg} kg")
        if contract.speed_mpm:
            spec_parts.append(f"Speed {contract.speed_mpm} m/min")
        if contract.motor_power:
            spec_parts.append(f"Power {contract.motor_power}")
        if spec_parts:
            lines.append("Core Specs: " + " | ".join(spec_parts))

        brand_parts = []
        if contract.motor_brand:
            brand_parts.append(f"Motor {contract.motor_brand}")
        if contract.control_system_brand:
            brand_parts.append(f"Control {contract.control_system_brand}")
        if contract.electrical_box_brand:
            brand_parts.append(f"Electrical Box {contract.electrical_box_brand}")
        if brand_parts:
            lines.append("Brands: " + " | ".join(brand_parts))

        if contract.number_of_stops or contract.number_of_floors_served:
            lines.append(
                f"Stops / Floors: {contract.number_of_stops or '-'} / {contract.number_of_floors_served or '-'}"
            )

        if plan is not None:
            lines.append(
                f"Plan Window: {plan.start_date:%Y-%m-%d} → {plan.installation_end:%Y-%m-%d}"
            )

        add("Status", status)
        return "\n".join(lines)

    #  Message helpers 
    def show_status(self, message: str) -> None:
        if self.statusbar:
            self.statusbar.showMessage(message, 5000)
        else:
            print(message)

    def warn(self, message: str) -> None:
        QMessageBox.warning(self, "Warning", message)

    def error(self, message: str) -> None:
        QMessageBox.critical(self, "Error", message)


def main() -> None:
    app = QApplication(sys.argv)
    window = ContractPlannerApp()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
