import sys
import threading
from pathlib import Path

from PySide6.QtCore import Qt, QObject, Signal
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QProgressBar,
    QPlainTextEdit,
    QSizePolicy,
    QSlider,
    QTabWidget,
    QVBoxLayout,
    QWidget,
    QFileDialog,
)


def _ensure_module_paths() -> None:
    base = Path(__file__).resolve().parent
    root = base.parent
    for name in ("ExcelCleaner", "ExcelImageOptimization", "ExcelByteReduce"):
        p = root / name
        if p.is_dir():
            sys.path.insert(0, str(p))


_ensure_module_paths()

from excel_suite_pipeline import run_pipeline_core, open_in_explorer_select
from settings import get_settings, save_settings


class PipelineWorker(QObject):
    log = Signal(str)
    status = Signal(str, float)
    finished = Signal(str)
    failed = Signal(str)

    def __init__(
        self,
        path: Path,
        use_clean: bool,
        use_image: bool,
        use_precision: bool,
        aggressive: bool,
        do_xml_cleanup: bool,
        force_custom: bool,
        parent: QObject | None = None,
    ) -> None:
        super().__init__(parent)
        self.path = path
        self.use_clean = use_clean
        self.use_image = use_image
        self.use_precision = use_precision
        self.aggressive = aggressive
        self.do_xml_cleanup = do_xml_cleanup
        self.force_custom = force_custom

    def run(self) -> None:
        """Run the shared pipeline core in a worker thread.

        이 메서드는 excel_suite_pipeline.run_pipeline_core 를 호출해서
        tkinter 버전과 완전히 동일한 파이프라인 로직을 재사용합니다.
        """

        last_progress = 0.0

        def log_cb(msg: str) -> None:
            self.log.emit(msg)

        def set_status_cb(text: str, progress: float | None) -> None:
            nonlocal last_progress
            if progress is not None:
                last_progress = progress
            self.status.emit(text, last_progress)

        def show_error_cb(title: str, text: str) -> None:
            # 실제 메시지 박스는 메인 스레드에서 처리하도록 failed 시그널로 위임
            self.failed.emit(text)

        def on_finished_cb(final_path: Path) -> None:
            self.finished.emit(str(final_path))

        try:
            run_pipeline_core(
                start_path=self.path,
                use_clean=self.use_clean,
                use_image=self.use_image,
                use_precision=self.use_precision,
                aggressive=self.aggressive,
                do_xml_cleanup=self.do_xml_cleanup,
                force_custom=self.force_custom,
                log=log_cb,
                set_status=set_status_cb,
                show_error=show_error_cb,
                on_finished=on_finished_cb,
            )
        except Exception as e:  # noqa: BLE001
            self.failed.emit(f"예기치 못한 오류: {e}")


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("ExcelSlimmer")
        self.resize(1120, 720)

        self._settings = get_settings()
        self._theme = self._settings.theme

        self._worker_thread: threading.Thread | None = None
        self._worker: PipelineWorker | None = None

        self._build_ui()

    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)

        root_layout = QVBoxLayout(central)
        root_layout.setContentsMargins(18, 14, 18, 18)
        root_layout.setSpacing(12)

        header_layout = QHBoxLayout()
        header_layout.setSpacing(8)
        root_layout.addLayout(header_layout)

        title = QLabel("ExcelSlimmer")
        title.setStyleSheet("font-size: 20px; font-weight: 700;")
        header_layout.addWidget(title, 0, Qt.AlignLeft | Qt.AlignVCenter)
        header_layout.addStretch(1)

        tabs = QTabWidget()
        root_layout.addWidget(tabs, 1)

        self.pipeline_tab = QWidget()
        self.settings_tab = QWidget()
        tabs.addTab(self.pipeline_tab, "슬리머 실행")
        tabs.addTab(self.settings_tab, "환경 설정")

        pipe_layout = QGridLayout(self.pipeline_tab)
        # 좌우는 12px, 상단은 약간 내려서 슬리머 실행 탭 상단과 라벨 사이 간격을 확보
        pipe_layout.setContentsMargins(12, 8, 12, 0)
        pipe_layout.setHorizontalSpacing(12)

        left_col = QWidget()
        left_layout = QVBoxLayout(left_col)
        left_layout.setContentsMargins(0, 0, 0, 0)
        # 라벨과 카드 사이 간격을 줄이기 위해 spacing을 약간 낮게 설정
        left_layout.setSpacing(6)

        right_col = QWidget()
        right_layout = QVBoxLayout(right_col)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(6)

        pipe_layout.addWidget(left_col, 0, 0)
        pipe_layout.addWidget(right_col, 0, 1)
        pipe_layout.setColumnStretch(0, 0)
        pipe_layout.setColumnStretch(1, 1)

        # 대상 파일 카드
        file_label = QLabel("대상 파일")
        file_label.setStyleSheet("font-weight: 600;")
        left_layout.addWidget(file_label)

        file_group = QGroupBox()
        file_group.setStyleSheet(self._card_style())
        fg_layout = QVBoxLayout(file_group)
        fg_layout.setSpacing(6)

        fg_layout.addWidget(QLabel("파일 경로:"))

        self.file_edit = QLineEdit()
        self.file_edit.setObjectName("file_path_edit")
        self.file_edit.setReadOnly(True)
        self.file_edit.setFrame(True)
        fg_layout.addWidget(self.file_edit)

        browse_btn = QPushButton("찾기...")
        browse_btn.setCursor(Qt.PointingHandCursor)
        browse_btn.clicked.connect(self._on_browse)
        fg_layout.addWidget(browse_btn, 0, Qt.AlignRight)

        left_layout.addWidget(file_group)

        # 실행할 기능 카드
        func_label = QLabel("실행할 기능")
        func_label.setStyleSheet("font-weight: 600;")
        left_layout.addWidget(func_label)

        func_group = QGroupBox()
        func_group.setStyleSheet(self._card_style())
        func_layout = QVBoxLayout(func_group)
        func_layout.setSpacing(4)

        settings = get_settings()
        self.clean_check = QCheckBox("이름 정의 정리 (definedNames 클린)")
        self.clean_check.setFocusPolicy(Qt.NoFocus)
        self.clean_check.setCursor(Qt.PointingHandCursor)
        self.image_check = QCheckBox("이미지 최적화 (이미지 리사이즈/압축)")
        self.image_check.setFocusPolicy(Qt.NoFocus)
        self.image_check.setCursor(Qt.PointingHandCursor)
        self.precision_check = QCheckBox("정밀 슬리머 (Precision Plus)")
        self.precision_check.setFocusPolicy(Qt.NoFocus)
        self.precision_check.setCursor(Qt.PointingHandCursor)

        func_layout.addWidget(self.clean_check)
            
        # 이미지 최적화 체크박스 및 해상도/품질 설정 (슬라이더 + 직접 입력)
        func_layout.addWidget(self.image_check)

        self.max_edge_label = QLabel(f"최대 해상도 (px): {settings.image_max_edge}")
        self.max_edge_slider = QSlider(Qt.Horizontal)
        self.max_edge_slider.setCursor(Qt.PointingHandCursor)
        self.max_edge_slider.setRange(1400, 4000)
        self.max_edge_slider.setSingleStep(100)
        self.max_edge_slider.setValue(settings.image_max_edge)

        self.max_edge_edit = QLineEdit(str(settings.image_max_edge))
        self.max_edge_edit.setObjectName("max_edge_edit")
        self.max_edge_edit.setFixedWidth(72)
        self.max_edge_edit.setMaxLength(4)
        self.max_edge_edit.setFrame(True)
        self.max_edge_edit.setCursor(Qt.PointingHandCursor)
        self.max_edge_edit.editingFinished.connect(self._on_max_edge_edit_finished)

        max_edge_row = QHBoxLayout()
        max_edge_row.setSpacing(6)
        max_edge_row.addWidget(self.max_edge_slider, 1)
        max_edge_row.addWidget(self.max_edge_edit, 0)

        self.quality_label = QLabel(f"JPEG 품질: {settings.image_quality}%")
        self.quality_slider = QSlider(Qt.Horizontal)
        self.quality_slider.setCursor(Qt.PointingHandCursor)
        self.quality_slider.setRange(70, 100)
        self.quality_slider.setSingleStep(5)
        self.quality_slider.setValue(settings.image_quality)

        self.quality_edit = QLineEdit(str(settings.image_quality))
        self.quality_edit.setObjectName("quality_edit")
        self.quality_edit.setFixedWidth(72)
        self.quality_edit.setMaxLength(3)
        self.quality_edit.setFrame(True)
        self.quality_edit.setCursor(Qt.PointingHandCursor)
        self.quality_edit.editingFinished.connect(self._on_quality_edit_finished)

        quality_row = QHBoxLayout()
        quality_row.setSpacing(6)
        quality_row.addWidget(self.quality_slider, 1)
        quality_row.addWidget(self.quality_edit, 0)

        self.max_edge_slider.valueChanged.connect(self._on_image_settings_changed)
        self.quality_slider.valueChanged.connect(self._on_image_settings_changed)

        func_layout.addWidget(self.max_edge_label)
        func_layout.addLayout(max_edge_row)

        self.max_edge_hint_label = QLabel("권장: 1600 px")
        self.max_edge_hint_label.setStyleSheet("color: #666666; font-size: 9pt;")
        func_layout.addWidget(self.max_edge_hint_label)

        func_layout.addWidget(self.quality_label)
        func_layout.addLayout(quality_row)

        self.quality_hint_label = QLabel("권장: 80 %")
        self.quality_hint_label.setStyleSheet("color: #666666; font-size: 9pt;")
        func_layout.addWidget(self.quality_hint_label)

        func_layout.addWidget(self.precision_check)

        warn = QLabel("주의: 정밀 슬리머 사용 시 엑셀에서 복구 여부를 물어볼 수 있습니다.")
        warn.setStyleSheet("color: #ff6666; font-size: 9pt;")
        func_layout.addWidget(warn)

        left_layout.addWidget(func_group)

        # 정밀 슬리머 옵션 카드
        opt_label = QLabel("정밀 슬리머 옵션")
        opt_label.setStyleSheet("font-weight: 600;")
        left_layout.addWidget(opt_label)

        opt_group = QGroupBox()
        opt_group.setStyleSheet(self._card_style())
        opt_layout = QVBoxLayout(opt_group)
        opt_layout.setSpacing(4)

        self.xmlcleanup_check = QCheckBox("XML 정리 (calcChain, printerSettings 등)")
        self.xmlcleanup_check.setFocusPolicy(Qt.NoFocus)
        self.xmlcleanup_check.setCursor(Qt.PointingHandCursor)
        self.force_custom_check = QCheckBox("숨은 XML 데이터 삭제 (customXml, 주의)")
        self.force_custom_check.setFocusPolicy(Qt.NoFocus)
        self.force_custom_check.setCursor(Qt.PointingHandCursor)
        self.aggressive_check = QCheckBox("이미지 포맷 변경 (PNG→JPG) + 참조 동기화 (고급)")
        self.aggressive_check.setFocusPolicy(Qt.NoFocus)
        self.aggressive_check.setCursor(Qt.PointingHandCursor)

        opt_layout.addWidget(self.xmlcleanup_check)
        opt_layout.addWidget(self.force_custom_check)
        opt_layout.addWidget(self.aggressive_check)

        opt_warn = QLabel("주의: 숨은 XML 데이터 삭제는 일반적인 경우 사용하지 마세요.")
        opt_warn.setStyleSheet("color: #ff6666; font-size: 9pt;")
        opt_layout.addWidget(opt_warn)

        opt_warn2 = QLabel("이미지 포맷 변경 옵션은 일부 도형/특수 이미지에서 예기치 않은 영향이 있을 수 있습니다.")
        opt_warn2.setStyleSheet("color: #ff6666; font-size: 9pt;")
        opt_layout.addWidget(opt_warn2)

        left_layout.addWidget(opt_group)

        # 실행/상태 카드
        run_label = QLabel("진행 상태")
        run_label.setStyleSheet("font-weight: 600;")
        left_layout.addWidget(run_label)

        run_group = QGroupBox()
        run_group.setStyleSheet(self._card_style())
        run_layout = QVBoxLayout(run_group)
        run_layout.setSpacing(8)

        self.run_button = QPushButton("슬리머 실행")
        self.run_button.setCursor(Qt.PointingHandCursor)
        self.run_button.clicked.connect(self._on_run_clicked)
        run_layout.addWidget(self.run_button, 0, Qt.AlignLeft)

        status_row = QHBoxLayout()
        status_row.setSpacing(8)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        status_row.addWidget(self.progress_bar, 1)

        self.status_label = QLabel("준비됨")
        status_row.addWidget(self.status_label)

        run_layout.addLayout(status_row)

        left_layout.addWidget(run_group)
        left_layout.addStretch(1)

        # 로그 카드
        log_label = QLabel("로그")
        log_label.setStyleSheet("font-weight: 600;")
        right_layout.addWidget(log_label)

        log_group = QGroupBox()
        log_group.setStyleSheet(self._card_style())
        log_layout = QVBoxLayout(log_group)
        log_layout.setContentsMargins(8, 6, 8, 8)
        self.log_edit = QPlainTextEdit()
        self.log_edit.setReadOnly(True)
        self.log_edit.setLineWrapMode(QPlainTextEdit.NoWrap)
        log_layout.addWidget(self.log_edit)

        right_layout.addWidget(log_group)

        # 환경 설정 탭
        s_layout = QVBoxLayout(self.settings_tab)
        s_layout.setContentsMargins(12, 12, 12, 12)
        s_layout.setSpacing(12)

        settings = get_settings()

        # 너무 가로로 넓어지지 않도록, 가운데에 최대 폭이 제한된 컨테이너를 둔다.
        settings_container = QWidget()
        settings_layout = QVBoxLayout(settings_container)
        settings_layout.setContentsMargins(0, 0, 0, 0)
        settings_layout.setSpacing(10)
        settings_container.setMaximumWidth(520)
        settings_container.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Preferred)

        # 경로/백업/출력 폴더 설정 카드
        path_label = QLabel("경로 설정")
        path_label.setStyleSheet("font-weight: 600;")

        path_group = QGroupBox()
        path_group.setStyleSheet(self._card_style())
        path_layout = QVBoxLayout(path_group)
        path_layout.setSpacing(6)

        self.keep_backup_check = QCheckBox("완성본 저장 폴더에 백업 파일 저장")
        self.keep_backup_check.setFocusPolicy(Qt.NoFocus)
        self.keep_backup_check.setCursor(Qt.PointingHandCursor)
        self.keep_backup_check.setChecked(settings.keep_backup)
        self.keep_backup_check.toggled.connect(self._on_keep_backup_toggled)
        hint = QLabel("기본값: OFF (원본 파일은 덮어쓰지 않으므로 일반적으로는 필요 없습니다.)")
        hint.setStyleSheet("color: #666666; font-size: 9pt;")

        out_label = QLabel("완성본 저장 폴더:")
        self.output_dir_edit = QLineEdit()
        self.output_dir_edit.setObjectName("output_dir_edit")
        self.output_dir_edit.setReadOnly(True)
        self.output_dir_edit.setFrame(True)
        if settings.output_dir:
            self.output_dir_edit.setText(settings.output_dir)
        else:
            self.output_dir_edit.setPlaceholderText("기본 위치 사용 (ExcelSlimmed 폴더)")

        out_btn_row = QHBoxLayout()
        out_change_btn = QPushButton("변경...")
        out_reset_btn = QPushButton("기본값")
        out_change_btn.setCursor(Qt.PointingHandCursor)
        out_reset_btn.setCursor(Qt.PointingHandCursor)
        out_change_btn.clicked.connect(self._on_change_output_dir)
        out_reset_btn.clicked.connect(self._on_reset_output_dir)
        out_btn_row.addWidget(out_change_btn)
        out_btn_row.addWidget(out_reset_btn)
        out_btn_row.addStretch(1)

        path_layout.addWidget(self.keep_backup_check)
        path_layout.addWidget(hint)
        path_layout.addWidget(out_label)
        path_layout.addWidget(self.output_dir_edit)
        path_layout.addLayout(out_btn_row)

        settings_layout.addWidget(path_label)
        settings_layout.addWidget(path_group)

        # 이미지 설정 카드 (최대 해상도 / JPEG 품질)

        # 로그 설정 카드 (로그 상세도 / 오류 시 로그 폴더 자동 열기)
        log_label = QLabel("로그 설정")
        log_label.setStyleSheet("font-weight: 600;")

        log_settings_group = QGroupBox()
        log_settings_group.setStyleSheet(self._card_style())
        log_settings_layout = QVBoxLayout(log_settings_group)
        log_settings_layout.setSpacing(6)

        self.verbose_check = QCheckBox("상세 로그 기록")
        self.verbose_check.setFocusPolicy(Qt.NoFocus)
        self.verbose_check.setCursor(Qt.PointingHandCursor)
        self.verbose_check.setChecked(settings.log_mode == "verbose")

        self.open_log_on_error_check = QCheckBox("오류 발생 시 관련 로그 폴더 자동 열기")
        self.open_log_on_error_check.setFocusPolicy(Qt.NoFocus)
        self.open_log_on_error_check.setCursor(Qt.PointingHandCursor)
        self.open_log_on_error_check.setChecked(settings.open_log_on_error)

        self.verbose_check.toggled.connect(self._on_log_settings_changed)
        self.open_log_on_error_check.toggled.connect(self._on_log_settings_changed)

        log_settings_layout.addWidget(self.verbose_check)
        log_settings_layout.addWidget(self.open_log_on_error_check)

        settings_layout.addWidget(log_label)
        settings_layout.addWidget(log_settings_group)

        # UI 설정 카드 (다크 모드)
        ui_label = QLabel("다크 모드 설정")
        ui_label.setStyleSheet("font-weight: 600;")

        ui_group = QGroupBox()
        ui_group.setStyleSheet(self._card_style())
        ui_layout = QVBoxLayout(ui_group)
        ui_layout.setSpacing(6)

        self.dark_mode_check = QCheckBox("다크 모드 사용")
        self.dark_mode_check.setFocusPolicy(Qt.NoFocus)
        self.dark_mode_check.setCursor(Qt.PointingHandCursor)
        self.dark_mode_check.setChecked(self._theme == "dark")
        self.dark_mode_check.toggled.connect(self._on_dark_mode_toggled)

        ui_layout.addWidget(self.dark_mode_check)

        settings_layout.addWidget(ui_label)
        settings_layout.addWidget(ui_group)

        settings_layout.addStretch(1)

        s_layout.addWidget(settings_container, 0, Qt.AlignTop | Qt.AlignHCenter)

        self._update_precision_options_state()
        self.precision_check.toggled.connect(self._update_precision_options_state)
        self._update_image_controls_state()
        self.image_check.toggled.connect(self._update_image_controls_state)

        self._apply_global_widget_style()

    def _on_keep_backup_toggled(self, checked: bool) -> None:
        settings = get_settings()
        settings.keep_backup = bool(checked)
        save_settings(settings)

    def _on_change_output_dir(self) -> None:
        settings = get_settings()
        start_dir = settings.output_dir or str(Path.home() / "Desktop")
        directory = QFileDialog.getExistingDirectory(
            self,
            "완성본 저장 폴더 선택",
            start_dir,
        )
        if directory:
            settings.output_dir = directory
            save_settings(settings)
            self.output_dir_edit.setText(directory)

    def _on_reset_output_dir(self) -> None:
        settings = get_settings()
        settings.output_dir = ""
        save_settings(settings)
        self.output_dir_edit.clear()
        self.output_dir_edit.setPlaceholderText("기본 위치 사용 (ExcelSlimmed 폴더)")

    def _on_image_settings_changed(self) -> None:
        settings = get_settings()
        settings.image_max_edge = self.max_edge_slider.value()
        settings.image_quality = self.quality_slider.value()
        save_settings(settings)
        self.max_edge_label.setText(f"최대 해상도 (px): {settings.image_max_edge}")
        self.quality_label.setText(f"JPEG 품질: {settings.image_quality}%")
        self.max_edge_edit.setText(str(settings.image_max_edge))
        self.quality_edit.setText(str(settings.image_quality))

    def _on_log_settings_changed(self) -> None:
        settings = get_settings()
        settings.log_mode = "verbose" if self.verbose_check.isChecked() else "minimal"
        settings.open_log_on_error = self.open_log_on_error_check.isChecked()
        save_settings(settings)

    def _on_dark_mode_toggled(self, checked: bool) -> None:
        settings = get_settings()
        settings.theme = "dark" if checked else "light"
        save_settings(settings)
        self._theme = settings.theme
        self._apply_global_widget_style()
        self._refresh_card_styles()

    def _card_style(self) -> str:
        theme = getattr(self, "_theme", "light")
        if theme == "dark":
            # 다크 모드에서는 전체 배경보다 살짝 밝은 회색 카드 배경을 사용
            # 개별 QGroupBox 위젯에 직접 적용되도록 셀렉터 없이 스타일을 반환한다.
            return (
                "background-color: #222426;"  # 메인보다 약간 밝은 다크 그레이
                "border: 0px;"               # 테두리는 없애고 배경만 유지
                "border-radius: 4px;"
                "margin-top: 0px;"
            )
        # 라이트 모드 기본 카드 스타일
        return (
            "background-color: #ffffff;"
            "border: 0px;"               # 테두리는 없애고 배경만 유지
            "border-radius: 4px;"
            "margin-top: 0px;"
        )

    def _apply_global_widget_style(self) -> None:
        """Apply a theme-aware, uniform border to inputs and buttons."""

        theme = getattr(self, "_theme", "light")
        if theme == "dark":
            # 모던 브라우저와 비슷한 부드러운 다크 테마
            # 전체 배경은 짙은 회색(#1c1e21), 카드(QGroupBox)는 약간 더 밝은 회색(#222426)을 사용
            # 내부 텍스트 위젯(QLineEdit, QPlainTextEdit)은 기본 상태에서 박스 없이 카드 배경 위의 글씨처럼 보이게 한다.
            # 체크박스는 파란 배경(#5b8cff)에 흰색 체크 아이콘(check_white.svg)이 보이도록 설정한다.
            checkbox_checked_icon = str(Path(__file__).resolve().parent / "check_white.svg").replace("\\", "/")
            self.setStyleSheet(
                "* {"
                "  border: 0px;"
                "}"
                "QWidget {"
                "  background: #1c1e21;"   # 전체 배경 (짙은 회색)
                "  color: #f5f5f5;"        # 눈이 편한 거의 흰색 텍스트
                "}"
                "QTabWidget::pane {"
                "  border: 0px;"           # 탭 내용 영역 외곽 박스 제거
                "}"
                "QLineEdit {"
                "  background: transparent;"
                "  color: #f5f5f5;"
                "  border: 0px;"
                "  padding: 3px 6px;"
                "}"
                "QLineEdit[readOnly=\"true\"] {"
                "  background: transparent;"
                "  border: 0px;"
                "}"
                "QLineEdit:focus {"
                "  background: transparent;"  # 포커스 시에도 테두리 박스를 만들지 않는다
                "  border: 0px;"
                "}"
                "QPlainTextEdit {"
                "  background: transparent;"
                "  color: #f5f5f5;"
                "  border: 0px;"
                "}"
                "QPlainTextEdit:focus {"
                "  background: transparent;"
                "  border: 0px;"
                "}"
                "QPlainTextEdit[readOnly=\"true\"] {"
                "  background: transparent;"
                "  border: 0px;"
                "}"
                "QLineEdit#file_path_edit,"
                "QLineEdit#max_edge_edit,"
                "QLineEdit#quality_edit,"
                "QLineEdit#output_dir_edit {"
                "  background: #25272b;"
                "  border: 1px solid #858a96;"  # 다크 모드에서 눈에 잘 보이는 연한 회색 테두리
                "  border-radius: 4px;"
                "  padding: 3px 6px;"
                "}"
                "QLineEdit#file_path_edit:focus,"
                "QLineEdit#max_edge_edit:focus,"
                "QLineEdit#quality_edit:focus,"
                "QLineEdit#output_dir_edit:focus {"
                "  border: 1px solid #bfc5d4;"  # 포커스 시 살짝 더 밝은 회색/파랑 톤
                "}"
                "QProgressBar {"
                "  border: 0px;"           # 진행바 외곽 박스 제거
                "  background: #222426;"   # 카드 배경과 자연스럽게 맞춤
                "  text-align: center;"
                "}"
                "QProgressBar::chunk {"
                "  background-color: #5b8cff;"
                "  border-radius: 2px;"
                "}"
                "QPushButton {"
                "  border: 1px solid #858a96;"  # 버튼도 동일한 연한 회색 테두리
                "  border-radius: 4px;"
                "  padding: 4px 10px;"
                "  background: #2b2f36;"
                "}"
                "QPushButton:hover {"
                "  background: #383c42;"
                "}"
                "QPushButton:pressed {"
                "  background: #42474e;"
                "}"
                "QCheckBox {"
                "  spacing: 6px;"
                "}"
                "QCheckBox::indicator {"
                "  width: 16px;"
                "  height: 16px;"
                "  border-radius: 3px;"
                "  border: 1px solid #70757d;"
                "  background: #25272b;"
                "}"
                "QCheckBox::indicator:hover {"
                "  border-color: #82aaff;"
                "}"
                "QCheckBox::indicator:checked {"
                "  background: #5b8cff;"   # 파란색으로 꽉 찬 체크 박스
                "  border-color: #5b8cff;"
                f"  image: url('{checkbox_checked_icon}');"
                "}"
                "QCheckBox:focus {"
                "  outline: none;"          # 체크박스 줄 전체를 감싸는 포커스 테두리 제거
                "}"
            )
        else:
            # 라이트 모드에서는 전체를 흰 배경 위에 평평하게 두되,
            # 텍스트 입력/로그 영역은 기본 상태에서 박스 없이 글씨만 보이도록 한다.
            # 체크박스는 파란 배경(#5b8cff)에 흰색 체크 아이콘(check_white.svg)이 보이도록 설정한다.
            checkbox_checked_icon = str(Path(__file__).resolve().parent / "check_white.svg").replace("\\", "/")
            self.setStyleSheet(
                "* {"
                "  border: 0px;"
                "}"
                "QTabWidget::pane {"
                "  border: 0px;"           # 탭 내용 영역 외곽 박스 제거
                "}"
                "QLineEdit {"
                "  background: transparent;"
                "  border: 0px;"
                "  padding: 3px 6px;"
                "}"
                "QLineEdit[readOnly=\"true\"] {"
                "  background: transparent;"
                "  border: 0px;"
                "}"
                "QLineEdit:focus {"
                "  border: 0px;"
                "  background: transparent;"
                "}"
                "QPlainTextEdit {"
                "  background: transparent;"
                "  border: 0px;"
                "}"
                "QPlainTextEdit:focus {"
                "  background: transparent;"
                "  border: 0px;"
                "}"
                "QPlainTextEdit[readOnly=\"true\"] {"
                "  background: transparent;"
                "  border: 0px;"
                "}"
                "QLineEdit#file_path_edit,"
                "QLineEdit#max_edge_edit,"
                "QLineEdit#quality_edit,"
                "QLineEdit#output_dir_edit {"
                "  background: #ffffff;"
                "  border: 1px solid #d0d0d0;"  # 라이트 모드에서 자연스러운 연한 회색 테두리
                "  border-radius: 4px;"
                "  padding: 3px 6px;"
                "}"
                "QLineEdit#file_path_edit:focus,"
                "QLineEdit#max_edge_edit:focus,"
                "QLineEdit#quality_edit:focus,"
                "QLineEdit#output_dir_edit:focus {"
                "  border: 1px solid #5b8cff;"  # 포커스 시만 살짝 파란색으로 강조
                "}"
                "QProgressBar {"
                "  border: 0px;"           # 진행바 외곽 박스 제거
                "  background: #f0f0f0;"
                "  text-align: center;"
                "}"
                "QProgressBar::chunk {"
                "  background-color: #5b8cff;"
                "  border-radius: 2px;"
                "}"
                "QCheckBox {"
                "  spacing: 6px;"
                "}"
                "QCheckBox::indicator {"
                "  width: 16px;"
                "  height: 16px;"
                "  border-radius: 3px;"
                "  border: 1px solid #b0b0b0;"
                "  background: #ffffff;"
                "}"
                "QCheckBox::indicator:hover {"
                "  border-color: #5b8cff;"
                "}"
                "QCheckBox::indicator:checked {"
                "  background: #5b8cff;"   # 라이트 모드에서도 파란색으로 꽉 찬 체크 박스
                "  border-color: #5b8cff;"
                f"  image: url('{checkbox_checked_icon}');"
                "}"
                "QCheckBox:focus {"
                "  outline: none;"          # 체크박스 줄 전체를 감싸는 포커스 테두리 제거
                "}"
                "QPushButton {"
                "  border: 1px solid #d0d0d0;"  # 버튼도 동일한 연한 회색 테두리
                "  border-radius: 4px;"
                "  padding: 4px 10px;"
                "  background: #ffffff;"
                "}"
                "QPushButton:hover {"
                "  background: #f5f5f5;"
                "}"
                "QPushButton:pressed {"
                "  background: #eaeaea;"
                "}"
            )

        if theme == "dark":
            # 다크 모드: 1px 연한 회색 테두리
            input_style = (
                "background: #25272b;"
                "border: 1px solid #858a96;"
                "border-radius: 4px;"
                "padding: 3px 6px;"
            )
            button_style = (
                "border: 1px solid #858a96;"
                "border-radius: 4px;"
                "padding: 4px 10px;"
                "background: #2b2f36;"
            )
        else:
            # 라이트 모드: 1px 연한 회색 테두리
            input_style = (
                "background: #ffffff;"
                "border: 1px solid #d0d0d0;"
                "border-radius: 4px;"
                "padding: 3px 6px;"
            )
            button_style = (
                "border: 1px solid #d0d0d0;"
                "border-radius: 4px;"
                "padding: 4px 10px;"
                "background: #ffffff;"
            )

        for edit in (
            self.file_edit,
            self.max_edge_edit,
            self.quality_edit,
            self.output_dir_edit,
        ):
            edit.setStyleSheet(input_style)

        for btn in self.findChildren(QPushButton):
            btn.setStyleSheet(button_style)

    def _refresh_card_styles(self) -> None:
        """Re-apply card styles for the current theme.

        QGroupBox 들은 생성 시점에만 _card_style()을 적용하므로,
        테마를 전환할 때는 전체 그룹박스에 대해 스타일을 다시 입혀준다.
        """

        style = self._card_style()
        for group in self.findChildren(QGroupBox):
            group.setStyleSheet(style)

    def _update_precision_options_state(self) -> None:
        enabled = self.precision_check.isChecked()
        for cb in (self.aggressive_check, self.xmlcleanup_check, self.force_custom_check):
            if not enabled:
                cb.setChecked(False)
            cb.setEnabled(enabled)

    def _update_image_controls_state(self) -> None:
        enabled = self.image_check.isChecked()
        for w in (
            self.max_edge_label,
            self.max_edge_slider,
            self.max_edge_edit,
            self.max_edge_hint_label,
            self.quality_label,
            self.quality_slider,
            self.quality_edit,
            self.quality_hint_label,
        ):
            w.setEnabled(enabled)

    def _on_max_edge_edit_finished(self) -> None:
        text = self.max_edge_edit.text().strip()
        if not text:
            # 빈 입력이면 현재 슬라이더 값으로 복구
            self.max_edge_edit.setText(str(self.max_edge_slider.value()))
            return
        try:
            value = int(text)
        except ValueError:
            QMessageBox.warning(self, "입력 오류", "숫자만 입력 가능합니다.")
            self.max_edge_edit.setText(str(self.max_edge_slider.value()))
            return

        min_v = self.max_edge_slider.minimum()
        max_v = self.max_edge_slider.maximum()
        if value < min_v:
            QMessageBox.warning(self, "입력 범위", f"최소 {min_v}까지 가능합니다.")
            value = min_v
        elif value > max_v:
            QMessageBox.warning(self, "입력 범위", f"최대 {max_v}까지 가능합니다.")
            value = max_v

        self.max_edge_slider.setValue(value)

    def _on_quality_edit_finished(self) -> None:
        text = self.quality_edit.text().strip()
        if not text:
            self.quality_edit.setText(str(self.quality_slider.value()))
            return
        try:
            value = int(text)
        except ValueError:
            QMessageBox.warning(self, "입력 오류", "숫자만 입력 가능합니다.")
            self.quality_edit.setText(str(self.quality_slider.value()))
            return

        min_v = self.quality_slider.minimum()
        max_v = self.quality_slider.maximum()
        if value < min_v:
            QMessageBox.warning(self, "입력 범위", f"최소 {min_v}까지 가능합니다.")
            value = min_v
        elif value > max_v:
            QMessageBox.warning(self, "입력 범위", f"최대 {max_v}까지 가능합니다.")
            value = max_v

        self.quality_slider.setValue(value)

    def _on_browse(self) -> None:
        # 기본 경로를 바탕 화면으로 지정 (없으면 기본값 사용)
        default_dir = Path.home() / "Desktop"
        start_dir = str(default_dir) if default_dir.exists() else ""

        path, _ = QFileDialog.getOpenFileName(
            self,
            "대상 Excel 파일 선택",
            start_dir,
            "Excel Files (*.xlsx *.xlsm)",
        )
        if path:
            self.file_edit.setText(path)

    def _append_log(self, text: str) -> None:
        self.log_edit.appendPlainText(text)
        self.log_edit.verticalScrollBar().setValue(self.log_edit.verticalScrollBar().maximum())

    def _set_status(self, text: str, progress: float | None = None) -> None:
        self.status_label.setText(text)
        if progress is not None:
            self.progress_bar.setValue(int(progress))

    def _reset_ui_after_finish(self) -> None:
        """파이프라인 완료 후 기본 상태로 되돌립니다 (로그는 유지)."""
        self.file_edit.clear()
        # 초기 모드: 실행할 기능은 아무것도 선택하지 않은 상태
        self.clean_check.setChecked(False)
        self.image_check.setChecked(False)
        self.precision_check.setChecked(False)
        self.aggressive_check.setChecked(False)
        self.xmlcleanup_check.setChecked(False)
        self.force_custom_check.setChecked(False)
        self._update_precision_options_state()
        self.progress_bar.setValue(0)
        self.status_label.setText("준비됨")

    def _on_run_clicked(self) -> None:
        path_str = self.file_edit.text().strip()
        if not path_str:
            QMessageBox.warning(self, "안내", "대상 파일을 먼저 선택하세요.")
            return
        path = Path(path_str)
        if not path.exists():
            QMessageBox.critical(self, "오류", f"파일을 찾을 수 없습니다:\n{path}")
            return
        if path.suffix.lower() not in (".xlsx", ".xlsm"):
            QMessageBox.critical(self, "오류", "지원 형식은 .xlsx / .xlsm 입니다.")
            return
        if not (
            self.clean_check.isChecked()
            or self.image_check.isChecked()
            or self.precision_check.isChecked()
        ):
            QMessageBox.information(self, "안내", "실행할 기능을 하나 이상 선택하세요.")
            return

        # 정밀 슬리머를 켠 경우, 하위 옵션을 하나 이상 선택해야 실행
        if self.precision_check.isChecked() and not (
            self.aggressive_check.isChecked()
            or self.xmlcleanup_check.isChecked()
            or self.force_custom_check.isChecked()
        ):
            QMessageBox.warning(self, "안내", "정밀 슬리머 옵션을 1가지 이상 선택해야 합니다.")
            return

        self.log_edit.clear()
        self.progress_bar.setValue(0)
        self.status_label.setText("작업 시작...")
        self.run_button.setEnabled(False)

        worker = PipelineWorker(
            path=path,
            use_clean=self.clean_check.isChecked(),
            use_image=self.image_check.isChecked(),
            use_precision=self.precision_check.isChecked(),
            aggressive=self.aggressive_check.isChecked(),
            do_xml_cleanup=self.xmlcleanup_check.isChecked(),
            force_custom=self.force_custom_check.isChecked(),
        )
        worker.log.connect(self._append_log)
        worker.status.connect(self._set_status)

        def on_finished(final_path: str) -> None:
            self._set_status("모든 작업 완료", 100.0)
            self.run_button.setEnabled(True)
            QMessageBox.information(
                self,
                "완료",
                f"모든 작업이 완료되었습니다.\n\n최종 결과 파일:\n{final_path}",
            )
            # 탐색기 열기는 별도 스레드에서 실행해 UI 블로킹을 방지합니다.
            try:
                threading.Thread(
                    target=open_in_explorer_select,
                    args=(Path(final_path),),
                    daemon=True,
                ).start()
            except Exception:
                # 탐색기 열기 실패는 치명적이지 않으므로 조용히 무시합니다.
                pass
            # 로그는 유지하고 나머지 UI 상태만 초기화합니다.
            self._reset_ui_after_finish()

        def on_failed(msg: str) -> None:
            self.run_button.setEnabled(True)
            self._set_status("오류 발생", None)
            QMessageBox.critical(self, "오류", msg)

        worker.finished.connect(on_finished)
        worker.failed.connect(on_failed)

        # QThread 대신 표준 Python 스레드를 사용해 파이프라인 코어를 실행한다.
        # Qt 객체 생성/소멸은 모두 메인 스레드에서만 일어나고, 백그라운드에서는
        # PipelineWorker.run 이 순수 파이썬 코드와 시그널 emit만 수행한다.
        thread = threading.Thread(target=worker.run, daemon=True)
        thread.start()

        self._worker_thread = thread
        self._worker = worker


def main() -> None:
    app = QApplication(sys.argv)
    # 사무용에 적합하면서도 너무 딱딱하지 않은 기본 폰트 설정
    # 맑은 고딕이 설치되지 않은 환경에서는 Qt가 자동으로 대체 폰트를 사용합니다.
    app.setFont(QFont("맑은 고딕", 10))
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
