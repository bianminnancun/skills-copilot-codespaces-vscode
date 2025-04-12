import sys
import os
import datetime
import json
import logging
from logging.handlers import RotatingFileHandler
from PyQt5.QtWidgets import (QApplication, QWidget, QTableWidget, QTableWidgetItem, QPushButton,
                             QVBoxLayout, QCheckBox, QHeaderView, QMessageBox, QRadioButton,
                             QLabel, QLineEdit, QInputDialog, QProgressBar, QGridLayout,
                             QSlider, QDesktopWidget, QFrame, QHBoxLayout, QSystemTrayIcon, QMenu)
from PyQt5.QtCore import QTimer, Qt, QUrl, QPropertyAnimation, QRect, QLocale, QPoint, QTranslator
from PyQt5.QtGui import QFont, QIntValidator, QPalette, QColor, QIcon, QLinearGradient
from PyQt5.QtMultimedia import QMediaPlayer, QMediaContent
import ctypes

# 配置日志系统
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        RotatingFileHandler('boss_timer.log', maxBytes=1024*1024, backupCount=3, encoding='utf-8'),
        logging.StreamHandler()
    ]
)

def handle_exceptions(func):
    """异常处理装饰器"""
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            logging.error(f"Error in {func.__name__}: {str(e)}")
            QMessageBox.critical(args[0], "错误", f"发生错误: {str(e)}")
    return wrapper

class MarqueeAlert(QWidget):
    """优化的跑马灯提示窗口"""

    def __init__(self, message):
        super().__init__()
        self.setup_ui(message)
        self.position_window()
        QTimer.singleShot(100, self.setup_animation)

    def setup_ui(self, message):
        self.setWindowFlags(
            Qt.WindowStaysOnTopHint |
            Qt.FramelessWindowHint |
            Qt.Tool
        )
        self.setAttribute(Qt.WA_TranslucentBackground)

        self.setFixedSize(400, 40)
        self.container = QWidget(self)
        self.container.setGeometry(0, 0, self.width(), self.height())
        self.container.setStyleSheet("""
            QWidget {
                background-color: rgba(0, 0, 0, 0.8);
                border-radius: 2px;
            }
        """)

        self.label = QLabel(message, self.container)
        self.label.setStyleSheet("""
            QLabel {
                color: #FF0000;
                font-family: "Microsoft YaHei";
                font-size: 16px;
                font-weight: bold;
            }
        """)
        self.label.adjustSize()
        self.label.move(self.width(), (self.height() - self.label.height()) // 2)

    def position_window(self):
        """固定在屏幕左上角"""
        screen = QDesktopWidget().screenGeometry()
        self.move(screen.left() + 20, screen.top() + 20)

    def setup_animation(self):
        """设置从右向左的文字移动动画"""
        if not hasattr(self, 'label'):
            return

        start_x = self.width()
        end_x = -self.label.width()
        y_pos = (self.height() - self.label.height()) // 2

        self.anim = QPropertyAnimation(self.label, b"pos")
        self.anim.setDuration(5000)
        self.anim.setStartValue(QPoint(start_x, y_pos))
        self.anim.setEndValue(QPoint(end_x, y_pos))
        self.anim.finished.connect(self.close)
        self.anim.start()

    def closeEvent(self, event):
        if hasattr(self, 'anim'):
            self.anim.stop()
        super().closeEvent(event)

class BossTimer(QWidget):
    """主应用程序类"""

    VERSION = "1.0.0"

    def __init__(self):
        super().__init__()
        self.ringing = False
        self.alerts = []
        self.audio_file = None

        self.setup_ui()
        self.setup_media_player()
        self.setup_timers()
        self.setup_tray_icon()
        self.setup_translator()
        self.load_config()
        self.center_window()

    def resource_path(self, relative_path):
        """获取资源的绝对路径"""
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        
        return os.path.join(base_path, relative_path)

    def center_window(self):
        """窗口居中显示"""
        screen = QDesktopWidget().screenGeometry()
        size = self.geometry()
        self.move((screen.width() - size.width()) // 2, (screen.height() - size.height()) // 2)

    def setup_ui(self):
        """初始化UI"""
        self.setWindowTitle(f"BOSS定时预警系统 v{self.VERSION}")
        self.resize(1000, 650)
        self.setMinimumSize(900, 500)

        self.setAutoFillBackground(True)
        palette = self.palette()
        gradient = QLinearGradient(0, 0, 0, self.height())
        gradient.setColorAt(0, QColor(30, 35, 42))
        gradient.setColorAt(1, QColor(50, 55, 62))
        palette.setBrush(QPalette.Window, gradient)
        self.setPalette(palette)

        self.create_widgets()
        self.setup_layout()
        self.setup_connections()

    def create_widgets(self):
        """创建所有控件"""
        self.title_label = QLabel(f"BOSS定时预警系统 v{self.VERSION}")
        self.title_label.setStyleSheet("""
            QLabel {
                color: #4FC3F7;
                font-family: "Microsoft YaHei";
                font-size: 24px;
                font-weight: bold;
                qproperty-alignment: AlignCenter;
            }
        """)

        self.time_label = QLabel()
        self.time_label.setStyleSheet("""
            QLabel {
                color: #AAAAAA;
                font-family: "Microsoft YaHei";
                font-size: 12px;
                qproperty-alignment: AlignCenter;
            }
        """)

        self.table = QTableWidget(0, 9)
        headers = ["序号", "BOSS名称", "周期(分)", "周期(秒)", "上次刷新", "下次刷新", "剩余时间", "启用", "进度"]
        self.table.setHorizontalHeaderLabels(headers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.setStyleSheet("""
            QTableWidget {
                background-color: rgba(255, 255, 255, 0.1);
                border: 1px solid rgba(255, 255, 255, 0.2);
                border-radius: 4px;
                gridline-color: rgba(255, 255, 255, 0.1);
                font-family: "Microsoft YaHei";
                font-size: 12px;
                color: #DDDDDD;
            }
            QHeaderView::section {
                background-color: rgba(79, 195, 247, 0.3);
                color: white;
                padding: 6px;
                border: none;
                font-weight: bold;
            }
            QTableWidget::item {
                border-bottom: 1px solid rgba(255, 255, 255, 0.05);
            }
            QTableWidget::item:selected {
                background-color: rgba(79, 195, 247, 0.5);
            }
        """)

        self.new_btn = self.create_button("新建", "#4FC3F7")
        self.delete_btn = self.create_button("删除", "#F44336")
        self.stop_alarm_btn = self.create_button("静音", "#FF9800")
        self.update_btn = self.create_button("刷新", "#8BC34A")
        self.test_btn = self.create_button("测试声音", "#9C27B0")
        self.check_update_btn = self.create_button("检查更新", "#607D8B")

        self.auto_time = QRadioButton("自动模式")
        self.manual_time = QRadioButton("手动模式")
        self.auto_time.setChecked(True)

        radio_style = """
            QRadioButton {
                color: #DDDDDD;
                font-family: "Microsoft YaHei";
                font-size: 12px;
                spacing: 5px;
            }
            QRadioButton::indicator {
                width: 16px;
                height: 16px;
            }
            QRadioButton::indicator::unchecked {
                border: 1px solid #AAAAAA;
                border-radius: 8px;
                background: rgba(255, 255, 255, 0.1);
            }
            QRadioButton::indicator::checked {
                border: 1px solid #4FC3F7;
                border-radius: 8px;
                background: #4FC3F7;
            }
        """
        self.auto_time.setStyleSheet(radio_style)
        self.manual_time.setStyleSheet(radio_style)

        self.volume_slider = QSlider(Qt.Horizontal)
        self.volume_slider.setRange(0, 100)
        self.volume_slider.setValue(70)
        self.volume_slider.setStyleSheet("""
            QSlider {
                height: 20px;
            }
            QSlider::groove:horizontal {
                background: rgba(255, 255, 255, 0.1);
                height: 4px;
                border-radius: 2px;
            }
            QSlider::handle:horizontal {
                background: #4FC3F7;
                width: 16px;
                height: 16px;
                margin: -6px 0;
                border-radius: 8px;
            }
            QSlider::sub-page:horizontal {
                background: #4FC3F7;
                border-radius: 2px;
            }
        """)

        self.global_progress = QProgressBar()
        self.global_progress.setStyleSheet("""
            QProgressBar {
                background: rgba(255, 255, 255, 0.1);
                border: 1px solid rgba(255, 255, 255, 0.2);
                border-radius: 4px;
                height: 20px;
                text-align: center;
                color: white;
            }
            QProgressBar::chunk {
                background-color: #4FC3F7;
                border-radius: 3px;
            }
        """)

        self.info_frame = QFrame()
        self.info_frame.setStyleSheet("""
            QFrame {
                background-color: rgba(0, 0, 0, 0.2);
                border-top: 1px solid rgba(255, 255, 255, 0.1);
            }
        """)

        info_layout = QHBoxLayout(self.info_frame)
        info_layout.setContentsMargins(10, 5, 10, 5)

        self.credit_label = QLabel("米哈游制作")
        self.credit_label.setStyleSheet("""
            QLabel {
                color: rgba(255, 255, 255, 0.7);
                font-family: "Microsoft YaHei";
                font-size: 12px;
            }
        """)

        self.contact_label = QLabel("联系方式: 15853286172")
        self.contact_label.setStyleSheet("""
            QLabel {
                color: rgba(255, 255, 255, 0.7);
                font-family: "Microsoft YaHei";
                font-size: 12px;
                qproperty-alignment: AlignRight;
            }
        """)

        info_layout.addWidget(self.credit_label)
        info_layout.addStretch()
        info_layout.addWidget(self.contact_label)

    def create_button(self, text, color):
        """创建现代化按钮"""
        btn = QPushButton(text)
        btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {color};
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px 12px;
                font-family: "Microsoft YaHei";
                font-size: 12px;
                min-width: 80px;
            }}
            QPushButton:hover {{
                background-color: {self.adjust_color(color, 20)};
            }}
            QPushButton:pressed {{
                background-color: {self.adjust_color(color, -20)};
            }}
        """)
        return btn

    def adjust_color(self, hex_color, amount):
        """调整颜色亮度"""
        hex_color = hex_color.lstrip('#')
        rgb = tuple(int(hex_color[i:i + 2], 16) for i in (0, 2, 4))
        adjusted = []
        for channel in rgb:
            new_channel = min(255, max(0, channel + amount))
            adjusted.append(f"{new_channel:02x}")
        return '#' + ''.join(adjusted)

    def setup_layout(self):
        """设置现代化布局"""
        main_layout = QGridLayout()
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 0)

        main_layout.addWidget(self.title_label, 0, 0, 1, 4)
        main_layout.addWidget(self.time_label, 1, 0, 1, 4)
        main_layout.addWidget(self.table, 2, 0, 1, 4)

        button_layout = QVBoxLayout()
        button_layout.setSpacing(10)
        for btn in [self.new_btn, self.delete_btn, self.stop_alarm_btn, 
                   self.update_btn, self.test_btn, self.check_update_btn]:
            btn.setMinimumHeight(35)
            button_layout.addWidget(btn)
        main_layout.addLayout(button_layout, 3, 0)

        control_layout = QVBoxLayout()
        control_layout.setSpacing(10)

        mode_group = QVBoxLayout()
        mode_group.setSpacing(8)
        mode_group.addWidget(QLabel("模式选择:"))
        mode_group.addWidget(self.auto_time)
        mode_group.addWidget(self.manual_time)
        control_layout.addLayout(mode_group)

        control_layout.addWidget(QLabel("音量控制:"))
        control_layout.addWidget(self.volume_slider)
        main_layout.addLayout(control_layout, 3, 1)

        main_layout.addWidget(self.global_progress, 4, 0, 1, 4)
        main_layout.addWidget(self.info_frame, 5, 0, 1, 4)

        self.setLayout(main_layout)

    def setup_connections(self):
        """连接信号槽"""
        self.new_btn.clicked.connect(self.add_boss_row)
        self.delete_btn.clicked.connect(self.delete_selected_row)
        self.stop_alarm_btn.clicked.connect(self.stop_alarm_sound)
        self.update_btn.clicked.connect(self.force_refresh)
        self.volume_slider.valueChanged.connect(self.adjust_volume)
        self.table.cellDoubleClicked.connect(self.handle_cell_edit)
        self.test_btn.clicked.connect(self.test_audio)
        self.check_update_btn.clicked.connect(self.check_for_updates)

    def setup_tray_icon(self):
        """设置系统托盘图标"""
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(QIcon(self.resource_path("icon.ico")))
        
        menu = QMenu()
        show_action = menu.addAction("显示窗口")
        show_action.triggered.connect(self.show_normal)
        exit_action = menu.addAction("退出")
        exit_action.triggered.connect(self.close)
        
        self.tray_icon.setContextMenu(menu)
        self.tray_icon.show()
        
    def show_normal(self):
        """从托盘恢复窗口"""
        self.show()
        self.setWindowState(self.windowState() & ~Qt.WindowMinimized | Qt.WindowActive)
        self.activateWindow()

    def setup_translator(self):
        """设置翻译器"""
        self.translator = QTranslator()
        if self.translator.load(self.resource_path("boss_timer_zh.qm")):
            QApplication.instance().installTranslator(self.translator)

    @handle_exceptions
    def setup_media_player(self):
        """初始化音频播放器"""
        print("初始化音频播放器...")
        print(f"当前工作目录: {os.getcwd()}")

        # 确保sounds目录存在
        sounds_dir = self.resource_path("sounds")
        if not os.path.exists(sounds_dir):
            os.makedirs(sounds_dir)

        self.media_player = QMediaPlayer()
        self.media_player.setVolume(self.volume_slider.value())

        # 检查音频文件
        alarm_path = self.find_audio_file("alarm.wav")
        warning_path = self.find_audio_file("warning.wav")

        # 加载主警报音效
        self.alarm_sound = None
        if alarm_path:
            try:
                self.alarm_sound = QMediaContent(QUrl.fromLocalFile(alarm_path))
                print(f"成功加载alarm.wav: {alarm_path}")
            except Exception as e:
                print(f"加载alarm.wav失败: {str(e)}")

        # 检查音频文件是否存在
        if not alarm_path or not warning_path:
            msg = "缺少音频文件:\n"
            if not alarm_path:
                msg += "- alarm.wav\n"
            if not warning_path:
                msg += "- warning.wav\n"
            msg += "将使用系统蜂鸣音替代"
            print(msg)
            QMessageBox.warning(self, "音频文件缺失", msg)

    def find_audio_file(self, filename="alarm.wav"):
        """查找音频文件"""
        possible_paths = [
            self.resource_path(filename),
            self.resource_path(os.path.join("sounds", filename)),
            filename,
            os.path.join("sounds", filename)
        ]

        for path in possible_paths:
            if os.path.exists(path):
                print(f"找到音频文件: {path}")
                return path

        print(f"未找到音频文件: {filename}")
        return None

    @handle_exceptions
    def play_sound(self, filename):
        """播放指定音频文件"""
        print(f"尝试播放: {filename}")
        try:
            filepath = self.find_audio_file(filename)
            if filepath:
                print(f"播放文件路径: {filepath}")
                media = QMediaContent(QUrl.fromLocalFile(filepath))
                self.media_player.setMedia(media)
                self.media_player.setVolume(self.volume_slider.value())
                self.media_player.play()

                # 检查播放状态
                def check_play_status():
                    if self.media_player.state() != QMediaPlayer.PlayingState:
                        print(f"播放失败，状态: {self.media_player.state()}")
                        QApplication.beep()

                QTimer.singleShot(500, check_play_status)
            else:
                print(f"未找到文件: {filename}")
                QApplication.beep()
        except Exception as e:
            print(f"播放{filename}失败: {str(e)}")
            QApplication.beep()

    def play_warning_sound(self):
        """播放3分钟预警音"""
        print("触发3分钟预警音")
        self.play_sound("warning.wav")

    def setup_timers(self):
        """初始化定时器系统"""
        self.clock_timer = QTimer(self)
        self.clock_timer.timeout.connect(self.update_clock)
        self.clock_timer.start(1000)

        self.refresh_timer = QTimer(self)
        self.refresh_timer.timeout.connect(self.refresh_all_timers)
        self.refresh_timer.start(1000)

    def update_clock(self):
        """更新系统时钟显示"""
        self.time_label.setText(datetime.datetime.now().strftime("系统时间: %Y-%m-%d %H:%M:%S"))

    @handle_exceptions
    def add_boss_row(self):
        """添加新BOSS行"""
        row = self.table.rowCount()
        self.table.insertRow(row)

        self.table.setItem(row, 0, QTableWidgetItem(str(row + 1)))
        self.table.setItem(row, 1, QTableWidgetItem("新BOSS"))

        min_edit = QLineEdit("60")
        min_edit.setValidator(QIntValidator(1, 1440, self))
        min_edit.textChanged.connect(lambda: self.validate_input(row, 2))
        min_edit.setStyleSheet("""
            QLineEdit {
                background-color: rgba(255, 255, 255, 0.1);
                color: #DDDDDD;
                border: 1px solid rgba(255, 255, 255, 0.2);
                border-radius: 3px;
                padding: 3px;
            }
        """)
        self.table.setCellWidget(row, 2, min_edit)

        sec_edit = QLineEdit("0")
        sec_edit.setValidator(QIntValidator(0, 59, self))
        sec_edit.textChanged.connect(lambda: self.validate_input(row, 3))
        sec_edit.setStyleSheet(min_edit.styleSheet())
        self.table.setCellWidget(row, 3, sec_edit)

        current_time = datetime.datetime.now().strftime("%H:%M:%S")
        self.table.setItem(row, 4, QTableWidgetItem(current_time))
        self.table.setItem(row, 5, QTableWidgetItem(""))
        self.table.setItem(row, 6, QTableWidgetItem(""))

        enable_check = QCheckBox(self)
        enable_check.setChecked(True)
        enable_check.setStyleSheet("""
            QCheckBox {
                color: #DDDDDD;
                spacing: 5px;
            }
            QCheckBox::indicator {
                width: 16px;
                height: 16px;
            }
            QCheckBox::indicator:unchecked {
                border: 1px solid #AAAAAA;
                background: rgba(255, 255, 255, 0.1);
                border-radius: 3px;
            }
            QCheckBox::indicator:checked {
                border: 1px solid #4FC3F7;
                background: #4FC3F7;
                border-radius: 3px;
            }
        """)
        self.table.setCellWidget(row, 7, enable_check)

        progress = QProgressBar()
        progress.setStyleSheet("""
            QProgressBar {
                background: rgba(255, 255, 255, 0.1);
                border: 1px solid rgba(255, 255, 255, 0.2);
                border-radius: 3px;
                height: 16px;
                text-align: center;
                color: white;
            }
            QProgressBar::chunk {
                background-color: #FF9800;
                border-radius: 2px;
            }
        """)
        self.table.setCellWidget(row, 8, progress)

    def validate_input(self, row, col):
        """验证输入有效性"""
        widget = self.table.cellWidget(row, col)
        if widget:
            if widget.hasAcceptableInput():
                widget.setStyleSheet("""
                    QLineEdit {
                        background-color: rgba(255, 255, 255, 0.1);
                        color: #DDDDDD;
                        border: 1px solid rgba(255, 255, 255, 0.2);
                        border-radius: 3px;
                        padding: 3px;
                    }
                """)
            else:
                widget.setStyleSheet("""
                    QLineEdit {
                        background-color: rgba(255, 0, 0, 0.2);
                        color: #DDDDDD;
                        border: 1px solid #F44336;
                        border-radius: 3px;
                        padding: 3px;
                    }
                """)

    @handle_exceptions
    def delete_selected_row(self):
        """删除选中行"""
        row = self.table.currentRow()
        if row >= 0:
            self.table.removeRow(row)
            self.renumber_rows()
            logging.info(f"已删除第 {row + 1} 行")

    def renumber_rows(self):
        """重新编号所有行"""
        for row in range(self.table.rowCount()):
            self.table.setItem(row, 0, QTableWidgetItem(str(row + 1)))

    @handle_exceptions
    def handle_cell_edit(self, row, col):
        """处理单元格编辑"""
        if col == 4 and self.manual_time.isChecked():
            current = self.table.item(row, 4).text()
            new_time, ok = QInputDialog.getText(
                self, "设置时间", "格式: HH:MM:SS", text=current
            )
            if ok and self.is_valid_time(new_time):
                self.table.setItem(row, 4, QTableWidgetItem(new_time))
                self.refresh_all_timers()

    def is_valid_time(self, time_str):
        """验证时间格式"""
        try:
            datetime.datetime.strptime(time_str, "%H:%M:%S")
            return True
        except ValueError:
            QMessageBox.warning(self, "格式错误", "请输入有效时间 (HH:MM:SS)")
            return False

    @handle_exceptions
    def refresh_all_timers(self):
        """刷新所有计时器状态"""
        now = datetime.datetime.now()
        for row in range(self.table.rowCount()):
            try:
                if not self.is_row_valid(row):
                    continue

                period_min = int(self.table.cellWidget(row, 2).text())
                period_sec = int(self.table.cellWidget(row, 3).text())
                last_time = datetime.datetime.strptime(
                    self.table.item(row, 4).text(), "%H:%M:%S"
                ).time()

                last_dt = datetime.datetime.combine(now.date(), last_time)
                if last_dt > now:
                    last_dt -= datetime.timedelta(days=1)

                period = datetime.timedelta(minutes=period_min, seconds=period_sec)
                next_time = last_dt + period
                while next_time < now:
                    next_time += period

                self.update_row_display(row, next_time, period, now)
                self.check_alert_conditions(row, next_time, now)

            except Exception as e:
                logging.error(f"刷新第 {row + 1} 行失败: {str(e)}")

    def is_row_valid(self, row):
        """验证行数据有效性"""
        return all([
            self.table.item(row, 1),
            self.table.cellWidget(row, 2),
            self.table.cellWidget(row, 3),
            self.table.item(row, 4)
        ])

    def update_row_display(self, row, next_time, period, now):
        """更新行显示"""
        self.table.setItem(row, 5, QTableWidgetItem(next_time.strftime("%H:%M:%S")))

        remaining = (next_time - now).total_seconds()
        remaining_str = str(datetime.timedelta(seconds=int(remaining)))
        self.table.setItem(row, 6, QTableWidgetItem(remaining_str))

        total_sec = period.total_seconds()
        if total_sec > 0:
            progress = int((1 - remaining / total_sec) * 100)
            self.table.cellWidget(row, 8).setValue(progress)
            self.table.cellWidget(row, 8).setFormat(f"{progress}%")

    def check_alert_conditions(self, row, next_time, now):
        """检查警报触发条件"""
        if not self.table.cellWidget(row, 7).isChecked():
            return

        remaining = (next_time - now).total_seconds()
        print(f"BOSS {self.table.item(row, 1).text()} 剩余时间: {remaining}秒")

        if 178 <= remaining <= 180:
            print("触发3分钟预警")
            boss_name = self.table.item(row, 1).text()
            self.show_alert(f"{boss_name} 将在3分钟后刷新！")
            self.play_warning_sound()

        if remaining <= 0 and not self.ringing:
            print("触发BOSS刷新警报")
            self.trigger_alarm(row, next_time)

    def show_alert(self, message):
        """显示跑马灯预警"""
        alert = MarqueeAlert(message)
        alert.show()
        self.alerts.append(alert)
        self.alerts = [a for a in self.alerts if a.isVisible()]

    @handle_exceptions
    def trigger_alarm(self, row, next_time):
        """触发完整警报"""
        self.ringing = True
        boss_name = self.table.item(row, 1).text()
        alert_message = f"{boss_name} 已刷新！\n时间: {next_time.strftime('%H:%M:%S')}"

        self.play_sound("alarm.wav")

        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("BOSS刷新")
        msg_box.setText(alert_message)
        msg_box.setStandardButtons(QMessageBox.Ok)
        msg_box.finished.connect(self.handle_alert_close)
        msg_box.show()

        if self.auto_time.isChecked():
            self.update_boss_time(row)

    @handle_exceptions
    def test_audio(self):
        """测试音频系统"""
        print("\n测试音频系统...")
        print("1. 测试预警音...")
        self.play_sound("warning.wav")

        QTimer.singleShot(2000, lambda: (
            print("\n2. 测试警报音..."),
            self.play_sound("alarm.wav")
        ))

    def handle_alert_close(self):
        """警报窗口关闭后的清理"""
        self.stop_alarm_sound()

    @handle_exceptions
    def update_boss_time(self, row):
        """安全更新时间"""
        try:
            new_time = datetime.datetime.now().strftime("%H:%M:%S")
            self.table.item(row, 4).setText(new_time)
            self.refresh_all_timers()
        except Exception as e:
            logging.error(f"更新时间失败: {str(e)}")

    @handle_exceptions
    def stop_alarm_sound(self):
        """停止音效"""
        try:
            if hasattr(self, 'media_player') and self.media_player.state() == QMediaPlayer.PlayingState:
                print("停止播放音效")
                self.media_player.stop()
                self.media_player.setPosition(0)
        except Exception as e:
            logging.error(f"停止音效失败: {str(e)}")
        finally:
            self.ringing = False
            QApplication.processEvents()

    @handle_exceptions
    def adjust_volume(self, value):
        """调整音量"""
        try:
            volume = max(0, min(100, value))
            self.volume_slider.blockSignals(True)
            self.volume_slider.setValue(volume)
            self.volume_slider.blockSignals(False)

            if hasattr(self, 'media_player'):
                self.media_player.setVolume(volume)
                print(f"音量调整为: {volume}%")

        except Exception as e:
            logging.error(f"音量调整失败: {str(e)}")

    @handle_exceptions
    def force_refresh(self):
        """手动强制刷新"""
        self.refresh_all_timers()
        QMessageBox.information(self, "刷新完成", "所有计时器已刷新！")

    @handle_exceptions
    def check_for_updates(self):
        """检查更新"""
        try:
            import requests
            response = requests.get("https://api.github.com/repos/yourname/bosstimer/releases/latest")
            latest_version = response.json()["tag_name"]
            if latest_version > self.VERSION:
                QMessageBox.information(self, "更新", f"发现新版本 {latest_version}")
            else:
                QMessageBox.information(self, "更新", "当前已是最新版本")
        except Exception as e:
            logging.error(f"检查更新失败: {str(e)}")
            QMessageBox.warning(self, "更新", "检查更新失败")

    @handle_exceptions
    def load_config(self):
        """加载配置文件"""
        try:
            config_path = self.resource_path("boss_config.json")
            if os.path.exists(config_path):
                with open(config_path, "r", encoding="utf-8") as f:
                    config = json.load(f)
                    for item in config:
                        self.add_boss_row()
                        row = self.table.rowCount() - 1

                        if self.table.item(row, 1):
                            self.table.item(row, 1).setText(item.get("name", "新BOSS"))
                        if self.table.cellWidget(row, 2):
                            self.table.cellWidget(row, 2).setText(str(item.get("minutes", 60)))
                        if self.table.cellWidget(row, 3):
                            self.table.cellWidget(row, 3).setText(str(item.get("seconds", 0)))
                        if self.table.item(row, 4):
                            self.table.item(row, 4).setText(item.get("last_time", "00:00:00"))
                        if self.table.cellWidget(row, 7):
                            self.table.cellWidget(row, 7).setChecked(item.get("enabled", True))

                logging.info("配置加载成功")
            else:
                logging.info("未找到配置文件，将使用默认配置")
        except Exception as e:
            logging.error(f"配置加载失败: {str(e)}")

    @handle_exceptions
    def save_config(self):
        """保存配置到文件"""
        try:
            config = []
            for row in range(self.table.rowCount()):
                config.append({
                    "name": self.table.item(row, 1).text(),
                    "minutes": int(self.table.cellWidget(row, 2).text()),
                    "seconds": int(self.table.cellWidget(row, 3).text()),
                    "last_time": self.table.item(row, 4).text(),
                    "enabled": self.table.cellWidget(row, 7).isChecked()
                })

            config_path = self.resource_path("boss_config.json")
            with open(config_path, "w", encoding="utf-8") as f:
                json.dump(config, f, indent=2, ensure_ascii=False)

            logging.info("配置保存成功")
        except Exception as e:
            logging.error(f"配置保存失败: {str(e)}")

    def closeEvent(self, event):
        """处理窗口关闭事件"""
        reply = QMessageBox.question(
            self, '确认',
            "确定要退出吗? 程序将继续在后台运行。",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            self.save_config()
            for alert in self.alerts:
                alert.close()
            QApplication.quit()
            event.accept()
        else:
            self.hide()
            event.ignore()

if __name__ == '__main__':
    try:
        ctypes.windll.kernel32.FreeConsole()
    except:
        pass

    app = QApplication(sys.argv)
    QLocale.setDefault(QLocale(QLocale.Chinese))

    # 确保资源目录存在
    if not os.path.exists("sounds"):
        os.makedirs("sounds")

    window = BossTimer()
    window.show()
    sys.exit(app.exec_())
