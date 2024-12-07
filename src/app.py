import os
import sys
import logging
import traceback
import multiprocessing
from concurrent.futures import ThreadPoolExecutor, as_completed
import time
import psutil
import GPUtil

import cv2
from fpdf import FPDF
from pptx import Presentation
from pptx.util import Inches, Pt

from PyQt5.QtWidgets import (
    QApplication, QFileDialog, QVBoxLayout, QHBoxLayout, QComboBox,
    QPushButton, QLabel, QProgressBar, QMainWindow, QWidget, QMessageBox, QSpinBox,
    QTextEdit, QSplitter, QFrame, QGraphicsDropShadowEffect, QDesktopWidget
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer, QPropertyAnimation, QEasingCurve
from PyQt5.QtGui import QFont, QTextCursor, QColor, QIcon, QPalette

# Import additional utility functions
from videos_processor import extract_frames, create_pdf_from_frames, create_pptx_from_frames

class AdvancedLoggingHandler(logging.Handler):
    """Enhanced logging handler with more detailed formatting and severity levels."""
    def __init__(self, signal):
        super().__init__()
        self.signal = signal
        self.formatter = logging.Formatter(
            "%(asctime)s | %(levelname)8s | %(message)s", 
            datefmt="%Y-%m-%d %H:%M:%S"
        )

    def emit(self, record):
        try:
            msg = self.format(record)
            # Differentiate log levels with color coding
            if record.levelno == logging.ERROR:
                msg = f"🔴 {msg}"
            elif record.levelno == logging.WARNING:
                msg = f"🟠 {msg}"
            elif record.levelno == logging.INFO:
                msg = f"🟢 {msg}"
            self.signal.emit(msg)
        except Exception:
            self.handleError(record)

class ResourceMonitor:
    """Comprehensive system resource monitoring with GPU support."""
    @staticmethod
    def get_comprehensive_stats():
        """Gather detailed system and GPU resources."""
        stats = {
            "CPU": {
                "usage": f"{psutil.cpu_percent()}%",
                "cores": multiprocessing.cpu_count(),
                "frequency": f"{psutil.cpu_freq().current:.2f} MHz"
            },
            "Memory": {
                "total": f"{psutil.virtual_memory().total / (1024**3):.2f} GB",
                "used": f"{psutil.virtual_memory().percent}%",
                "available": f"{psutil.virtual_memory().available / (1024**3):.2f} GB"
            },
            "GPU": []
        }

        try:
            gpus = GPUtil.getGPUs()
            for gpu in gpus:
                stats["GPU"].append({
                    "name": gpu.name,
                    "memory_total": f"{gpu.memoryTotal} MB",
                    "memory_used": f"{gpu.memoryUsed} MB",
                    "memory_free": f"{gpu.memoryFree} MB",
                    "gpu_load": f"{gpu.load*100:.2f}%"
                })
        except Exception as e:
            stats["GPU"] = [{"error": str(e)}]

        return stats


class VideoConverterWorker(QThread):
    """Enhanced worker thread with comprehensive progress tracking."""
    progress_signal = pyqtSignal(int)
    detailed_progress_signal = pyqtSignal(dict)
    completed_signal = pyqtSignal(str)
    log_signal = pyqtSignal(str)
    resource_signal = pyqtSignal(dict)

    def __init__(self, video_paths, output_folder, output_format, frame_interval, max_workers=None):
        super().__init__()
        self.video_paths = video_paths
        self.output_folder = output_folder
        self.output_format = output_format
        self.frame_interval = frame_interval
        self.max_workers = max_workers or max(1, multiprocessing.cpu_count() - 1)

        # Enhanced logging setup
        self.logger = logging.getLogger(f"Worker-{id(self)}")
        self.logger.setLevel(logging.INFO)

        log_handler = AdvancedLoggingHandler(self.log_signal)
        self.logger.addHandler(log_handler)
        self.logger.propagate = False

        # Resource monitoring timer
        self.monitoring_timer = QTimer()
        self.monitoring_timer.timeout.connect(self.emit_resource_stats)
        self.monitoring_timer.start(2000)  # Every 2 seconds

    def emit_resource_stats(self):
        """Periodically emit system and GPU resource statistics."""
        try:
            stats = ResourceMonitor.get_comprehensive_stats()
            self.resource_signal.emit(stats)
        except Exception as e:
            self.logger.error(f"Resource monitoring error: {e}")

    def run(self):
        total_files = len(self.video_paths)
        progress_step = 100 / total_files
        start_time = time.time()

        # Detailed progress tracking
        progress_details = {
            "total_files": total_files,
            "processed_files": 0,
            "successful_conversions": 0,
            "failed_conversions": 0
        }



        for idx, video_path in enumerate(self.video_paths, 1):
            try:
                result = self.process_video(video_path)  # Call the processing function directly
                if result:
                    progress_details["successful_conversions"] += 1
                    self.logger.info(f"✅ Processed: {video_path} \n ----------------------------------------")
                else:
                    progress_details["failed_conversions"] += 1
                    self.completed_signal.emit(f"❌ Failed to process: {video_path} \n ---------------------------------------- ")
            except Exception as e:
                progress_details["failed_conversions"] += 1
                self.completed_signal.emit(f"❌ Error processing {video_path}: {e} \n ----------------------------------------")

            progress_details["processed_files"] = idx
            current_progress = int(idx * progress_step)
            
            # Emit progress signals
            self.progress_signal.emit(current_progress)
            self.detailed_progress_signal.emit(progress_details)


        # with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
        #     futures = {
        #         executor.submit(self.process_video, video_path): video_path
        #         for video_path in self.video_paths
        #     }

        #     for idx, future in enumerate(as_completed(futures), 1):
        #         video_path = futures[future]
        #         try:
        #             result = future.result()
        #             if result:
        #                 progress_details["successful_conversions"] += 1
        #                 # self.completed_signal.emit(f"✅ Processed: {video_path}")
        #             else:
        #                 progress_details["failed_conversions"] += 1
        #                 self.completed_signal.emit(f"❌ Failed to process: {video_path}")
        #         except Exception as e:
        #             progress_details["failed_conversions"] += 1
        #             self.completed_signal.emit(f"❌ Error processing {video_path}: {e}")

        #         progress_details["processed_files"] = idx
        #         current_progress = int(idx * progress_step)
                
        #         self.progress_signal.emit(current_progress)
        #         self.detailed_progress_signal.emit(progress_details)

        # Stop resource monitoring
        self.monitoring_timer.stop()

        total_time = time.time() - start_time
        final_message = (
            f"🏁 Conversion complete.\n"
            f"Total time: {total_time:.2f} seconds\n"
            f"Total files: {total_files}\n"
            f"Successful: {progress_details['successful_conversions']}\n"
            f"Failed: {progress_details['failed_conversions']}"
        )
        self.logger.info(final_message)
        self.completed_signal.emit(final_message)


    def process_video(self, video_path):
        try:
            self.logger.info(f"Processing: {video_path}")
            frames = extract_frames(video_path,interval=self.frame_interval, logger=self.logger)

            if not frames:
                self.logger.error(f"No frames extracted from {video_path}")
                return False

            base_name = os.path.splitext(os.path.basename(video_path))[0]  # Extract base name
            self.logger.info(f"Base Name : {base_name} ---- {video_path}")
            
            # Fix: Make sure you pass the folder as the second argument.
            if self.output_format == "pdf":
                return create_pdf_from_frames(frames, base_name, self.output_folder, logger=self.logger)  # Pass both frames and output_folder
            elif self.output_format == "pptx":
                return create_pptx_from_frames(frames, base_name, self.output_folder, logger=self.logger)  # Same for pptx
            else:
                self.logger.error(f"Unsupported format: {self.output_format}")
                return False
        except Exception as e:
            self.logger.error(f"Comprehensive error processing {video_path}: {traceback.format_exc()}")
            return False





class VideoConverterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("🎬 Advanced Video Converter")
        self.setup_global_styles()
        self.setup_ui()
        self.center_on_screen()

        self.frame_interval = 30  # Initialize frame_interval with the default value

    def setup_global_styles(self):
        """Set up global application styling."""
        # Define a consistent color palette
        self.COLOR_PRIMARY = "#3498db"      # Bright blue
        self.COLOR_SECONDARY = "#2ecc71"    # Bright green
        self.COLOR_BACKGROUND = "#ecf0f1"   # Light gray background
        self.COLOR_TEXT = "#2c3e50"         # Dark blue-gray for text
        self.COLOR_ACCENT = "#e74c3c"       # Bright red for highlights

        # Set up a global font
        global_font = QFont("Segoe UI", 10)
        QApplication.setFont(global_font)

    def setup_ui(self):
        """Comprehensive UI setup with modern design principles."""
        self.setGeometry(100, 100, 1400, 900)
        self.setStyleSheet(f"""
            QMainWindow {{
                background-color: {self.COLOR_BACKGROUND};
            }}
            QWidget {{
                color: {self.COLOR_TEXT};
            }}
        """)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        
        main_layout = QVBoxLayout(self.central_widget)
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(30, 30, 30, 30)

        # Create and add sections with consistent styling
        top_section = self.create_top_section()
        progress_section = self.create_progress_section()
        log_section = self.create_log_section()

        # Add shadow or elevation effect to sections
        for section in [top_section, progress_section, log_section]:
            section.setStyleSheet(f"""
                QWidget {{
                    background-color: white;
                    border-radius: 10px;
                    padding: 20px;
                }}
            """)

        main_layout.addWidget(top_section)
        main_layout.addWidget(progress_section)
        main_layout.addWidget(log_section)

    def update_frame_interval(self):
        """Update the frame_interval value."""
        self.frame_interval = self.frame_interval_spin_box.value() 
        print(f"frame_interval set to: {self.frame_interval} seconds")

    def create_top_section(self):
        """Create the top section with file and folder selection."""
        top_widget = QWidget()
        top_layout = QHBoxLayout(top_widget)
        top_layout.setSpacing(20)

        # Consistent button and input styling
        button_style = f"""
            QPushButton {{
                background-color: {self.COLOR_PRIMARY};
                color: white;
                border: none;
                border-radius: 6px;
                padding: 10px 15px;
                font-weight: bold;
            }}
            QPushButton:hover {{
                background-color: {self.COLOR_SECONDARY};
            }}
        """

        # File selection
        file_section = QVBoxLayout()
        file_label = self.create_section_label("📁 Selected Files")
        self.file_path_label = QLabel("No files selected")
        self.file_button = QPushButton("Browse Files")
        self.file_button.setStyleSheet(button_style)
        self.file_button.clicked.connect(self.select_files)
        file_section.addWidget(file_label)
        file_section.addWidget(self.file_path_label)
        file_section.addWidget(self.file_button)
        top_layout.addLayout(file_section)

        # Folder selection
        folder_section = QVBoxLayout()
        folder_label = self.create_section_label("💾 Output Folder")
        self.folder_path_label = QLabel("No folder selected")
        self.folder_button = QPushButton("Browse Folder")
        self.folder_button.setStyleSheet(button_style)
        self.folder_button.clicked.connect(self.select_folder)
        folder_section.addWidget(folder_label)
        folder_section.addWidget(self.folder_path_label)
        folder_section.addWidget(self.folder_button)
        top_layout.addLayout(folder_section)


        # Frame interval selection
        frame_interval_section = QVBoxLayout()
        frame_interval_label = self.create_section_label("⏱️ Frame Extraction Interval (seconds)")
        self.frame_interval_default_label = QLabel(f"Default: 30 seconds")
        # self.frame_interval_default_label.setStyleSheet(f"""
        #     QLabel {{
        #         font-size: 12px;
        #         color: {self.COLOR_PRIMARY};
        #         margin-bottom: 5px;
        #     }}
        # """)

        self.frame_interval_spin_box = QSpinBox()
        self.frame_interval_spin_box.setStyleSheet(f"""
            QSpinBox {{
                border: 1px solid {self.COLOR_PRIMARY};
                border-radius: 6px;
                padding: 8px;
                background-color: white;
            }}
            QSpinBox::up-button, QSpinBox::down-button {{
                background-color: {self.COLOR_PRIMARY};
                color: white;
                width: 20px;
                border-radius: 3px;
            }}
        """)
        self.frame_interval_spin_box.setMinimum(1)
        self.frame_interval_spin_box.setMaximum(60)
        self.frame_interval_spin_box.setValue(30)  # Default value
        self.frame_interval_spin_box.valueChanged.connect(self.update_frame_interval)

        frame_interval_section.addWidget(frame_interval_label)
        frame_interval_section.addWidget(self.frame_interval_default_label)  # Show default value
        frame_interval_section.addWidget(self.frame_interval_spin_box)
        top_layout.addLayout(frame_interval_section)


        # Conversion controls
        controls_section = QVBoxLayout()
        controls_label = self.create_section_label("⚙️ Conversion Settings")
        
        self.output_format = QComboBox()
        self.output_format.setStyleSheet(f"""
            QComboBox {{
                border: 1px solid {self.COLOR_PRIMARY};
                border-radius: 6px;
                padding: 8px;
                background-color: white;
            }}
            QComboBox::drop-down {{
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 30px;
                border-left-width: 1px;
                border-left-color: {self.COLOR_PRIMARY};
                border-left-style: solid;
                background-color: {self.COLOR_PRIMARY};
            }}
        """)
        self.output_format.addItems(["PDF", "PPTX"])
        
        # Set default value to PPTX
        self.output_format.setCurrentText("PPTX")  

        self.convert_button = QPushButton("Start Conversion")
        self.convert_button.setStyleSheet(f"""
            QPushButton {{
                background-color: {self.COLOR_ACCENT};
                color: white;
                border: none;
                border-radius: 6px;
                padding: 12px 20px;
                font-weight: bold;
                font-size: 12pt;
            }}
            QPushButton:hover {{
                background-color: #c0392b;
            }}
        """)
        self.convert_button.clicked.connect(self.convert_videos)
        
        controls_section.addWidget(controls_label)
        controls_section.addWidget(self.output_format)
        controls_section.addWidget(self.convert_button)
        top_layout.addLayout(controls_section)



        # # Conversion controls
        # controls_section = QVBoxLayout()
        # controls_label = self.create_section_label("⚙️ Conversion Settings")
        # self.output_format = QComboBox()
        # self.output_format.setStyleSheet(f"""
        #     QComboBox {{
        #         border: 1px solid {self.COLOR_PRIMARY};
        #         border-radius: 6px;
        #         padding: 8px;
        #         background-color: white;
        #     }}
        #     QComboBox::drop-down {{
        #         subcontrol-origin: padding;
        #         subcontrol-position: top right;
        #         width: 30px;
        #         border-left-width: 1px;
        #         border-left-color: {self.COLOR_PRIMARY};
        #         border-left-style: solid;
        #         background-color: {self.COLOR_PRIMARY};
        #     }}
        # """)
        # self.output_format.addItems(["PDF", "PPTX"])
        # self.convert_button = QPushButton("Start Conversion")
        # self.convert_button.setStyleSheet(f"""
        #     QPushButton {{
        #         background-color: {self.COLOR_ACCENT};
        #         color: white;
        #         border: none;
        #         border-radius: 6px;
        #         padding: 12px 20px;
        #         font-weight: bold;
        #         font-size: 12pt;
        #     }}
        #     QPushButton:hover {{
        #         background-color: #c0392b;
        #     }}
        # """)
        # self.convert_button.clicked.connect(self.convert_videos)
        # controls_section.addWidget(controls_label)
        # controls_section.addWidget(self.output_format)
        # controls_section.addWidget(self.convert_button)
        # top_layout.addLayout(controls_section)

        return top_widget



    def create_progress_section(self):
        """Create progress and system stats section with improved design."""
        progress_widget = QWidget()
        progress_layout = QVBoxLayout(progress_widget)

        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFormat("%p% - Processing Videos")
        self.progress_bar.setStyleSheet(f"""
            QProgressBar {{
                border: 2px solid {self.COLOR_PRIMARY};
                border-radius: 8px;
                text-align: center;
                color: {self.COLOR_TEXT};
                background-color: {self.COLOR_BACKGROUND};
                padding: 4px;
                height: 25px;
            }}
            QProgressBar::chunk {{
                background: qlineargradient(
                    x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 {self.COLOR_PRIMARY}, 
                    stop: 1 {self.COLOR_SECONDARY}
                );
                border-radius: 6px;
            }}
        """)

        self.detailed_progress_label = QLabel("Detailed Progress: Waiting to start")
        self.detailed_progress_label.setStyleSheet(f"font-weight: bold; color: {self.COLOR_TEXT};")
        
        self.system_stats_label = QLabel("System Resources: Idle")
        self.system_stats_label.setWordWrap(True)
        self.system_stats_label.setStyleSheet(f"color: {self.COLOR_TEXT}; opacity: 0.7;")

        progress_layout.addWidget(self.progress_bar)
        progress_layout.addWidget(self.detailed_progress_label)
        progress_layout.addWidget(self.system_stats_label)

        return progress_widget
    
    def create_log_section(self):
        """Create log output section with enhanced design."""
        log_widget = QWidget()
        log_layout = QVBoxLayout(log_widget)
        log_label = self.create_section_label("📋 Conversion Logs")
        
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setStyleSheet(f"""
            QTextEdit {{
                border: 1px solid {self.COLOR_PRIMARY};
                border-radius: 8px;
                padding: 15px;
                background-color: white;
                font-family: 'Consolas', 'Courier New', monospace;
            }}
        """)
        
        log_layout.addWidget(log_label)
        log_layout.addWidget(self.log_output)

        return log_widget


    def create_section_label(self, text):
        """Create a consistent and stylish section label."""
        label = QLabel(text)
        label.setStyleSheet(f"""
            font-weight: bold;
            color: {self.COLOR_TEXT};
            margin-bottom: 10px;
            font-size: 12pt;
        """)
        return label

    def center_on_screen(self):
        """Center the application window on the screen."""
        frame_geometry = self.frameGeometry()
        screen_center = QDesktopWidget().availableGeometry().center()
        frame_geometry.moveCenter(screen_center)
        self.move(frame_geometry.topLeft())

    def apply_modern_styling(self):
        """Apply modern, sleek styling to the application."""
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f0f0f0;
                font-family: 'Segoe UI', Arial, sans-serif;
            }
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border-radius: 5px;
                padding: 8px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QLabel {
                color: #333;
                font-size: 14px;
            }
            QProgressBar {
                border: 2px solid #4CAF50;
                border-radius: 5px;
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: #4CAF50;
            }
            QTextEdit {
                background-color: white;
                border: 1px solid #ddd;
                border-radius: 5px;
            }
        """)

    def select_files(self):
        """Enhanced file selection with multiple file support and preview."""
        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.ExistingFiles)
        file_dialog.setNameFilter("Video Files (*.mp4 *.mkv *.avi *.mov)")
        
        if file_dialog.exec_():
            self.video_paths = file_dialog.selectedFiles()
            preview_text = ", ".join([os.path.basename(path) for path in self.video_paths[:5]])
            if len(self.video_paths) > 5:
                preview_text += f" ...and {len(self.video_paths) - 5} more"
            
            self.file_path_label.setText(f"{len(self.video_paths)} files selected: {preview_text}")

    def select_folder(self):
        """Enhanced folder selection with path preview."""
        folder_dialog = QFileDialog()
        self.output_folder = folder_dialog.getExistingDirectory(self, "Select Output Folder")
        
        if self.output_folder:
            self.folder_path_label.setText(f"📂 {self.output_folder}")



    def convert_videos(self):
        """Enhanced video conversion with comprehensive error checking."""
        if not hasattr(self, 'video_paths') or not self.video_paths:
            QMessageBox.warning(self, "Error", "Please select video files.")
            return

        if not hasattr(self, 'output_folder') or not self.output_folder:
            QMessageBox.warning(self, "Error", "Please select an output folder.")
            return

        # Reset the progress bar to 0
        self.progress_bar.setValue(0)
        
        output_format = self.output_format.currentText().lower()
        
        self.worker_thread = VideoConverterWorker(
            video_paths=self.video_paths,
            output_folder=self.output_folder,
            output_format=output_format,
            frame_interval=self.frame_interval
        )

        # Connect signals for progress, logging, and completion
        self.worker_thread.progress_signal.connect(self.update_progress_bar)
        self.worker_thread.detailed_progress_signal.connect(self.update_detailed_progress)
        self.worker_thread.completed_signal.connect(self.show_completion_message)
        self.worker_thread.log_signal.connect(self.append_to_log)
        self.worker_thread.resource_signal.connect(self.update_system_stats)

        # Start the worker thread
        self.worker_thread.start()

        # Disable the convert button while processing
        self.convert_button.setEnabled(False)
        self.convert_button.setText("Processing...")


    def update_progress_bar(self, progress):
        """Update the main progress bar."""
        self.progress_bar.setValue(progress)

    def update_detailed_progress(self, details):
        """Update the detailed progress label."""
        self.detailed_progress_label.setText(
            f"Processed: {details['processed_files']}/{details['total_files']} | "
            f"Successful: {details['successful_conversions']} | "
            f"Failed: {details['failed_conversions']}"
        )

    def show_completion_message(self, message):
        """Show a message when the conversion is complete."""
        QMessageBox.information(self, "Conversion Complete", message)

        # Reset the progress bar and set the custom completion message
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("Click on Start Conversion again")
        
        # Re-enable the convert button after processing
        self.convert_button.setEnabled(True)
        self.convert_button.setText("Start Conversion")



    def append_to_log(self, log_message):
        """Append log messages to the log output."""
        self.log_output.append(log_message)
        self.log_output.moveCursor(QTextCursor.End)

    def update_system_stats(self, stats):
        """Update system resource statistics."""
        cpu_info = stats["CPU"]
        memory_info = stats["Memory"]
        gpu_info = stats["GPU"]

        gpu_stats = "\n".join([
            f"{gpu['name']}: {gpu['memory_used']}/{gpu['memory_total']} MB, Load: {gpu['gpu_load']}"
            for gpu in gpu_info
        ]) if gpu_info else "No GPU detected."

        self.system_stats_label.setText(
            f"CPU Usage: {cpu_info['usage']} | Cores: {cpu_info['cores']} | Freq: {cpu_info['frequency']}\n"
            f"Memory Usage: {memory_info['used']} | Available: {memory_info['available']}\n"
            f"GPU Stats:\n{gpu_stats}"
        )


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = VideoConverterApp()
    window.show()
    sys.exit(app.exec_())
