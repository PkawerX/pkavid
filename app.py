import sys
import json
from pathlib import Path
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QPushButton, QFileDialog, 
                             QGroupBox, QComboBox)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
import cv2
import win32gui
import win32con
import win32api
import win32ui
from ctypes import windll, Structure, c_long, WINFUNCTYPE, c_void_p, POINTER
from struct import pack
import time

class RECT(Structure):
    _fields_ = [('left', c_long), ('top', c_long), ('right', c_long), ('bottom', c_long)]

def get_monitors():
    """Get all monitor information including primary monitor."""
    monitors = []
    def callback(hMonitor, hdcMonitor, lprcMonitor, dwData):
        rect = lprcMonitor.contents
        monitor_info = win32api.GetMonitorInfo(hMonitor)
        is_primary = monitor_info.get('Flags') == 1
        monitors.append({
            'handle': hMonitor,
            'x': rect.left,
            'y': rect.top,
            'width': rect.right - rect.left,
            'height': rect.bottom - rect.top,
            'is_primary': is_primary,
            'device': monitor_info.get('Device', 'Unknown')
        })
        return True

    MONITORENUMPROC = WINFUNCTYPE(c_void_p, c_void_p, c_void_p, POINTER(RECT), c_void_p)
    windll.user32.EnumDisplayMonitors(None, None, MONITORENUMPROC(callback), 0)
    return monitors

def find_workerw():
    """Find the WorkerW window that contains the desktop background."""
    def callback(hwnd, hwnds):
        if win32gui.GetClassName(hwnd) == "WorkerW":
            temp = win32gui.FindWindowEx(hwnd, 0, "SHELLDLL_DefView", None)
            if temp:
                worker_w = win32gui.FindWindowEx(0, hwnd, "WorkerW", None)
                hwnds.append(worker_w)
        return True

    hwnds = []
    progman = win32gui.FindWindow("Progman", None)
    win32gui.SendMessageTimeout(progman, 0x052C, 0, 0, win32con.SMTO_NORMAL, 1000)
    win32gui.EnumWindows(callback, hwnds)
    return hwnds[0] if hwnds else None

def create_bmi_header(width, height):
    """Creates the BITMAPINFOHEADER structure for the bitmap."""
    return pack(
        'liiHHIIIIII',
        40,  # Size of header
        width,
        -height,  # Negative height for top-down bitmap
        1,  # Planes
        32,  # Bits per pixel
        0,  # Compression 
        width * height * 4,  # Image size
        0,  # X pixels per meter
        0,  # Y pixels per meter
        0,  # Colors used
        0   # Important colors
    )

class VideoPlayerThread(QThread):
    error_occurred = pyqtSignal(str)
    fps_updated = pyqtSignal(float)  # Signal to update the live FPS
    
    def __init__(self, monitor_config):
        super().__init__()
        self.monitor_config = monitor_config
        self.running = True
    
    def run(self):
        caps = {}
        save_dc = None
        mfc_dc = None
        hwnd_dc = None
        bmp = None
        worker_w = None

        try:
            worker_w = find_workerw()
            if not worker_w:
                self.error_occurred.emit("Could not find WorkerW window")
                return

            hwnd_dc = win32gui.GetWindowDC(worker_w)
            mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
            save_dc = mfc_dc.CreateCompatibleDC()

            monitors = get_monitors()
            min_x = min(monitor['x'] for monitor in monitors)
            min_y = min(monitor['y'] for monitor in monitors)
            total_width = max(monitor['x'] + monitor['width'] - min_x for monitor in monitors)
            total_height = max(monitor['y'] + monitor['height'] - min_y for monitor in monitors)

            bmp = win32ui.CreateBitmap()
            bmp.CreateCompatibleBitmap(mfc_dc, total_width, total_height)
            save_dc.SelectObject(bmp)

            for monitor_id, config in self.monitor_config.items():
                if 'video_path' in config and config['video_path']:
                    cap = cv2.VideoCapture(config['video_path'])
                    if cap.isOpened():
                        caps[monitor_id] = {
                            'cap': cap,
                            'info': config['monitor_info'],
                            'fps': config.get('fps', 30)  # Get FPS value from config
                        }
                        # Update the ComboBox with supported FPS from the video
                        fps = cap.get(cv2.CAP_PROP_FPS)
                        self.fps_updated.emit(fps)
                    else:
                        self.error_occurred.emit(f"Could not open video: {config['video_path']}")

            last_time = time.time()
            while self.running and caps:
                save_dc.FillSolidRect((0, 0, total_width, total_height), win32api.RGB(0, 0, 0))

                for monitor_id, data in caps.items():
                    cap = data['cap']
                    monitor = data['info']
                    fps = data['fps']
                    
                    ret, frame = cap.read()
                    if not ret:
                        cap.set(cv2.CAP_PROP_POS_FRAMES, 0)
                        ret, frame = cap.read()
                        if not ret:
                            continue

                    frame = cv2.cvtColor(frame, cv2.COLOR_BGR2BGRA)
                    video_width = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
                    video_height = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))

                    windll.gdi32.StretchDIBits(
                        save_dc.GetHandleOutput(),
                        monitor['x'] - min_x, monitor['y'] - min_y,  # Adjusted position
                        monitor['width'], monitor['height'],  # Destination dimensions
                        0, 0,  # Source position
                        video_width, video_height,  # Source dimensions
                        frame.tobytes(),
                        create_bmi_header(video_width, video_height),
                        win32con.DIB_RGB_COLORS,
                        win32con.SRCCOPY
                    )

                try:
                    mfc_dc.BitBlt(
                        (0, 0),
                        (total_width, total_height),
                        save_dc,
                        (0, 0),
                        win32con.SRCCOPY
                    )
                except Exception as e:
                    print(f"Error in BitBlt operation: {str(e)}")

                # Update live FPS every second
                if time.time() - last_time >= 1:
                    last_time = time.time()
                    self.fps_updated.emit(fps)

                cv2.waitKey(int(1000 / fps))  # Use selected FPS

        except Exception as e:
            self.error_occurred.emit(str(e))
        
        finally:
            for data in caps.values():
                data['cap'].release()
            if save_dc:
                save_dc.DeleteDC()
            if mfc_dc:
                mfc_dc.DeleteDC()
            if hwnd_dc and worker_w:
                win32gui.ReleaseDC(worker_w, hwnd_dc)
            if bmp:
                win32gui.DeleteObject(bmp.GetHandle())

    def stop(self):
        self.running = False
        self.wait()

class LiveWallpaperManager(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Live Wallpaper Manager")
        self.setMinimumSize(800, 400)
        
        self.config_file = Path("wallpaper_config.json")
        self.monitor_configs = {}
        self.video_player = None
        
        self.init_ui()
        self.load_config()
        
    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        monitors_group = QGroupBox("Monitor Configuration")
        monitors_layout = QVBoxLayout()
        
        self.monitors = get_monitors()
        self.monitor_widgets = {}
        
        for i, monitor in enumerate(self.monitors):
            monitor_id = f"monitor_{monitor['handle']}"
            title = f"Monitor {i + 1} - {'Primary' if monitor['is_primary'] else 'Secondary'}"
            group = QGroupBox(title)
            layout = QVBoxLayout()
            
            info_label = QLabel(
                f"Position: ({monitor['x']}, {monitor['y']})\n"
                f"Size: {monitor['width']}x{monitor['height']}\n"
                f"Device: {monitor['device']}"
            )
            layout.addWidget(info_label)
            
            file_layout = QHBoxLayout()
            path_label = QLabel("No video selected")
            path_label.setWordWrap(True)
            select_btn = QPushButton("Select Video")
            select_btn.clicked.connect(lambda checked, mid=monitor_id: 
                                       self.select_video(mid))
            
            file_layout.addWidget(path_label)
            file_layout.addWidget(select_btn)
            layout.addLayout(file_layout)
            
            # Add FPS selection ComboBox
            fps_combo = QComboBox()
            fps_combo.addItems(['24', '30', '60', '120'])
            fps_combo.currentIndexChanged.connect(lambda index, mid=monitor_id: 
                                                   self.update_fps(mid, fps_combo))
            layout.addWidget(QLabel("Select FPS:"))
            layout.addWidget(fps_combo)
            
            # Add label to show live FPS
            self.fps_label = QLabel("Live FPS: 0")
            layout.addWidget(self.fps_label)
            
            group.setLayout(layout)
            monitors_layout.addWidget(group)
            
            self.monitor_widgets[monitor_id] = {
                'path_label': path_label,
                'select_btn': select_btn,
                'monitor_info': monitor,
                'fps_combo': fps_combo,
                'fps_label': self.fps_label  # Reference to live FPS label
            }
        
        monitors_group.setLayout(monitors_layout)
        main_layout.addWidget(monitors_group)
        
        controls_layout = QHBoxLayout()
        
        start_btn = QPushButton("Start Wallpapers")
        start_btn.clicked.connect(self.start_wallpapers)
        
        stop_btn = QPushButton("Stop Wallpapers")
        stop_btn.clicked.connect(self.stop_wallpapers)
        
        controls_layout.addWidget(start_btn)
        controls_layout.addWidget(stop_btn)
        
        main_layout.addLayout(controls_layout)
    
    def select_video(self, monitor_id):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Video File",
            "",
            "Video Files (*.mp4 *.avi *.mkv);;All Files (*.*)"
        )
        
        if file_path:
            monitor_info = self.monitor_widgets[monitor_id]['monitor_info']
            self.monitor_widgets[monitor_id]['path_label'].setText(file_path)
            self.monitor_configs[monitor_id] = {
                'video_path': file_path,
                'monitor_info': monitor_info,
                'fps': 30  # Default FPS value
            }
            self.save_config()
    
    def update_fps(self, monitor_id, fps_combo):
        fps = int(fps_combo.currentText())
        if monitor_id in self.monitor_configs:
            self.monitor_configs[monitor_id]['fps'] = fps
            self.save_config()

    def load_config(self):
        if self.config_file.exists():
            with open(self.config_file, 'r') as f:
                self.monitor_configs = json.load(f)
                
            for monitor_id, config in self.monitor_configs.items():
                if monitor_id in self.monitor_widgets:
                    self.monitor_widgets[monitor_id]['path_label'].setText(
                        config.get('video_path', 'No video selected')
                    )
                    self.monitor_widgets[monitor_id]['fps_combo'].setCurrentText(
                        str(config.get('fps', 30))
                    )
                    self.monitor_widgets[monitor_id]['monitor_info'] = config['monitor_info']

    def save_config(self):
        config_data = {
            monitor_id: {
                'video_path': config['video_path'],
                'monitor_info': config['monitor_info'],
                'fps': config.get('fps', 30)
            }
            for monitor_id, config in self.monitor_configs.items()
        }
        
        with open(self.config_file, 'w') as f:
            json.dump(config_data, f, indent=4)
    
    def start_wallpapers(self):
        if self.video_player and self.video_player.isRunning():
            self.stop_wallpapers()
        
        self.video_player = VideoPlayerThread(self.monitor_configs)
        self.video_player.error_occurred.connect(self.handle_error)
        self.video_player.fps_updated.connect(self.update_live_fps)
        self.video_player.start()
    
    def stop_wallpapers(self):
        if self.video_player:
            self.video_player.stop()
            self.video_player = None
    
    def handle_error(self, error_msg):
        print(f"Error: {error_msg}")
    
    def update_live_fps(self, fps):
        for monitor_id, monitor_widget in self.monitor_widgets.items():
            monitor_widget['fps_label'].setText(f"Live FPS: {fps:.2f}")
    
    def closeEvent(self, event):
        self.stop_wallpapers()
        event.accept()

def main():
    app = QApplication(sys.argv)
    window = LiveWallpaperManager()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
