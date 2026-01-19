import tkinter as tk
from tkinter import ttk, messagebox
from tkinter import font as tkfont
import threading, time, sys, os, ctypes, json
from loguru import logger
from SimConnect import AircraftRequests, SimConnect, AircraftEvents

# Config file path
def get_config_path():
    """Get config file path in user's APPDATA"""
    appdata = os.environ.get("APPDATA", os.path.expanduser("~"))
    config_dir = os.path.join(appdata, "MSFS-SimRateMonitor")
    os.makedirs(config_dir, exist_ok=True)
    return os.path.join(config_dir, "config.json")

def load_config():
    """Load config from file"""
    try:
        config_path = get_config_path()
        if os.path.exists(config_path):
            with open(config_path, "r") as f:
                return json.load(f)
    except Exception as e:
        logger.debug(f"Could not load config: {e}")
    return {}

def save_config(config):
    """Save config to file"""
    try:
        config_path = get_config_path()
        with open(config_path, "w") as f:
            json.dump(config, f, indent=2)
    except Exception as e:
        logger.debug(f"Could not save config: {e}")

# Load custom font for overlay
def load_custom_font():
    """Load JetBrains Mono font for overlay display"""
    try:
        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
        
        font_path = os.path.join(base_path, "fonts", "JetBrainsMono-Bold.ttf")
        if os.path.exists(font_path):
            # Windows: Add font resource temporarily
            ctypes.windll.gdi32.AddFontResourceExW(font_path, 0x10, 0)  # FR_PRIVATE
            logger.info(f"Loaded custom font: {font_path}")
            return True
    except Exception as e:
        logger.debug(f"Could not load custom font: {e}")
    return False

CUSTOM_FONT_LOADED = load_custom_font()
OVERLAY_FONT = "JetBrains Mono" if CUSTOM_FONT_LOADED else "Consolas"


# Configure loguru for detailed logging
logger.remove()  # Remove default handler

# 只在有 stdout 时才添加控制台日志
if sys.stdout is not None:
    logger.add(sys.stdout, 
              format="<green>{time:HH:mm:ss}</green> | <level>{level: <8}</level> | <cyan>{name}</cyan>:<cyan>{function}</cyan>:<cyan>{line}</cyan> - <level>{message}</level>",
              level="DEBUG")

    # 总是添加文件日志
    logger.add("msfs_mini_gui.log", 
            format="{time:YYYY-MM-DD HH:mm:ss} | {level: <8} | {name}:{function}:{line} - {message}",
            level="DEBUG", 
            rotation="1 MB", 
            retention=1)       # 只保留 1 个文件（旧的自动删除）

# =============== 主题（非夜间） ===============
THEMES = {
    "modern_light_gray": {
        "bg": "#F7F7F7",
        "fg": "#2E2E2E",
        "accent": "#357C55",
        "control_bg": "#EDEDED",
        "divider": "#DFDFDF",
    },
    "sky_blue": {
        "bg": "#F5F9FF", "fg": "#333333", "accent": "#007ACC",
        "control_bg": "#E0ECF8", "divider": "#D5E2EE",
    },
    "ivory_minimal": {
        "bg": "#FAF9F6", "fg": "#4B4B4B", "accent": "#C27C0E",
        "control_bg": "#EDE9E3", "divider": "#E2DDD5",
    },
    "mint_fresh": {
        "bg": "#F2FFF8", "fg": "#444444", "accent": "#2BAE66",
        "control_bg": "#DFF5E1", "divider": "#D0E8D2",
    },
    "morning_orange": {
        "bg": "#FFF8F2", "fg": "#3D3D3D", "accent": "#FF6F00",
        "control_bg": "#FFE0CC", "divider": "#FFD0B2",
    },
}
CURRENT_THEME_NAME = "modern_light_gray"
CURRENT_THEME = THEMES[CURRENT_THEME_NAME]

# ---- helpers: 颜色混合，做选中/悬停底色 ----
def _hex_to_rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))

def _rgb_to_hex(rgb):
    return "#{:02X}{:02X}{:02X}".format(*rgb)

def _blend(c1, c2, t=0.85):
    r1,g1,b1 = _hex_to_rgb(c1); r2,g2,b2 = _hex_to_rgb(c2)
    r = int(r1*(1-t) + r2*t); g = int(g1*(1-t) + g2*t); b = int(b1*(1-t) + b2*t)
    return _rgb_to_hex((r,g,b))

class SegmentedRadio(tk.Frame):
    """更现代的分段单选控件：选中高亮、悬停变浅、支持主题色"""
    def __init__(self, parent, variable, options, theme, command=None, min_item_width=36, **kw):
        """
        options: [(value, text), ...]
        variable: tk.StringVar
        theme: dict，使用你的 CURRENT_THEME
        command: 回调（选项改变时）
        """
        super().__init__(parent, bg=theme["control_bg"], **kw)
        self.var = variable
        self.options = options
        self.t = theme
        self.command = command
        self._labels = []

        # 颜色体系（在 accent 与 control_bg 之间做混色）
        self.sel_bg   = _blend(self.t["accent"], self.t["control_bg"], 0.82)  # 选中底
        self.sel_fg   = self.t["accent"]                                      # 选中文本
        self.hov_bg   = _blend(self.t["accent"], self.t["control_bg"], 0.90)  # 悬停底
        self.nor_bg   = self.t["control_bg"]                                  # 常态底
        self.nor_fg   = self.t["fg"]                                          # 常态文本
        self.border   = self.t["divider"]

        # 外层边框（形成胶囊容器）
        self.container = tk.Frame(self, bg=self.t["control_bg"], highlightthickness=1,
                                  highlightbackground=self.border, bd=0)
        self.container.pack(fill="x")

        # 生成标签项
        cols = len(self.options)
        for i, (value, text) in enumerate(self.options):
            self.container.grid_columnconfigure(i, weight=1, uniform="seg")
            cell = tk.Frame(self.container, bg=self.nor_bg)
            cell.grid(row=0, column=i, sticky="nsew")

            lbl = tk.Label(cell, text=text, bg=self.nor_bg, fg=self.nor_fg,
                           font=("Segoe UI", 10, "bold" if self.var.get()==value else "normal"),
                           padx=12, pady=6, cursor="hand2")
            lbl.pack(fill="both", expand=True)
            lbl.bind("<Button-1>", lambda e, v=value: self._on_click(v))
            lbl.bind("<Enter>", lambda e, L=lbl: self._on_hover(L, True))
            lbl.bind("<Leave>", lambda e, L=lbl: self._on_hover(L, False))

            # 让每项最小宽度更协调
            lbl.update_idletasks()
            if lbl.winfo_width() < min_item_width:
                lbl.config(width=int(min_item_width/7))

            self._labels.append((value, lbl))

        # 变量联动
        self.var.trace_add("write", lambda *_: self._sync())
        self._sync()

    def _on_click(self, value):
        if self.var.get() != value:
            self.var.set(value)
        if self.command:
            self.command()

    def _on_hover(self, lbl, entering):
        value = None
        for v, l in self._labels:
            if l is lbl:
                value = v
                break
        if value is None:
            return
        if self.var.get() == value:
            # 已选中：悬停不改变（保持稳重）
            return
        lbl.configure(bg=self.hov_bg)

        if not entering:
            # 离开还原
            lbl.configure(bg=self.nor_bg)

    def _sync(self):
        cur = self.var.get()
        for v, lbl in self._labels:
            if v == cur:
                lbl.configure(bg=self.sel_bg, fg=self.sel_fg, font=("Segoe UI", 10, "bold"))
            else:
                lbl.configure(bg=self.nor_bg, fg=self.nor_fg, font=("Segoe UI", 10, "normal"))


class SimRateMonitor:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("MSFS Sim Rate Monitor")
        # Increase height to ensure all options are visible (from 230 to 320)
        self.root.geometry("340x330")
        self.root.resizable(False, False)

        try:
            # PyInstaller 资源路径处理
            if getattr(sys, 'frozen', False):
                # 打包后的环境
                base_path = sys._MEIPASS
            else:
                # 开发环境
                base_path = os.path.dirname(os.path.abspath(__file__))
            
            icon_path = os.path.join(base_path, "mini_gui_icon.ico")
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except (tk.TclError, Exception):
            pass

        self.sim_rate = tk.StringVar(value="-- x")
        self.overlay_size = tk.StringVar(value="l")  # Pain point 2: Default to Large
        self.auto_hide = tk.BooleanVar(value=True)    # Pain point 3: Hide when switched out
        self.sim_connect = None
        self.aircraft_requests = None
        self.aircraft_events = None
        self.connected = False
        self.running = True
        self.overlay_window = None
        self.overlay_hidden = False  # Track if overlay is hidden (not destroyed)
        
        # Load saved position from config
        config = load_config()
        saved_pos = config.get("overlay_position", [100, 100])
        self.overlay_position = tuple(saved_pos)
        
        self.is_msfs_active = False  # FIX: Initialize to False, not True
        self.startup_var = tk.BooleanVar(value=self._check_startup_exists())
        self._last_overlay_visible = False  # Track overlay visibility state
        self._hide_timer_id = None  # For debouncing hide

        self.size_configs = {
            "s": {"width": 80, "height": 25, "font_size": 10},
            "m": {"width": 120, "height": 40, "font_size": 16},
            "l": {"width": 160, "height": 55, "font_size": 20},
            "xl": {"width": 200, "height": 70, "font_size": 24},
            "xxl": {"width": 250, "height": 85, "font_size": 28},
            "hide": None,
        }

        # 存放 tk.Radiobutton
        self._size_radios = {}

        self.setup_ui()
        self.start_simconnect_thread()

    def setup_ui(self):
        t = CURRENT_THEME

        # ttk 只负责“外围”，卡片内部改用 tk 控件保证背景完全一致
        style = ttk.Style()
        try: style.theme_use("clam")
        except Exception: pass
        style.configure("App.TFrame", background=t["bg"])
        style.configure("App.Title.TLabel", background=t["bg"], foreground=t["accent"],
                        font=("Segoe UI", 28, "bold"))

        self.root.configure(bg=t["bg"])

        # 主容器
        main = ttk.Frame(self.root, padding=16, style="App.TFrame")
        main.grid(row=0, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main.columnconfigure(0, weight=1)

        # 大号速率显示
        rate_label = ttk.Label(main, textvariable=self.sim_rate, style="App.Title.TLabel")
        rate_label.grid(row=0, column=0, pady=(6, 10), sticky="n")

        # 分隔线
        divider = tk.Frame(main, height=1, bg=t["divider"], bd=0, highlightthickness=0)
        divider.grid(row=1, column=0, sticky="ew", pady=(4, 10))

        # 控制卡片（纯 tk，确保所有文字背景=control_bg）
        card = tk.Frame(main, bg=t["control_bg"], bd=0, highlightthickness=0)
        card.grid(row=2, column=0, sticky="ew")
        card.grid_columnconfigure(0, weight=1)
        # 内边距
        card_pad = tk.Frame(card, bg=t["control_bg"], padx=12, pady=10)
        card_pad.grid(row=0, column=0, sticky="ew")

        # Section 标题（bg=control_bg）
        title = tk.Label(card_pad, text="Overlay Size",
                         bg=t["control_bg"], fg=t["fg"], font=("Segoe UI", 10))
        title.grid(row=0, column=0, sticky="w")

        # 分段单选区域（现代风）
        size_frame = tk.Frame(card_pad, bg=t["control_bg"])
        size_frame.grid(row=1, column=0, sticky="ew", pady=(8, 0))

        size_options = [
            ("hide", "Hide"), ("s", "S"), ("m", "M"),
            ("l", "L"), ("xl", "XL")
        ]
        self.size_seg = SegmentedRadio(
            size_frame, self.overlay_size, size_options,
            theme=t, command=self.on_size_change
        )
        self.size_seg.pack(side="left", fill="x", expand=True)

        # --- Section 2: Options ---
        options_frame = tk.Frame(main, bg=t["bg"])
        options_frame.grid(row=3, column=0, sticky="ew", pady=(15, 0))
        
        # Style for checkboxes
        cb_style = {"bg": t["bg"], "fg": t["fg"], "activebackground": t["bg"], 
                    "activeforeground": t["accent"], "selectcolor": t["bg"],
                    "font": ("Segoe UI", 9)}

        # Auto-hide checkbox
        self.cb_hide = tk.Checkbutton(options_frame, text="Auto-hide when MSFS not in focus",
                                      variable=self.auto_hide, **cb_style)
        self.cb_hide.pack(anchor="w", pady=2)

        # Startup shortcut checkbox
        self.cb_startup = tk.Checkbutton(options_frame, text="Start with Windows (Startup Folder)",
                                         variable=self.startup_var, **cb_style,
                                         command=self.toggle_startup)
        self.cb_startup.pack(anchor="w", pady=2)

    # 选中态：粗体 + accent，未选：常规 + fg
    def _refresh_size_radio_styles(self):
        t = CURRENT_THEME
        selected = self.overlay_size.get()
        for value, rb in self._size_radios.items():
            if value == selected:
                rb.configure(fg=t["accent"], activeforeground=t["accent"],
                             font=("Segoe UI", 10, "bold"))
            else:
                rb.configure(fg=t["fg"], activeforeground=t["fg"],
                             font=("Segoe UI", 10))

    # ===== SimConnect 线程逻辑保持不变 =====
    def start_simconnect_thread(self):
        self.simconnect_thread = threading.Thread(target=self.simconnect_worker, daemon=True)
        self.simconnect_thread.start()

    def simconnect_worker(self):
        logger.info("Starting SimConnect worker thread")
        while self.running:
            try:
                if not self.connected:
                    self.connect_to_msfs()
                    if not self.connected:
                        logger.debug("Connection failed, waiting 3 seconds before retry...")
                        # Also check window even when not connected
                        self._check_active_window()
                    time.sleep(3.0)
                else:
                    self.update_sim_rate()
                    self._check_active_window()
                    time.sleep(0.5)
            except Exception as e:
                logger.error(f"SimConnect worker thread error: {type(e).__name__}: {e}")
                self.handle_disconnect()
                time.sleep(3.0)
        logger.info("SimConnect worker thread stopped")

    def _check_active_window(self):
        """Detect active window to determine if overlay should show"""
        try:
            # Using ctypes to get foreground window title
            hwnd = ctypes.windll.user32.GetForegroundWindow()
            length = ctypes.windll.user32.GetWindowTextLengthW(hwnd)
            buff = ctypes.create_unicode_buffer(length + 1)
            ctypes.windll.user32.GetWindowTextW(hwnd, buff, length + 1)
            title = buff.value
            
            # Check if MSFS or our main app is active
            is_msfs = "Microsoft Flight Simulator" in title or "FlightSimulator" in title
            is_our_app = "MSFS Sim Rate Monitor" in title
            
            # For overlay: check if title is empty (overlay has no title)
            # This is a heuristic - empty title + overlay exists = likely our overlay
            is_overlay = (title == "" and self.overlay_window and not self.overlay_hidden)
            
            new_active = is_msfs or is_our_app or is_overlay
            
            if new_active != self.is_msfs_active:
                self.is_msfs_active = new_active
                logger.debug(f"Focus changed: active = {new_active} (title: '{title[:50]}')") 
                
                if new_active:
                    # Cancel any pending hide and show immediately
                    if self._hide_timer_id:
                        self.root.after_cancel(self._hide_timer_id)
                        self._hide_timer_id = None
                    self.root.after(0, self._show_overlay)
                else:
                    # Debounce hide - delay by 300ms to allow focus to return
                    if not self._hide_timer_id:
                        self._hide_timer_id = self.root.after(300, self._hide_overlay_debounced)
        except Exception as e:
            logger.debug(f"Error checking active window: {e}")
    
    def _show_overlay(self):
        """Show overlay if conditions are met"""
        if self.overlay_size.get() == "hide":
            return
        if not self.connected:
            return
        
        if not self.overlay_window:
            self.create_overlay()
            rate = self.sim_rate.get()
            if rate != "-- x":
                self._update_overlay_rate(rate)
        elif self.overlay_hidden:
            self.overlay_window.deiconify()
            self.overlay_hidden = False
    
    def _hide_overlay_debounced(self):
        """Hide overlay after debounce delay"""
        self._hide_timer_id = None
        
        # Double check that we should still hide
        if self.is_msfs_active or not self.auto_hide.get():
            return  # Don't hide
        
        if self.overlay_window and not self.overlay_hidden:
            self._save_overlay_position()
            self.overlay_window.withdraw()
            self.overlay_hidden = True

    def _update_overlay_visibility(self):
        """Update overlay visibility based on connection and focus state"""
        if self.overlay_size.get() == "hide":
            if self.overlay_window:
                self._save_overlay_position()
                self.destroy_overlay()
            return
        
        # This method is now mainly used for size changes
        # Show/hide is handled by _show_overlay and _hide_overlay_debounced
        should_show = self.connected and (self.is_msfs_active or not self.auto_hide.get())
        
        if should_show:
            self._show_overlay()
        elif self.overlay_window and not self.overlay_hidden:
            self._save_overlay_position()
            self.overlay_window.withdraw()
            self.overlay_hidden = True
    
    def _save_overlay_position(self):
        """Save current overlay position to memory and config file"""
        if self.overlay_window:
            try:
                x = self.overlay_window.winfo_x()
                y = self.overlay_window.winfo_y()
                if x >= 0 and y >= 0:  # Valid position
                    self.overlay_position = (x, y)
                    config = load_config()
                    config["overlay_position"] = [x, y]
                    save_config(config)
            except:
                pass


    def connect_to_msfs(self):
        logger.info("Attempting to connect to MSFS...")
        try:
            # Clean up existing connections
            if self.aircraft_requests:
                logger.debug("Cleaning up existing AircraftRequests instance")
                self.aircraft_requests = None
            if self.sim_connect:
                logger.debug("Cleaning up existing SimConnect instance")
                try: 
                    self.sim_connect.exit()
                except Exception as e: 
                    logger.debug(f"Error cleaning up existing SimConnect: {e}")
                self.sim_connect = None

            logger.debug("Creating new SimConnect instance")
            self.sim_connect = SimConnect(auto_connect=False)
            
            logger.debug("Calling SimConnect.connect()")
            self.sim_connect.connect()
            logger.debug("SimConnect.connect() completed successfully")

            logger.debug("Creating AircraftRequests instance with SimConnect")
            self.aircraft_requests = AircraftRequests(self.sim_connect)
            logger.debug("AircraftRequests instance created successfully")

            logger.debug("Creating AircraftEvents instance with SimConnect")
            self.aircraft_events = AircraftEvents(self.sim_connect)
            logger.debug("AircraftEvents instance created successfully")
            
            # Test the connection by trying to get simulation rate
            logger.debug("Testing connection by requesting simulation rate")
            test_rate = self.aircraft_requests.get("SIMULATION_RATE")
            logger.debug(f"Test request result: {test_rate}")
            
            self.connected = True
            logger.success("✅ Successfully connected to MSFS via AircraftRequests")
            # Show overlay immediately on connection (don't wait for focus check)
            self.is_msfs_active = True  # Assume main window is active since user just opened app
            self.root.after(0, self._show_overlay)
        except Exception as e:
            logger.error(f"❌ Failed to connect to MSFS: {type(e).__name__}: {e}")
            logger.debug(f"Connection error details:", exc_info=True)
            self.handle_disconnect()

    def update_sim_rate(self):
        try:
            if not self.aircraft_requests:
                self.handle_disconnect(); return
            
            rate = self.aircraft_requests.get("SIMULATION_RATE")
            if rate is not None and isinstance(rate, (int, float)) and rate >= 0:
                rate_text = f"{rate:.2f}x"
                
                # 使用 after_idle 而不是 after(0) 来提高响应性
                self.root.after_idle(lambda: self._update_ui_rate(rate_text))
        except Exception as e:
            logger.debug(f"Update sim rate error: {e}")
            self.handle_disconnect()
    
    def _update_ui_rate(self, rate_text):
        """在主线程中更新UI"""
        try:
            self.sim_rate.set(rate_text)
            if self.overlay_window and self.overlay_size.get() != "hide":
                self.update_overlay(rate_text)
        except Exception as e:
            logger.debug(f"UI update error: {e}")

    def handle_disconnect(self):
        logger.warning("Handling MSFS disconnection")
        self.connected = False
        if self.aircraft_requests:
            self.aircraft_requests = None
        if self.aircraft_events:
            self.aircraft_events = None
        if self.sim_connect:
            try: 
                self.sim_connect.exit()
                logger.debug("SimConnect instance cleaned up")
            except Exception as e:
                logger.debug(f"Error during SimConnect cleanup: {e}")
            self.sim_connect = None
        self.root.after(0, lambda: self.sim_rate.set("-- x"))
        if self.overlay_window and self.overlay_size.get() != "hide":
            self.root.after(0, lambda: self.update_overlay("-- x"))

    def on_size_change(self):
        self._refresh_size_radio_styles()
        selected_size = self.overlay_size.get()
        if selected_size == "hide":
            self.destroy_overlay()
        else:
            # Destroy and recreate for size change
            if self.overlay_window:
                self.destroy_overlay()
            # Only show if conditions are met
            self._update_overlay_visibility()

    def sim_rate_incr(self):
        """Pain point 4: Increase sim rate"""
        if self.aircraft_events:
            logger.info("Sending SIM_RATE_INCR")
            # This triggers the event via AircraftEvents helper
            # We use the event name directly as it looks up in EventList
            try:
                event = self.aircraft_events.find("SIM_RATE_INCR")
                if event: event()
                else: logger.warning("SIM_RATE_INCR event not found")
            except Exception as e:
                logger.error(f"Error sending SIM_RATE_INCR: {e}")

    def sim_rate_decr(self):
        """Pain point 4: Decrease sim rate"""
        if self.aircraft_events:
            logger.info("Sending SIM_RATE_DECR")
            try:
                event = self.aircraft_events.find("SIM_RATE_DECR")
                if event: event()
                else: logger.warning("SIM_RATE_DECR event not found")
            except Exception as e:
                logger.error(f"Error sending SIM_RATE_DECR: {e}")

    # ===== Modern Overlay Design v3 (Speed-based colors) =====
    def create_overlay(self):
        if self.overlay_window or self.overlay_size.get() == "hide": return
        size_config = self.size_configs[self.overlay_size.get()]
        if size_config is None: return
        
        current_size = self.overlay_size.get()
        base_width, base_height, font_size = size_config["width"], size_config["height"], size_config["font_size"]
        
        # Adjust dimensions based on size and whether buttons are shown
        show_buttons = current_size not in ["s"]  # No buttons for S size
        width = base_width + (45 if show_buttons else 0)
        height = base_height

        self.overlay_window = tk.Toplevel()
        self.overlay_window.title("")
        # Use saved position
        x, y = self.overlay_position
        self.overlay_window.geometry(f"{width}x{height}+{x}+{y}")
        self.overlay_window.resizable(False, False)
        self.overlay_window.overrideredirect(True)
        self.overlay_window.wm_attributes("-topmost", True)
        self.overlay_window.wm_attributes("-alpha", 0.95)
        self.overlay_hidden = False  # Reset hidden flag
        
        # Modern dark theme colors with speed-based coloring
        bg_color = "#0D1117"
        border_color = "#30363D"
        btn_bg = "#21262D"
        btn_hover = "#30363D"
        
        # Speed-based text colors
        self._color_normal = "#E6EDF3"   # White - normal speed (1x)
        self._color_slow = "#79C0FF"     # Cyan - slower than normal (0.25, 0.5)
        self._color_fast = "#FFA657"     # Orange - faster than normal (2, 4, 8, 16)
        self._color_btn = "#58A6FF"      # Blue for buttons
        
        self.overlay_window.configure(bg=bg_color)
        
        # Main container
        self.overlay_container = tk.Frame(
            self.overlay_window, 
            bg=bg_color,
            highlightthickness=2,
            highlightbackground=border_color
        )
        self.overlay_container.pack(expand=True, fill="both")
        
        # Horizontal content layout
        content = tk.Frame(self.overlay_container, bg=bg_color)
        content.pack(expand=True, fill="both", padx=6, pady=2)
        
        # Rate display using JetBrains Mono - color changes based on speed
        self.overlay_label = tk.Label(
            content, 
            text="1x",
            font=(OVERLAY_FONT, font_size, "bold"),
            fg=self._color_normal, 
            bg=bg_color
        )
        self.overlay_label.pack(side="left", expand=True)
        
        # Control buttons (horizontal layout with JetBrains Mono)
        if show_buttons:
            btn_frame = tk.Frame(content, bg=bg_color)
            btn_frame.pack(side="right", padx=(4, 2))
            
            # Button font - use JetBrains Mono for consistency
            btn_font = (OVERLAY_FONT, max(10, font_size // 2))
            
            # Minus button (left arrow)
            self.btn_minus = tk.Label(
                btn_frame, text="<", 
                font=btn_font, fg=self._color_btn, bg=btn_bg,
                padx=3, pady=0, cursor="hand2"
            )
            self.btn_minus.pack(side="left", padx=1)
            self.btn_minus.bind("<Button-1>", lambda e: self.sim_rate_decr())
            self.btn_minus.bind("<Enter>", lambda e: self.btn_minus.configure(bg=btn_hover))
            self.btn_minus.bind("<Leave>", lambda e: self.btn_minus.configure(bg=btn_bg))
            
            # Plus button (right arrow)
            self.btn_plus = tk.Label(
                btn_frame, text=">", 
                font=btn_font, fg=self._color_btn, bg=btn_bg,
                padx=3, pady=0, cursor="hand2"
            )
            self.btn_plus.pack(side="left", padx=1)
            self.btn_plus.bind("<Button-1>", lambda e: self.sim_rate_incr())
            self.btn_plus.bind("<Enter>", lambda e: self.btn_plus.configure(bg=btn_hover))
            self.btn_plus.bind("<Leave>", lambda e: self.btn_plus.configure(bg=btn_bg))
        
        # Bind drag events
        for widget in [self.overlay_label, content, self.overlay_container]:
            widget.bind("<Button-1>", self.start_move)
            widget.bind("<B1-Motion>", self.do_move)
        
        self._overlay_bg = bg_color

    def _update_overlay_rate(self, rate_text):
        """Update overlay text and color based on speed"""
        if hasattr(self, "overlay_label") and self.overlay_label:
            self.overlay_label.config(text=rate_text)
            
            # Determine color based on rate value
            try:
                # Extract numeric value from rate_text (e.g., "2.00x" -> 2.0)
                rate_val = float(rate_text.replace("x", "").strip())
                
                if hasattr(self, "_color_normal"):
                    if rate_val < 1.0:
                        color = self._color_slow   # Cyan for slow
                    elif rate_val > 1.0:
                        color = self._color_fast   # Orange for fast
                    else:
                        color = self._color_normal  # White for normal
                    
                    self.overlay_label.config(fg=color)
            except:
                pass  # Keep current color if parsing fails

    # Startup Folder logic
    def _get_startup_path(self):
        return os.path.join(os.environ["APPDATA"], r"Microsoft\Windows\Start Menu\Programs\Startup")

    def _get_shortcut_path(self):
        return os.path.join(self._get_startup_path(), "MSFS-SimRateMonitor.lnk")

    def _check_startup_exists(self):
        return os.path.exists(self._get_shortcut_path())

    def toggle_startup(self):
        lnk = self._get_shortcut_path()
        if self.startup_var.get():
            # Create shortcut
            try:
                import winshell
                from win32com.client import Dispatch
                
                target = sys.executable if not getattr(sys, 'frozen', False) else sys.executable
                arguments = f'"{os.path.abspath(__file__)}"' if not getattr(sys, 'frozen', False) else ""
                if "--startup" not in arguments and not getattr(sys, 'frozen', False):
                    arguments += " --startup"
                elif getattr(sys, 'frozen', False):
                    arguments = "--startup"
                
                shell = Dispatch('WScript.Shell')
                shortcut = shell.CreateShortCut(lnk)
                shortcut.Targetpath = target
                shortcut.Arguments = arguments
                shortcut.WorkingDirectory = os.path.dirname(target)
                shortcut.IconLocation = target
                shortcut.save()
                logger.info(f"Startup shortcut created: {lnk}")
            except ImportError:
                # Fallback to a simple powershell command to create shortcut if win32com missing
                target = sys.executable if not getattr(sys, 'frozen', False) else sys.executable
                args = f'\\"{os.path.abspath(__file__)}\\" --startup' if not getattr(sys, 'frozen', False) else "--startup"
                ps_cmd = f'$s=(New-Object -ComObject WScript.Shell).CreateShortcut("{lnk}");$s.TargetPath="{target}";$s.Arguments="{args}";$s.Save()'
                os.system(f'powershell -Command {ps_cmd}')
                logger.info("Startup shortcut created via PowerShell")
            except Exception as e:
                logger.error(f"Failed to create startup shortcut: {e}")
                messagebox.showerror("Error", f"Failed to create startup shortcut:\n{e}")
                self.startup_var.set(False)
        else:
            # Remove shortcut
            if os.path.exists(lnk):
                try:
                    os.remove(lnk)
                    logger.info(f"Startup shortcut removed: {lnk}")
                except Exception as e:
                    logger.error(f"Failed to remove startup shortcut: {e}")
                    messagebox.showerror("Error", f"Failed to remove shortcut:\n{e}")
                    self.startup_var.set(True)

    def start_move(self, event):
        self.x, self.y = event.x, event.y

    def do_move(self, event):
        x = self.overlay_window.winfo_x() + (event.x - self.x)
        y = self.overlay_window.winfo_y() + (event.y - self.y)
        self.overlay_window.geometry(f"+{x}+{y}")

    def update_overlay(self, rate_text):
        try:
            if self.overlay_window and hasattr(self, "overlay_label"):
                self.overlay_label.config(text=rate_text)
        except Exception as e:
            logger.debug(f"Error updating overlay: {e}")
            if self.overlay_size.get() != "hide":
                self.create_overlay()

    def destroy_overlay(self):
        if self.overlay_window:
            self.overlay_window.destroy()
            self.overlay_window = None

    def on_closing(self):
        self.running = False
        
        # 关闭连接
        if self.aircraft_requests:
            self.aircraft_requests = None
        if self.sim_connect:
            try:
                self.sim_connect.exit()
            except: pass
            self.sim_connect = None
        
        try: self.destroy_overlay()
        except: pass
        
        # 快速退出，不等待线程
        self.root.quit()
        self.root.destroy()

    def run(self):
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.mainloop()


if __name__ == "__main__":
    logger.info("🚀 Starting MSFS Simulation Rate Monitor")
    
    # Handle --startup argument
    is_startup = "--startup" in sys.argv
    if is_startup:
        logger.info("Starting in silent mode (startup)")
    
    app = SimRateMonitor()
    
    if is_startup:
        # Hide main window if started via startup
        app.root.withdraw()
    
    # Don't force overlay on startup - let the worker thread handle it based on connection + focus
    app.run()
