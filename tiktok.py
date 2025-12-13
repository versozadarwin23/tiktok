import tkinter as tk
from tkinter import messagebox, filedialog
from PIL import Image, ImageTk
import subprocess
import threading
import multiprocessing
import time
import os
import io
import customtkinter as ctk
import queue
import random
import concurrent.futures
import requests
import json
from pathlib import Path
import re
import sys
import shutil
import tempfile
import hashlib
import uuid
import xml.etree.ElementTree as ET

# --- App Version and Update URL ---
__version__ = "1.4.0"  # Updated version number
UPDATE_URL = "https://raw.githubusercontent.com/versozadarwin23/adbtool/refs/heads/main/main.py"
VERSION_CHECK_URL = "https://raw.githubusercontent.com/versozadarwin23/adbtool/refs/heads/main/version.txt"

# --- GLOBAL CONSTANTS for TikTok Lite (UPDATED) ---
TIKTOK_LITE_PACKAGE = "com.zhiliaoapp.musically.go"
TIKTOK_LITE_ACTIVITY = "com.ss.android.ugc.aweme.main.homepage.MainActivity"

# --- Global Flag for Stopping Commands ---
is_stop_requested = threading.Event()


def run_adb_command(command, serial):
    """
    Executes a single ADB command for a specific device with a timeout, checking for a stop signal.

    Returns: (bool success, str output_or_error)
    """
    if is_stop_requested.is_set():
        return False, "Stop requested."

    try:
        process = subprocess.Popen(['adb', '-s', serial] + command, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

        timeout_seconds = 60
        start_time = time.time()
        while process.poll() is None and (time.time() - start_time < timeout_seconds):
            if is_stop_requested.is_set():
                process.terminate()
                return False, "Terminated due to stop request."

        if process.poll() is None:
            process.terminate()
            raise subprocess.TimeoutExpired(cmd=['adb', '-s', serial] + command, timeout=timeout_seconds)

        stdout, stderr = process.communicate()

        if process.returncode != 0:
            return False, stderr.decode()
        else:
            return True, stdout.decode()

    except subprocess.CalledProcessError as e:
        return False, e.stderr.decode()
    except FileNotFoundError:
        return False, "ADB not found. Please install it and add to PATH."
    except subprocess.TimeoutExpired:
        return False, "Command timed out."
    except Exception as e:
        return False, str(e)


def run_text_command(text_to_send, serial):
    """
    Sends a specific text string character-by-character with delay and proper space escaping.
    """
    if is_stop_requested.is_set():
        return

    if not text_to_send:
        return

    formatted_text = text_to_send
    DELAY_PER_CHAR = 0.02

    try:
        for char in formatted_text:
            if is_stop_requested.is_set():
                return

            adb_char = char.replace(' ', '%s')

            command_args = ['shell', 'input', 'text', adb_char]

            subprocess.run(['adb', '-s', serial] + command_args,
                           stdout=subprocess.DEVNULL,
                           stderr=subprocess.DEVNULL,
                           check=True,
                           timeout=5)

            time.sleep(DELAY_PER_CHAR)

    except Exception:
        pass


def create_and_run_updater_script(new_file_path, old_file_path):
    """Handles the file replacement and app restart for updates."""
    try:
        time.sleep(2)
        shutil.move(str(new_file_path), str(old_file_path))

        if sys.platform.startswith('win'):
            os.startfile(str(old_file_path))
        else:
            subprocess.Popen(['python3', str(old_file_path)])

        os._exit(0)
    except Exception as e:
        messagebox.showerror("Update Error", f"Failed to replace file: {e}")


# --- AdbControllerApp Class ---
class AdbControllerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # --- Configuration ---
        self.title(f"TikTok Lite ADB Commander V{__version__}")
        self.geometry("1400x900")
        self.state('zoomed')
        ctk.set_appearance_mode("Dark")
        ctk.set_default_color_theme("blue")

        # --- NEW TikTok Neon Dark Theme Palette ---
        self.COLOR_BACKGROUND = "#000000"
        self.COLOR_FRAME = "#1A1A1A"
        self.COLOR_BORDER = "#333333"
        self.COLOR_ACCENT = "#FF0050"  # Electric Pink/TikTok Red
        self.COLOR_ACCENT_HOVER = "#FF3377"
        self.COLOR_SUCCESS = "#00FFFF"  # Cyan
        self.COLOR_SUCCESS_HOVER = "#33FFFF"
        self.COLOR_DANGER = "#FF6600"  # Warning Orange
        self.COLOR_DANGER_HOVER = "#FF8533"
        self.COLOR_WARNING = "#FFFF00"
        self.COLOR_TEXT_PRIMARY = "#FFFFFF"
        self.COLOR_TEXT_SECONDARY = "#AAAAAA"

        # --- NEW Standardized Fonts (Larger and modern) ---
        self.FONT_TITLE = ctk.CTkFont(family="Consolas", size=36, weight="bold")
        self.FONT_HEADING = ctk.CTkFont(family="Consolas", size=20, weight="bold")
        self.FONT_SUBHEADING = ctk.CTkFont(family="Consolas", size=18, weight="bold")
        self.FONT_BODY = ctk.CTkFont(family="Consolas", size=16)
        self.FONT_BUTTON = ctk.CTkFont(family="Consolas", size=16, weight="bold")
        self.FONT_MONO = ctk.CTkFont(family="Consolas", size=16)
        self.FONT_STATUS = ctk.CTkFont(family="Consolas", size=14, weight="normal")

        # --- App State Variables ---
        self.device_frames = {}
        self.device_canvases = {}
        self.device_images = {}
        self.press_start_coords = {}
        self.press_time = {}
        self.selected_device_serial = None
        self.devices = []
        self.long_press_duration = 0.5
        self.drag_threshold = 20
        self.capture_running = {}
        self.screenshot_queue = queue.Queue()
        self.capture_thread = None
        self.update_image_id = None
        self.is_capturing = False
        self.apk_path = None
        self.is_muted = False
        self.update_check_job = None
        self.is_update_prompt_showing = False
        self.share_pairs = []
        self.share_pair_frame = None
        self.is_auto_typing = threading.Event()

        self.executor = concurrent.futures.ThreadPoolExecutor(max_workers=multiprocessing.cpu_count() * 4)

        # --- Main Window Grid Configuration ---
        self.grid_columnconfigure(0, weight=1, minsize=500)
        self.grid_columnconfigure(1, weight=3)
        self.grid_rowconfigure(0, weight=1)
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.configure(fg_color=self.COLOR_BACKGROUND)

        # --- [LEFT] Control Panel Setup ---
        self.control_panel = ctk.CTkFrame(self, corner_radius=0, fg_color=self.COLOR_FRAME)
        self.control_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 1), pady=0)

        self.control_panel.grid_columnconfigure(0, weight=1)
        self.control_panel.grid_rowconfigure(4, weight=1)
        self.control_panel.grid_rowconfigure(5, weight=0)

        # --- Row 0: Title ---
        ctk.CTkLabel(self.control_panel, text=f"TIKTOK LITE COMMANDER V{__version__}",
                     font=self.FONT_TITLE,
                     text_color=self.COLOR_ACCENT).grid(
            row=0, column=0, pady=(20, 10), padx=25, sticky='w')

        # --- Row 1: Global Stop Button ---
        self.stop_all_button = ctk.CTkButton(self.control_panel, text="🛑 TERMINATE ALL OPERATIONS 🛑",
                                             command=self.stop_all_commands,
                                             fg_color=self.COLOR_DANGER,
                                             hover_color=self.COLOR_DANGER_HOVER,
                                             text_color=self.COLOR_TEXT_PRIMARY,
                                             corner_radius=8,
                                             font=self.FONT_HEADING, height=60)
        self.stop_all_button.grid(row=1, column=0, sticky='ew', padx=25, pady=15)

        # --- Row 2: Device Management Frame (Simplified/Restructured) ---
        device_mgmt_frame = ctk.CTkFrame(self.control_panel, fg_color="transparent")
        device_mgmt_frame.grid(row=2, column=0, sticky="ew", padx=25, pady=(10, 5))
        device_mgmt_frame.grid_columnconfigure(0, weight=1)
        device_mgmt_frame.grid_columnconfigure(1, weight=1)
        device_mgmt_frame.grid_columnconfigure(2, weight=1)

        self.device_count_label = ctk.CTkLabel(device_mgmt_frame, text="DEVICES: 0",
                                               font=self.FONT_SUBHEADING, text_color=self.COLOR_TEXT_SECONDARY)
        self.device_count_label.grid(row=0, column=0, sticky='w', padx=(0, 10))

        self.detect_button = ctk.CTkButton(device_mgmt_frame, text="🔄 REFRESH", command=self.detect_devices,
                                           corner_radius=8,
                                           fg_color=self.COLOR_ACCENT,
                                           hover_color=self.COLOR_ACCENT_HOVER,
                                           font=self.FONT_BUTTON, height=45, text_color=self.COLOR_BACKGROUND)
        self.detect_button.grid(row=0, column=1, sticky='ew', padx=(5, 5))

        self.update_button = ctk.CTkButton(device_mgmt_frame, text=f"📥 UPDATE (V{__version__})",
                                           command=self.update_app,
                                           fg_color="transparent", hover_color=self.COLOR_BORDER, corner_radius=8,
                                           font=self.FONT_BUTTON,
                                           text_color=self.COLOR_SUCCESS, border_color=self.COLOR_SUCCESS,
                                           border_width=2,
                                           height=45)
        self.update_button.grid(row=0, column=2, sticky='e', padx=(5, 0))

        # --- Row 3: Device Selection ---
        device_select_frame = ctk.CTkFrame(self.control_panel, fg_color="transparent")
        device_select_frame.grid(row=3, column=0, sticky="ew", padx=25, pady=5)
        device_select_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(device_select_frame, text="LIVE VIEW:",
                     font=self.FONT_BUTTON, text_color=self.COLOR_TEXT_SECONDARY).grid(row=0, column=0, sticky='w')

        self.device_selector_var = ctk.StringVar(value="No devices found")
        self.device_option_menu = ctk.CTkOptionMenu(device_select_frame,
                                                    variable=self.device_selector_var,
                                                    command=self.on_device_select_menu,
                                                    values=["No devices found"],
                                                    state="disabled",
                                                    font=self.FONT_MONO,
                                                    dropdown_font=self.FONT_MONO,
                                                    fg_color=self.COLOR_FRAME,
                                                    button_color=self.COLOR_BORDER,
                                                    button_hover_color=self.COLOR_ACCENT,
                                                    dropdown_fg_color=self.COLOR_FRAME,
                                                    dropdown_hover_color=self.COLOR_BORDER,
                                                    corner_radius=8,
                                                    height=45)
        self.device_option_menu.grid(row=0, column=1, sticky='ew', padx=(15, 0))

        # --- Row 4: Tab View ---
        self.tab_view = ctk.CTkTabview(self.control_panel,
                                       fg_color=self.COLOR_FRAME,
                                       segmented_button_selected_color=self.COLOR_ACCENT,
                                       segmented_button_selected_hover_color=self.COLOR_ACCENT_HOVER,
                                       segmented_button_unselected_hover_color=self.COLOR_BORDER,
                                       segmented_button_unselected_color=self.COLOR_FRAME,
                                       text_color=self.COLOR_TEXT_PRIMARY,
                                       border_color=self.COLOR_BORDER,
                                       border_width=2,
                                       corner_radius=8)
        self.tab_view.grid(row=4, column=0, sticky="nsew", padx=25, pady=15)

        self.tab_view.add("TIKTOK AUTOMATION")
        self.tab_view.add("UTILITIES")
        self.tab_view.set("TIKTOK AUTOMATION")

        self._configure_tab_layouts()

        # --- Row 5: Status Bar ---
        self.status_label = ctk.CTkLabel(self.control_panel, text="Awaiting Command...", anchor='w',
                                         font=self.FONT_STATUS, text_color=self.COLOR_TEXT_SECONDARY, height=35,
                                         fg_color=self.COLOR_FRAME)
        self.status_label.grid(row=5, column=0, sticky='sew', padx=25, pady=(5, 15))

        # --- [RIGHT] Device View Panel Setup ---
        self.device_view_panel = ctk.CTkFrame(self, fg_color=self.COLOR_BACKGROUND, corner_radius=0)
        self.device_view_panel.grid(row=0, column=1, sticky="nsew", padx=0, pady=0)
        self.device_view_panel.grid_columnconfigure(0, weight=1)
        self.device_view_panel.grid_rowconfigure(0, weight=1)

        # --- Initial Setup ---
        self.detect_devices()
        self.check_for_updates()
        self.start_periodic_update_check()

    # --- Section Helper for Professional Look ---
    def _create_section_header(self, parent, text, row):
        """Creates a standardized, styled section header using Cyan for highlight."""
        ctk.CTkLabel(parent, text=f"• {text} •",
                     font=self.FONT_HEADING, text_color=self.COLOR_SUCCESS).grid(
            row=row, column=0, sticky='w', padx=15, pady=(20, 5))

    def _create_section_frame(self, parent, row):
        """Creates a standardized frame for grouping widgets."""
        frame = ctk.CTkFrame(parent, fg_color=self.COLOR_FRAME, corner_radius=12,
                             border_width=1, border_color=self.COLOR_BORDER)
        frame.grid(row=row, column=0, sticky='ew', padx=15, pady=5)
        frame.grid_columnconfigure(0, weight=1)
        return frame

    # --- Configuration Methods (Unchanged Logic) ---

    def start_periodic_update_check(self):
        self.update_check_job = self.after(60000, self._periodic_check_updates)

    def _periodic_check_updates(self):
        threading.Thread(target=self._check_and_reschedule, daemon=True).start()

    def _check_and_reschedule(self):
        try:
            response = requests.get(VERSION_CHECK_URL, timeout=10)
            response.raise_for_status()

            latest_version = response.text.strip()
            try:
                local_v = float(__version__)
                remote_v = float(latest_version)

                if remote_v > local_v:
                    self.after(0, self.ask_for_update, latest_version)
            except ValueError:
                if latest_version > __version__:
                    self.after(0, self.ask_for_update, latest_version)
        except requests.exceptions.RequestException:
            pass
        except Exception:
            pass
        finally:
            self.update_check_job = self.after(60000, self._periodic_check_updates)

    def check_for_updates(self):
        def _check_in_thread():
            try:
                response = requests.get(VERSION_CHECK_URL, timeout=10)
                response.raise_for_status()

                latest_version = response.text.strip()
                try:
                    local_v = float(__version__)
                    remote_v = float(latest_version)

                    if remote_v > local_v:
                        self.after(0, self.ask_for_update, latest_version)
                except ValueError:
                    if latest_version > __version__:
                        self.after(0, self.ask_for_update, latest_version)

            except requests.exceptions.HTTPError:
                self.after(0, lambda: self.status_label.configure(
                    text=f"❌ ERROR: Failed to check for update. HTTP Error.",
                    text_color=self.COLOR_DANGER))
            except requests.exceptions.RequestException:
                self.after(0, lambda: self.status_label.configure(
                    text="❌ ERROR: Failed to check for update. Connection Error.",
                    text_color=self.COLOR_DANGER))
            except Exception:
                self.after(0, lambda: self.status_label.configure(
                    text="❌ ERROR: An unexpected error occurred during version check.",
                    text_color=self.COLOR_DANGER))

        update_thread = threading.Thread(target=_check_in_thread, daemon=True)
        update_thread.start()

    def ask_for_update(self, latest_version):
        if self.is_update_prompt_showing:
            return

        try:
            self.is_update_prompt_showing = True
            title = "New TikTok Lite Commander Update!"
            # --- TRANSLATED TO ENGLISH ---
            message = (
                f"A new version ({latest_version}) is available!\n\n"
                "It features improved performance and new functionalities.\n\n"
                "The application will close and restart to apply the update. Update now?"
            )
            # --- END TRANSLATED ---

            response = messagebox.askyesno(title, message)
            if response:
                self.update_app()
        finally:
            self.is_update_prompt_showing = False

    def on_closing(self):
        if self.update_check_job:
            self.after_cancel(self.update_check_job)

        self.is_auto_typing.clear()
        is_stop_requested.set()

        self.stop_capture()
        self.executor.shutdown(wait=False)
        self.destroy()

    def _configure_tab_layouts(self):
        """
        Configures the grid layout for each tab using the Neon Dark design with DISTINCT sizing/layout.
        """

        # --- Configure "TIKTOK AUTOMATION" Tab ---
        tiktok_tab_container = self.tab_view.tab("TIKTOK AUTOMATION")

        tiktok_frame = ctk.CTkScrollableFrame(tiktok_tab_container, fg_color="transparent")
        tiktok_frame.pack(fill="both", expand=True, padx=0, pady=0)
        tiktok_frame.columnconfigure(0, weight=1)

        # --- Section: App Control (STACKED, LARGER BUTTONS) ---
        self._create_section_header(tiktok_frame, "TIKTOK APP CONTROL", 0)
        tiktok_app_frame = self._create_section_frame(tiktok_frame, 1)
        tiktok_app_frame.columnconfigure(0, weight=1)

        # Launch button uses Cyan (Success) and is full width
        self.launch_tiktok_lite_button = ctk.CTkButton(tiktok_app_frame, text="🚀 LAUNCH TIKTOK LITE",
                                                       command=self.launch_tiktok_lite,
                                                       corner_radius=10, fg_color=self.COLOR_SUCCESS,
                                                       hover_color=self.COLOR_SUCCESS_HOVER,
                                                       height=55, font=self.FONT_HEADING,
                                                       text_color=self.COLOR_BACKGROUND)
        self.launch_tiktok_lite_button.grid(row=0, column=0, sticky='ew', padx=15, pady=(15, 5))

        # Force Stop button uses Danger (Orange) and is full width
        self.force_stop_tiktok_lite_button = ctk.CTkButton(tiktok_app_frame, text="🛑 FORCE STOP TIKTOK",
                                                           command=self.force_stop_tiktok_lite,
                                                           fg_color=self.COLOR_DANGER,
                                                           hover_color=self.COLOR_DANGER_HOVER, corner_radius=10,
                                                           text_color=self.COLOR_TEXT_PRIMARY, height=55,
                                                           font=self.FONT_HEADING)
        self.force_stop_tiktok_lite_button.grid(row=1, column=0, sticky='ew', padx=15, pady=(5, 15))

        # --- Section: Single Video Visit (LARGER ENTRY) ---
        self._create_section_header(tiktok_frame, "SINGLE VIDEO VISIT", 2)
        tiktok_single_frame = self._create_section_frame(tiktok_frame, 3)

        self.tiktok_url_entry = ctk.CTkEntry(tiktok_single_frame, placeholder_text="Enter TikTok URL...", height=50,
                                             corner_radius=10, font=self.FONT_BODY,
                                             fg_color=self.COLOR_FRAME, border_color=self.COLOR_ACCENT, border_width=2)
        self.tiktok_url_entry.grid(row=0, column=0, sticky='ew', padx=15, pady=(15, 10))

        # Visit button uses Accent color (Pink)
        self.tiktok_button = ctk.CTkButton(tiktok_single_frame, text="▶️ VISIT VIDEO (DEEPLINK)",
                                           command=self.open_tiktok_lite_deeplink,
                                           fg_color=self.COLOR_ACCENT, hover_color=self.COLOR_ACCENT_HOVER, height=50,
                                           font=self.FONT_BUTTON, corner_radius=10, text_color=self.COLOR_BACKGROUND)
        self.tiktok_button.grid(row=1, column=0, sticky='ew', padx=15, pady=(10, 15))

        # --- Section: Multi-Video Automation ---
        self._create_section_header(tiktok_frame, "MULTI-LINK & COMMENT PAIRS", 4)

        # Container para sa mga dynamic na entry
        self.share_pair_frame = ctk.CTkScrollableFrame(tiktok_frame, fg_color=self.COLOR_FRAME, height=200,
                                                       corner_radius=12, border_color=self.COLOR_BORDER, border_width=1)
        self.share_pair_frame.grid(row=5, column=0, sticky='ew', padx=15, pady=5)
        self.share_pair_frame.columnconfigure(0, weight=1)

        # Add Link/Caption Button uses Cyan (Success) and is larger
        add_link_button = ctk.CTkButton(tiktok_frame, text="➕ ADD LINK / COMMENT PAIR", command=self.add_share_pair,
                                        fg_color=self.COLOR_SUCCESS, hover_color=self.COLOR_SUCCESS_HOVER, height=50,
                                        font=self.FONT_BUTTON, corner_radius=10,
                                        text_color=self.COLOR_BACKGROUND)
        add_link_button.grid(row=6, column=0, sticky='ew', padx=15, pady=(10, 15))

        # Initial pair upon startup
        self.add_share_pair(is_initial=True)

        # --- Section: Automation Actions ---
        self._create_section_header(tiktok_frame, "TEXT FILE TOOLS", 7)
        action_frame = self._create_section_frame(tiktok_frame, 8)
        action_frame.columnconfigure(0, weight=1)
        action_frame.columnconfigure(1, weight=1)

        self.send_button = ctk.CTkButton(action_frame, text="RANDOM TEXT ✉️",
                                         command=self.send_text_to_devices,
                                         fg_color=self.COLOR_BORDER, hover_color=self.COLOR_TEXT_SECONDARY, height=50,
                                         font=self.FONT_BUTTON, text_color=self.COLOR_TEXT_PRIMARY,
                                         corner_radius=10)
        self.send_button.grid(row=0, column=0, sticky='ew', padx=(15, 7), pady=15)

        self.remove_emoji_button = ctk.CTkButton(action_frame, text="REMOVE EMOJIS 🚫",
                                                 command=self.remove_emojis_from_file,
                                                 fg_color=self.COLOR_DANGER, hover_color=self.COLOR_DANGER_HOVER,
                                                 height=50,
                                                 font=self.FONT_BUTTON,
                                                 text_color=self.COLOR_TEXT_PRIMARY, corner_radius=10)
        self.remove_emoji_button.grid(row=0, column=1, sticky='ew', padx=(7, 15), pady=15)

        # --- Prominent "START AUTO-TYPE" Button (LARGER FONT) ---
        self.find_click_type_button = ctk.CTkButton(tiktok_frame, text="⚡ START AUTO-TYPE ⚡",
                                                    command=self.toggle_auto_type_loop,
                                                    fg_color=self.COLOR_ACCENT, hover_color=self.COLOR_ACCENT_HOVER,
                                                    height=70,
                                                    font=self.FONT_TITLE,
                                                    text_color=self.COLOR_BACKGROUND,
                                                    corner_radius=12)
        self.find_click_type_button.grid(row=9, column=0, sticky='ew', padx=15, pady=(20, 20))

        # --- Configure "UTILITIES" Tab ---
        utility_tab_container = self.tab_view.tab("UTILITIES")

        utility_frame = ctk.CTkScrollableFrame(utility_tab_container, fg_color="transparent")
        utility_frame.pack(fill="both", expand=True, padx=0, pady=0)
        utility_frame.columnconfigure(0, weight=1)

        # --- Section: App Management (Stacked) ---
        self._create_section_header(utility_frame, "APK MANAGEMENT", 0)
        apk_frame = self._create_section_frame(utility_frame, 1)

        self.apk_path_entry = ctk.CTkEntry(apk_frame, placeholder_text="Path: No APK selected...", height=45,
                                           corner_radius=10, font=self.FONT_BODY)
        self.apk_path_entry.grid(row=0, column=0, sticky='ew', padx=15, pady=(15, 5))

        apk_button_frame = ctk.CTkFrame(apk_frame, fg_color="transparent")
        apk_button_frame.grid(row=1, column=0, sticky='ew', padx=15, pady=(5, 15))
        apk_button_frame.columnconfigure(0, weight=1)
        apk_button_frame.columnconfigure(1, weight=1)

        browse_apk_button = ctk.CTkButton(apk_button_frame, text="BROWSE", command=self.browse_apk_file,
                                          fg_color=self.COLOR_BORDER, hover_color=self.COLOR_TEXT_SECONDARY,
                                          corner_radius=10, height=50,
                                          font=self.FONT_BUTTON)
        browse_apk_button.grid(row=0, column=0, sticky='ew', padx=(0, 7))

        install_apk_button = ctk.CTkButton(apk_button_frame, text="INSTALL APK ⬇️", command=self.install_apk_to_devices,
                                           fg_color=self.COLOR_ACCENT, hover_color=self.COLOR_ACCENT_HOVER,
                                           corner_radius=10,
                                           height=50,
                                           font=self.FONT_BUTTON, text_color=self.COLOR_BACKGROUND)
        install_apk_button.grid(row=0, column=1, sticky='ew', padx=(7, 0))

        # --- Section: Device Control ---
        self._create_section_header(utility_frame, "NETWORK/DEVICE CONTROL", 2)
        device_control_frame = self._create_section_frame(utility_frame, 3)
        device_control_frame.columnconfigure(0, weight=1)
        device_control_frame.columnconfigure(1, weight=1)

        enable_airplane_button = ctk.CTkButton(device_control_frame, text="ENABLE AIRPLANE ✈️",
                                               command=self.enable_airplane_mode,
                                               fg_color=self.COLOR_BORDER, hover_color=self.COLOR_TEXT_SECONDARY,
                                               corner_radius=10, height=50,
                                               font=self.FONT_BUTTON)
        enable_airplane_button.grid(row=0, column=0, sticky='ew', padx=(15, 7), pady=15)

        disable_airplane_button = ctk.CTkButton(device_control_frame, text="DISABLE AIRPLANE 📶",
                                                command=self.disable_airplane_mode,
                                                fg_color=self.COLOR_SUCCESS, hover_color=self.COLOR_SUCCESS_HOVER,
                                                corner_radius=10,
                                                height=50, text_color=self.COLOR_BACKGROUND,
                                                font=self.FONT_BUTTON)
        disable_airplane_button.grid(row=0, column=1, sticky='ew', padx=(7, 15), pady=15)

        # --- Section: Image Sharing ---
        self._create_section_header(utility_frame, "SHARE IMAGE TO TIKTOK LITE", 4)
        image_frame = self._create_section_frame(utility_frame, 5)

        self.image_file_name_entry = ctk.CTkEntry(image_frame,
                                                  placeholder_text="Enter image name in /sdcard/Download...",
                                                  height=45,
                                                  corner_radius=10, font=self.FONT_BODY)
        self.image_file_name_entry.grid(row=0, column=0, sticky='ew', padx=15, pady=(15, 10))

        self.share_image_button = ctk.CTkButton(image_frame, text="SHARE IMAGE",
                                                command=self.share_image_to_tiktok_lite,
                                                fg_color=self.COLOR_ACCENT, hover_color=self.COLOR_ACCENT_HOVER,
                                                height=50,
                                                font=self.FONT_BUTTON, corner_radius=10,
                                                text_color=self.COLOR_BACKGROUND)
        self.share_image_button.grid(row=1, column=0, sticky='ew', padx=15, pady=(10, 15))

    # --- ADB Utility Methods (Unchanged Logic) ---

    def _threaded_airplane_mode(self, mode):
        if not self.devices:
            self.after(0, lambda: self.status_label.configure(text="⚠️ No devices detected.",
                                                              text_color=self.COLOR_WARNING))
            return

        state = '1' if mode == 'enable' else '0'
        name = 'ENABLE' if mode == 'enable' else 'DISABLE'

        self.after(0, lambda: self.status_label.configure(
            text=f"[CMD] Sending {name} AIRPLANE MODE command...", text_color=self.COLOR_ACCENT))

        set_cmd = ['shell', 'settings', 'put', 'global', 'airplane_mode_on', state]
        broadcast_cmd = ['shell', 'am', 'broadcast', '-a', 'android.intent.action.AIRPLANE_MODE']

        for serial in self.devices:
            self.executor.submit(run_adb_command, set_cmd, serial)
            self.executor.submit(run_adb_command, broadcast_cmd, serial)

        self.after(0, lambda: self.status_label.configure(
            text=f"✅ AIRPLANE MODE {name} command sent to all devices.", text_color=self.COLOR_SUCCESS))

    def enable_airplane_mode(self):
        threading.Thread(target=self._threaded_airplane_mode, args=('enable',), daemon=True).start()

    def disable_airplane_mode(self):
        threading.Thread(target=self._threaded_airplane_mode, args=('disable',), daemon=True).start()

    def browse_apk_file(self):
        file_path = filedialog.askopenfilename(
            defaultextension=".apk",
            filetypes=[("APK files", "*.apk")]
        )
        if file_path:
            self.apk_path = file_path
            self.apk_path_entry.delete(0, tk.END)
            self.apk_path_entry.insert(0, os.path.basename(file_path))
            self.status_label.configure(text=f"✅ APK SELECTED: {os.path.basename(file_path)}",
                                        text_color=self.COLOR_SUCCESS)

    def install_apk_to_devices(self):
        if not self.apk_path or not os.path.exists(self.apk_path):
            self.status_label.configure(text="⚠️ Please select a valid APK file first.", text_color=self.COLOR_WARNING)
            return

        if not self.devices:
            self.status_label.configure(text="⚠️ No devices detected.", text_color=self.COLOR_WARNING)
            return

        self.status_label.configure(text=f"[CMD] Installing {os.path.basename(self.apk_path)} on all devices...",
                                    text_color=self.COLOR_ACCENT)

        command = ['install', '-r', self.apk_path]

        results = []

        def _install_task(serial):
            success, output = run_adb_command(command, serial)
            results.append((serial, success, output))

        futures = [self.executor.submit(_install_task, serial) for serial in self.devices]
        concurrent.futures.wait(futures)

        all_success = all(success for _, success, _ in results)
        if all_success:
            self.status_label.configure(text="✅ APK INSTALL SUCCESSFUL.", text_color=self.COLOR_SUCCESS)
        else:
            error_count = sum(1 for _, success, _ in results if not success)
            self.status_label.configure(text=f"❌ INSTALLATION FAILED on {error_count} device(s).",
                                        text_color=self.COLOR_DANGER)

    def update_app(self):
        def _update_in_thread():
            try:
                self.status_label.configure(text="[SYS] Downloading latest version...", text_color=self.COLOR_ACCENT)
                response = requests.get(UPDATE_URL)
                response.raise_for_status()

                desktop_path = Path.home() / "Desktop"
                old_file_path = Path(sys.executable) if getattr(sys, 'frozen', False) else Path(sys.argv[0])

                if not old_file_path.is_file():
                    new_file_path = desktop_path / "adb_tool_by_dars.py"
                elif old_file_path.suffix == '.py':
                    new_file_path = old_file_path.parent / old_file_path.name
                else:
                    new_file_path = desktop_path / old_file_path.name

                with open(new_file_path, 'wb') as f:
                    f.write(response.content)

                messagebox.showinfo("Update Complete",
                                    "The new version has been downloaded. The application will now close and update.")

                create_and_run_updater_script(new_file_path, old_file_path)

                self.destroy()

            except requests.exceptions.HTTPError:
                self.after(0, lambda: self.status_label.configure(
                    text=f"❌ ERROR: Update download failed. HTTP Error.",
                    text_color=self.COLOR_DANGER))
            except requests.exceptions.RequestException:
                self.after(0, lambda: self.status_label.configure(
                    text="❌ ERROR: Update download failed. Connection Error.",
                    text_color=self.COLOR_DANGER))
            except Exception as e:
                self.after(0, lambda: self.status_label.configure(
                    text=f"❌ ERROR: An unexpected update error occurred: {e}",
                    text_color=self.COLOR_DANGER))

        update_thread = threading.Thread(target=_update_in_thread, daemon=True)
        update_thread.start()

    # --- Dynamic Link/Caption Pair Management (Unchanged Logic) ---

    def add_share_pair(self, is_initial=False):
        # Use a secondary color for the inner frame to create depth
        frame = ctk.CTkFrame(self.share_pair_frame, fg_color=self.COLOR_FRAME, corner_radius=8, border_width=1,
                             border_color=self.COLOR_ACCENT_HOVER)
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=0)

        # Style the entry field
        share_url_entry = ctk.CTkEntry(frame,
                                       placeholder_text=f"Link #{len(self.share_pairs) + 1}: Enter link to share...",
                                       height=40, corner_radius=8, font=self.FONT_BODY,
                                       fg_color=self.COLOR_BACKGROUND, border_color=self.COLOR_BORDER, border_width=1)
        share_url_entry.grid(row=0, column=0, sticky='ew', padx=10, pady=(10, 5))

        if not is_initial:
            remove_button = ctk.CTkButton(frame, text="✖️", width=40, height=40, corner_radius=8,
                                          fg_color=self.COLOR_DANGER, hover_color=self.COLOR_DANGER_HOVER,
                                          command=lambda: self.remove_share_pair(frame), font=self.FONT_BUTTON)
            remove_button.grid(row=0, column=1, sticky='e', padx=(0, 10), pady=(10, 5))

        caption_frame = ctk.CTkFrame(frame, fg_color="transparent")
        caption_frame.grid(row=1, column=0, columnspan=2, sticky='ew', padx=10, pady=(0, 10))
        caption_frame.columnconfigure(0, weight=1)
        caption_frame.columnconfigure(1, weight=0)

        file_path_entry = ctk.CTkEntry(caption_frame, placeholder_text="Comment File Path: Select a text file...",
                                       height=40, corner_radius=8, font=self.FONT_BODY,
                                       fg_color=self.COLOR_BACKGROUND, border_color=self.COLOR_BORDER, border_width=1)
        file_path_entry.grid(row=0, column=0, sticky='ew', padx=(0, 7))

        browse_button = ctk.CTkButton(caption_frame, text="BROWSE TXT", corner_radius=8, width=120, height=40,
                                      fg_color=self.COLOR_SUCCESS, hover_color=self.COLOR_SUCCESS_HOVER,
                                      font=self.FONT_BUTTON, text_color=self.COLOR_BACKGROUND,
                                      command=lambda: self.browse_share_pair_file(file_path_entry))
        browse_button.grid(row=0, column=1, sticky='e')

        self.share_pairs.append({
            'frame': frame,
            'url_entry': share_url_entry,
            'file_entry': file_path_entry
        })
        frame.pack(fill='x', padx=5, pady=10)
        self.share_pair_frame.update_idletasks()

    def remove_share_pair(self, pair_frame_to_remove):
        for i, pair in enumerate(self.share_pairs):
            if pair['frame'] == pair_frame_to_remove:
                pair['frame'].destroy()
                self.share_pairs.pop(i)
                self.status_label.configure(text=f"✅ Link/Comment Pair removed.", text_color=self.COLOR_SUCCESS)
                if self.is_auto_typing.is_set():
                    self.stop_auto_type_loop()
                    self.after(100, self.start_auto_type_loop)
                return

    def browse_share_pair_file(self, target_entry):
        file_path = filedialog.askopenfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if file_path:
            target_entry.delete(0, tk.END)
            target_entry.insert(0, file_path)
            self.status_label.configure(text=f"✅ FILE SELECTED: {os.path.basename(file_path)}",
                                        text_color=self.COLOR_SUCCESS)
            if self.is_auto_typing.is_set():
                self.stop_auto_type_loop()
                self.after(100, self.start_auto_type_loop)

    def _threaded_send_text(self):
        file_paths = []
        for pair in self.share_pairs:
            file_path = pair['file_entry'].get()
            if file_path and os.path.exists(file_path):
                file_paths.append(file_path)

        if not file_paths:
            self.status_label.configure(text="⚠️ Please select a text file for at least one pair.",
                                        text_color=self.COLOR_WARNING)
            return

        if not self.devices:
            self.status_label.configure(text="⚠️ No devices detected.", text_color=self.COLOR_WARNING)
            return

        random_file_path = random.choice(file_paths)

        try:
            with open(random_file_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()

            clean_lines = [line.strip() for line in lines if line.strip()]

            if not clean_lines:
                self.status_label.configure(
                    text=f"⚠️ The selected file '{os.path.basename(random_file_path)}' is empty.",
                    text_color=self.COLOR_WARNING)
                return

            self.status_label.configure(
                text=f"[CMD] Sending random text from file '{os.path.basename(random_file_path)}' to all devices...",
                text_color=self.COLOR_ACCENT)

            for device_serial in self.devices:
                random_text = random.choice(clean_lines)
                self.executor.submit(run_text_command, random_text, device_serial)

            self.status_label.configure(text=f"✅ Text commands submitted.", text_color=self.COLOR_SUCCESS)


        except FileNotFoundError:
            self.status_label.configure(text="❌ ERROR: File not found.", text_color=self.COLOR_DANGER)
        except Exception as e:
            self.status_label.configure(text=f"❌ ERROR: An error occurred: {e}", text_color=self.COLOR_DANGER)

    def send_text_to_devices(self):
        send_thread = threading.Thread(target=self._threaded_send_text, daemon=True)
        send_thread.start()

    # --- Auto-Type Logic (Updated for Deep Link Fix) ---

    def start_auto_type_loop(self):
        if self.is_auto_typing.is_set():
            return

        valid_pairs = []
        for pair in self.share_pairs:
            share_url = pair['url_entry'].get()
            file_path = pair['file_entry'].get()

            if share_url:
                valid_pairs.append({'url': share_url, 'file': file_path})

        if not valid_pairs:
            self.status_label.configure(text="⚠️ No valid Links found. Please enter at least one URL.",
                                        text_color=self.COLOR_WARNING)
            return

        if not self.devices:
            self.status_label.configure(text="⚠️ No devices detected.", text_color=self.COLOR_WARNING)
            return

        self.is_auto_typing.set()

        self.find_click_type_button.configure(text="🛑 STOP AUTO-TYPE 🛑",
                                              fg_color=self.COLOR_DANGER,
                                              hover_color=self.COLOR_DANGER_HOVER,
                                              text_color=self.COLOR_TEXT_PRIMARY)

        self.status_label.configure(text="[CMD] Auto-type loop STARTED.", text_color=self.COLOR_SUCCESS)

        threading.Thread(target=self._threaded_find_click_type_LOOP, args=(valid_pairs,), daemon=True).start()

    def stop_auto_type_loop(self):
        self.is_auto_typing.clear()

        if hasattr(self, 'find_click_type_button') and self.find_click_type_button.winfo_exists():
            self.find_click_type_button.configure(text="⚡ START AUTO-TYPE ⚡",
                                                  fg_color=self.COLOR_ACCENT,
                                                  hover_color=self.COLOR_ACCENT_HOVER,
                                                  text_color=self.COLOR_BACKGROUND)

    def toggle_auto_type_loop(self):
        if self.is_auto_typing.is_set():
            self.stop_auto_type_loop()
        else:
            self.start_auto_type_loop()

    def _run_task_with_retry(self, serial, text_to_send, pair_index, max_retries=5):
        for attempt in range(max_retries):
            if not self.is_auto_typing.is_set() or is_stop_requested.is_set():
                return False, "Stop requested"

            success, message = self._run_find_click_type_on_device(serial, text_to_send)

            if success:
                self.after(0, lambda: self.status_label.configure(
                    text=f"✅ Pair {pair_index} on {serial} SUCCESSFUL (Attempt {attempt + 1}).",
                    text_color=self.COLOR_SUCCESS))
                return True, message
            else:
                if attempt < max_retries - 1:
                    wait_time = 3 + attempt * 2
                    self.after(0, lambda: self.status_label.configure(
                        text=f"⚠️ Pair {pair_index} on {serial}: Failed ({message}). Retrying in {wait_time}s (Attempt {attempt + 2}/{max_retries}).",
                        text_color=self.COLOR_WARNING))

                    time.sleep(wait_time)
                else:
                    self.after(0, lambda: self.status_label.configure(
                        text=f"❌ Pair {pair_index} on {serial}: FAILED after {max_retries} attempts ({message}). Moving to next pair.",
                        text_color=self.COLOR_DANGER))
                    return False, message

        return False, "Max retries reached"

    def _threaded_find_click_type_LOOP(self, valid_pairs):
        try:
            while self.is_auto_typing.is_set() and not is_stop_requested.is_set():

                success_achieved_in_this_cycle = False

                if not self.devices:
                    self.after(0, lambda: self.status_label.configure(text="⚠️ No devices, stopping loop.",
                                                                      text_color=self.COLOR_WARNING))
                    break

                for index, selected_pair in enumerate(valid_pairs):
                    if not self.is_auto_typing.is_set() or is_stop_requested.is_set():
                        break

                    share_url = selected_pair['url']
                    file_path = selected_pair['file']
                    pair_index = index + 1
                    total_pairs = len(valid_pairs)

                    self.after(0, lambda: self.status_label.configure(
                        text=f"[CMD] Processing Pair {pair_index}/{total_pairs}: Opening {share_url[:20]}...",
                        # Changed text to reflect deep linking
                        text_color=self.COLOR_ACCENT))

                    clean_lines = []
                    has_caption = False

                    if file_path and os.path.exists(file_path):
                        try:
                            with open(file_path, 'r', encoding='utf-8') as f:
                                lines = f.readlines()
                            clean_lines = [line.strip() for line in lines if line.strip()]

                            if clean_lines:
                                has_caption = True
                            else:
                                self.after(0, lambda: self.status_label.configure(
                                    text=f"⚠️ Comment file '{os.path.basename(file_path)}' is empty. Share-only mode.",
                                    text_color=self.COLOR_WARNING))
                        except Exception as e:
                            self.after(0, lambda: self.status_label.configure(
                                text=f"❌ Error reading file: {e}. Share-only mode.", text_color=self.COLOR_DANGER))
                    else:
                        self.after(0, lambda: self.status_label.configure(
                            text=f"ℹ️ No comment file for Pair {pair_index}. Share-only mode.",
                            text_color=self.COLOR_TEXT_SECONDARY))

                    # --- FIXED DEEP LINKING COMMAND ---
                    share_command = [
                        'shell', 'am', 'start',
                        '-a', 'android.intent.action.VIEW',  # ACTION_VIEW for deep linking
                        '-d', f'"{share_url}"',  # -d for the data URI (URL)
                        TIKTOK_LITE_PACKAGE
                    ]
                    # --- END FIXED DEEP LINKING COMMAND ---

                    share_futures = []
                    for serial in self.devices:
                        if not self.is_auto_typing.is_set() or is_stop_requested.is_set():
                            break
                        share_futures.append(self.executor.submit(run_adb_command, share_command, serial))

                    concurrent.futures.wait(share_futures)

                    time.sleep(5)

                    if not self.is_auto_typing.is_set() or is_stop_requested.is_set():
                        break

                    if has_caption:
                        self.after(0, lambda: self.status_label.configure(
                            text=f"[CMD] Pair {pair_index}: Starting typing and retry attempts...",
                            text_color=self.COLOR_ACCENT))

                        futures = []
                        for serial in self.devices:
                            if not self.is_auto_typing.is_set() or is_stop_requested.is_set():
                                break

                            random_text = random.choice(clean_lines)
                            futures.append(
                                self.executor.submit(self._run_task_with_retry, serial, random_text, pair_index))

                        concurrent.futures.wait(futures)

                        pair_success = False
                        for future in futures:
                            if future.exception() is None:
                                success, _ = future.result()
                                if success:
                                    pair_success = True
                    else:
                        self.after(0, lambda: self.status_label.configure(
                            text=f"✅ Pair {pair_index}: SHARE-ONLY complete.",
                            text_color=self.COLOR_SUCCESS))
                        pair_success = True

                    if pair_success:
                        success_achieved_in_this_cycle = True

                    COOLDOWN = 10
                    self.after(0, lambda: self.status_label.configure(
                        text=f"[SYS] Pair {pair_index} processed. Waiting {COOLDOWN}s before next pair...",
                        text_color=self.COLOR_TEXT_SECONDARY))

                    for _ in range(COOLDOWN):
                        if not self.is_auto_typing.is_set() or is_stop_requested.is_set():
                            break
                        time.sleep(1)

                    if not self.is_auto_typing.is_set() or is_stop_requested.is_set():
                        break

                if success_achieved_in_this_cycle:
                    self.after(0, lambda: self.status_label.configure(
                        text="✅ AUTO-TYPE SUCCESSFUL (Posted/Shared). Stopping loop.",
                        text_color=self.COLOR_SUCCESS))
                    break
                else:
                    self.after(0, lambda: self.status_label.configure(
                        text="[SYS] All pairs processed (No successful post). Waiting 5s for next cycle...",
                        text_color=self.COLOR_TEXT_PRIMARY))

                    wait_duration = 5
                    for _ in range(wait_duration):
                        if not self.is_auto_typing.is_set() or is_stop_requested.is_set():
                            break
                        time.sleep(1)

        except Exception as e:
            self.after(0, lambda: self.status_label.configure(
                text=f"❌ CRITICAL ERROR in auto-type task: {e}", text_color=self.COLOR_DANGER))
        finally:
            self.after(0, self.stop_auto_type_loop)

    def _run_find_click_type_on_device(self, serial, text_to_send):
        local_xml_file = f"ui_dump_{serial}_{uuid.uuid4()}.xml"

        # --- CONSTANTS ---
        COMMENT_BUTTON_ID = "com.zhiliaoapp.musically.go:id/dj4"
        POST_BUTTON_ID = "com.zhiliaoapp.musically.go:id/e_x"

        try:
            # 1. DUMP UI (Pre-tap)
            if self.is_auto_typing.is_set() and not is_stop_requested.is_set():
                dump_cmd = ['shell', 'uiautomator', 'dump', '/data/local/tmp/ui.xml']
                success, out = run_adb_command(dump_cmd, serial)
                if not success:
                    return False, "Failed to dump UI (Pre-tap)"
            else:
                return False, "Stop requested"

            # 2. PULL UI XML (Pre-tap)
            if self.is_auto_typing.is_set() and not is_stop_requested.is_set():
                pull_cmd = ['pull', '/data/local/tmp/ui.xml', local_xml_file]
                success, out = run_adb_command(pull_cmd, serial)
                if not success:
                    return False, "Failed to pull UI XML (Pre-tap)"
            else:
                return False, "Stop requested"

            if not os.path.exists(local_xml_file):
                return False, "XML file not found (Pre-tap)"

            tree = ET.parse(local_xml_file)
            root = tree.getroot()

            # 3. FIND AND TAP COMMENT BUTTON (dj4)
            if self.is_auto_typing.is_set() and not is_stop_requested.is_set():

                comment_button_node = root.find(f'.//node[@resource-id="{COMMENT_BUTTON_ID}"]')

                if comment_button_node is not None:
                    bounds_str = comment_button_node.get('bounds')
                    coords = re.findall(r'\d+', bounds_str)
                    if len(coords) < 4:
                        return False, f"Invalid bounds string for comment button '{COMMENT_BUTTON_ID}'"

                    x1, y1, x2, y2 = map(int, coords[:4])

                    tap_x = (x1 + x2) // 2
                    tap_y = (y1 + y2) // 2

                    # Tap the comment button
                    tap_cmd = ['shell', 'input', 'tap', str(tap_x), str(tap_y)]
                    success, out = run_adb_command(tap_cmd, serial)
                    if not success:
                        return False, "Failed to tap the comment button (dj4)"

                    # --- MODIFIED: Reduced delay from 3s to 1.5s ---
                    time.sleep(1.5)

                    # RERUN DUMP/PULL TO GET THE NEW SCREEN WITH EDITTEXT

            if self.is_auto_typing.is_set() and not is_stop_requested.is_set():
                # Re-dump UI
                dump_cmd = ['shell', 'uiautomator', 'dump', '/data/local/tmp/ui.xml']
                success, out = run_adb_command(dump_cmd, serial)
                if not success:
                    return False, "Failed to dump UI (Post-tap)"

                # Re-pull XML
                pull_cmd = ['pull', '/data/local/tmp/ui.xml', local_xml_file]
                success, out = run_adb_command(pull_cmd, serial)
                if not success:
                    return False, "Failed to pull UI XML (Post-tap)"

                tree = ET.parse(local_xml_file)
                root = tree.getroot()
            else:
                return False, "Stop requested"

            # 4. FIND & TAP EDITTEXT
            edit_text_node = root.find('.//node[@class="android.widget.EditText"]')

            if edit_text_node is None:
                return False, "No EditText found (Comment box not ready/visible)"

            bounds_str = edit_text_node.get('bounds')
            if not bounds_str:
                return False, "EditText found but has no bounds"

            coords = re.findall(r'\d+', bounds_str)
            x1, y1, x2, y2 = map(int, coords[:4])
            tap_x = (x1 + x2) // 2
            tap_y = (y1 + y2) // 2

            if self.is_auto_typing.is_set() and not is_stop_requested.is_set():
                # Tap EditText
                tap_cmd = ['shell', 'input', 'tap', str(tap_x), str(tap_y)]
                success, out = run_adb_command(tap_cmd, serial)
                if not success:
                    return False, "Failed to tap EditText"

                # --- MODIFIED: Reduced delay from 1s to 0.5s ---
                time.sleep(0.5)

                # 5. TYPE TEXT & SEND ENTER
            if self.is_auto_typing.is_set() and not is_stop_requested.is_set():
                # Type text
                run_text_command(text_to_send, serial)

                # --- MODIFIED: Reduced delay from 1s to 0.5s ---
                time.sleep(0.5)

                # Send 'ENTER' keyevent (Keycode 66) - This often finalizes the text input
                post_cmd = ['shell', 'input', 'keyevent', '66']
                run_adb_command(post_cmd, serial)

                # --- MODIFIED: Reduced delay from 1s to 0.5s ---
                time.sleep(0.5)
            else:
                return False, "Stop requested"

            # --- START: FIND AND TAP POST BUTTON (e_x) ---

            # Re-dump UI to capture the screen state after typing (where the 'Post' button might activate/move)
            if self.is_auto_typing.is_set() and not is_stop_requested.is_set():
                dump_cmd = ['shell', 'uiautomator', 'dump', '/data/local/tmp/ui.xml']
                run_adb_command(dump_cmd, serial)

                pull_cmd = ['pull', '/data/local/tmp/ui.xml', local_xml_file]
                run_adb_command(pull_cmd, serial)

                tree = ET.parse(local_xml_file)
                root = tree.getroot()
                time.sleep(0.5)  # Small pause after dump/pull
            else:
                return False, "Stop requested"

            # Find the 'Post' button by resource-id
            post_button_node = root.find(f'.//node[@resource-id="{POST_BUTTON_ID}"]')

            if post_button_node is None:
                return False, f"Post button '{POST_BUTTON_ID}' not found after typing/enter. Failed to post."
            else:
                bounds_str = post_button_node.get('bounds')
                if not bounds_str:
                    return False, f"Post button '{POST_BUTTON_ID}' found but has no bounds"

                coords = re.findall(r'\d+', bounds_str)
                if len(coords) < 4:
                    return False, f"Invalid bounds string for Post button '{POST_BUTTON_ID}'"

                x1, y1, x2, y2 = map(int, coords[:4])

                tap_x = (x1 + x2) // 2
                tap_y = (y1 + y2) // 2

                # Tap the post button to send the comment
                tap_cmd = ['shell', 'input', 'tap', str(tap_x), str(tap_y)]
                success, out = run_adb_command(tap_cmd, serial)
                if not success:
                    return False, "Failed to tap the Post button (e_x)"

                # --- MODIFIED: Reduced delay from 1s to 0.5s ---
                time.sleep(0.5)
                # --- END: FIND AND TAP POST BUTTON (e_x) ---

            return True, "Success"

        except ET.ParseError:
            return False, "Failed to parse XML"
        except Exception as e:
            return False, str(e)
        finally:
            if os.path.exists(local_xml_file):
                os.remove(local_xml_file)

    def remove_emojis_from_file(self):
        if not self.share_pairs:
            self.status_label.configure(text="⚠️ Please add a Link/Comment Pair first.", text_color=self.COLOR_WARNING)
            return

        file_path = self.share_pairs[0]['file_entry'].get()
        if not file_path:
            self.status_label.configure(text="⚠️ Please select a text file for the first pair.",
                                        text_color=self.COLOR_WARNING)
            return

        try:
            emoji_pattern = re.compile("["
                                       "\U0001F600-\U0001F64F"
                                       "\U0001F300-\U0001F5FF"
                                       "\U0001F680-\U0001F6FF"
                                       "\U0001F700-\U0001F77F"
                                       "\U0001F780-\U0001F7FF"
                                       "\U0001F800-\U0001F8FF"
                                       "\U0001F900-\U0001F9FF"
                                       "\U0001FA00-\U0001FA6F"
                                       "\U0001FA70-\U0001FAFF"
                                       "\U00002702-\U000027B0"
                                       "\U00002600-\U000026FF"
                                       "\U000025A0-\U000025FF"
                                       "]+", flags=re.UNICODE)

            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()

            cleaned_content = emoji_pattern.sub(r'', content)

            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(cleaned_content)

            self.status_label.configure(text=f"✅ EMOJIS REMOVED from file: {os.path.basename(file_path)}.",
                                        text_color=self.COLOR_SUCCESS)

        except FileNotFoundError:
            self.status_label.configure(text="❌ ERROR: File not found.", text_color=self.COLOR_DANGER)
        except Exception as e:
            self.status_label.configure(text=f"❌ ERROR: An error occurred: {e}", text_color=self.COLOR_DANGER)

    # --- Device Management & UI (Fixed Logic) ---

    def detect_devices(self):
        self.stop_capture()

        for widget in self.device_view_panel.winfo_children():
            widget.destroy()

        self.device_frames = {}
        self.device_canvases = {}
        self.device_images = {}
        self.press_start_coords = {}
        self.press_time = {}
        self.selected_device_serial = None
        self.devices = []
        self.status_label.configure(text="[SYS] Detecting devices...", text_color=self.COLOR_ACCENT)

        try:
            result = subprocess.run(['adb', 'devices'], capture_output=True, text=True, check=True, timeout=10)
            devices_output = result.stdout.strip().split('\n')[1:]
            self.devices = [line.split('\t')[0] for line in devices_output if line.strip() and 'device' in line]
        except (subprocess.CalledProcessError, FileNotFoundError, subprocess.TimeoutExpired):
            messagebox.showerror("Error", "ADB is not installed, not in your system PATH, or timed out.")
            self.status_label.configure(text="❌ ERROR: ADB not found or timed out.", text_color=self.COLOR_DANGER)
            self.device_count_label.configure(text="DEVICES: 0")
            self.device_option_menu.configure(values=["No devices found"], state="disabled")
            self.device_selector_var.set("No devices found")
            return

        self.device_count_label.configure(text=f"DEVICES: {len(self.devices)}")

        if not self.devices:
            no_devices_label = ctk.CTkLabel(self.device_view_panel,
                                            text="NO DEVICES FOUND.\nEnsure USB debugging is enabled.",
                                            font=self.FONT_HEADING, text_color=self.COLOR_TEXT_SECONDARY)
            no_devices_label.pack(expand=True)
            self.status_label.configure(text="⚠️ No devices detected.", text_color=self.COLOR_WARNING)
            self.device_option_menu.configure(values=["No devices found"], state="disabled")
            self.device_selector_var.set("No devices found")
        else:
            self.status_label.configure(text=f"✅ {len(self.devices)} devices connected.", text_color=self.COLOR_SUCCESS)
            self.device_option_menu.configure(values=self.devices, state="normal")
            self.device_selector_var.set(self.devices[0])
            self.on_device_select_menu(self.devices[0])

    def on_device_select_menu(self, selected_serial):
        """FIXED: Uses selected_serial parameter correctly."""
        if not selected_serial or selected_serial == "No devices found":
            return

        self.stop_capture()
        self.selected_device_serial = selected_serial  # FIXED: Changed 'serial' to 'selected_serial'

        for widget in self.device_view_panel.winfo_children():
            widget.destroy()

        self.device_frames = {}
        self.device_canvases = {}
        self.device_images = {}
        self.press_start_coords = {}
        self.press_time = {}

        self.create_device_frame(self.selected_device_serial)
        self.start_capture_process()

    def stop_capture(self):
        self.is_capturing = False
        if self.update_image_id:
            self.after_cancel(self.update_image_id)
            self.update_image_id = None
        if self.capture_thread and self.capture_thread.is_alive():
            pass
        self.screenshot_queue.queue.clear()

    def start_capture_process(self):
        if self.is_capturing:
            return

        self.is_capturing = True
        self.capture_thread = threading.Thread(target=self.capture_screen_loop, daemon=True)
        self.capture_thread.start()
        self.update_image_id = self.after(100, self.update_image)

    def capture_screen_loop(self):
        while self.is_capturing:
            try:
                if not self.selected_device_serial:
                    self.is_capturing = False
                    break

                process = subprocess.run(['adb', '-s', self.selected_device_serial, 'exec-out', 'screencap', '-p'],
                                         capture_output=True, check=True, timeout=5)
                self.screenshot_queue.put(process.stdout)
            except subprocess.CalledProcessError:
                self.is_capturing = False
            except subprocess.TimeoutExpired:
                pass
            except Exception:
                self.is_capturing = False

    def update_image(self):
        try:
            if not self.selected_device_serial or not self.is_capturing:
                return

            canvas = self.device_canvases.get(self.selected_device_serial)
            if not canvas or not canvas.winfo_exists():
                return

            if not self.screenshot_queue.empty():
                image_data = self.screenshot_queue.get()
                pil_image = Image.open(io.BytesIO(image_data))

                canvas_width = canvas.winfo_width()
                canvas_height = canvas.winfo_height()
                if canvas_width > 0 and canvas_height > 0:
                    img_width, img_height = pil_image.size
                    aspect_ratio = img_width / img_height

                    if canvas_width / canvas_height > aspect_ratio:
                        new_height = canvas_height
                        new_width = int(new_height * aspect_ratio)
                    else:
                        new_width = canvas_width
                        new_height = int(new_width / aspect_ratio)

                    if new_width > 0 and new_height > 0:
                        resized_image = pil_image.resize((new_width, new_height), Image.Resampling.LANCZOS)
                        tk_image = ImageTk.PhotoImage(resized_image)

                        self.device_images[self.selected_device_serial] = {'pil_image': pil_image, 'tk_image': tk_image}

                        x_pos = canvas_width / 2
                        y_pos = canvas_height / 2

                        if 'item_id' in self.device_images.get(self.selected_device_serial, {}):
                            image_item_id = self.device_images[self.selected_device_serial]['item_id']
                            canvas.coords(image_item_id, x_pos, y_pos)
                            canvas.itemconfig(image_item_id, image=tk_image)
                        else:
                            image_item_id = canvas.create_image(x_pos, y_pos, image=tk_image)
                            self.device_images[self.selected_device_serial]['item_id'] = image_item_id
                            canvas.itemconfig(image_item_id, anchor=tk.CENTER)

            if self.is_capturing:
                self.update_image_id = self.after(100, self.update_image)

        except Exception:
            self.stop_capture()

    def create_device_frame(self, serial):
        device_frame = ctk.CTkFrame(self.device_view_panel, fg_color="transparent")
        device_frame.pack(padx=25, pady=25, fill=tk.BOTH, expand=True)
        self.device_frames[serial] = device_frame

        title = ctk.CTkLabel(device_frame, text=f"LIVE CONTROL: {serial}", font=self.FONT_HEADING,
                             text_color=self.COLOR_ACCENT)
        title.pack(pady=(0, 15))

        canvas_container = ctk.CTkFrame(device_frame, fg_color=self.COLOR_FRAME, corner_radius=12,
                                        border_width=1, border_color=self.COLOR_BORDER)
        canvas_container.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)
        canvas_container.bind("<Configure>", self.on_canvas_container_resize)

        canvas = tk.Canvas(canvas_container, bg=self.COLOR_FRAME, highlightthickness=0)
        canvas.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
        self.device_canvases[serial] = canvas

        canvas.bind("<ButtonPress-1>", lambda event: self.start_press(event, serial))
        canvas.bind("<ButtonRelease-1>", lambda event: self.handle_release(event, serial))

        # --- Button Frame (TWO ROWS) ---
        button_frame = ctk.CTkFrame(device_frame, fg_color="transparent")
        button_frame.pack(pady=(20, 0), fill="x")

        button_frame.columnconfigure(0, weight=1)
        button_frame.columnconfigure(1, weight=1)
        button_frame.columnconfigure(2, weight=1)

        # Define base style (all common args)
        button_style = {'corner_radius': 10,
                        'fg_color': self.COLOR_FRAME,
                        'hover_color': self.COLOR_BORDER,
                        'text_color': self.COLOR_TEXT_PRIMARY,
                        'border_color': self.COLOR_BORDER, 'border_width': 1,
                        'height': 45, 'font': self.FONT_BUTTON}

        button_padx = 7

        # ROW 0: Navigation
        home_button = ctk.CTkButton(button_frame, text="HOME 🏠", command=lambda: self.send_adb_keyevent(3),
                                    **button_style)
        home_button.grid(row=0, column=0, padx=button_padx, pady=(0, 10), sticky="ew")

        back_button = ctk.CTkButton(button_frame, text="BACK ↩️", command=lambda: self.send_adb_keyevent(4),
                                    **button_style)
        back_button.grid(row=0, column=1, padx=button_padx, pady=(0, 10), sticky="ew")

        recents_button = ctk.CTkButton(button_frame, text="RECENTS", command=lambda: self.send_adb_keyevent(187),
                                       **button_style)
        recents_button.grid(row=0, column=2, padx=button_padx, pady=(0, 10), sticky="ew")

        # ROW 1: Actions
        scroll_down_button = ctk.CTkButton(button_frame, text="SCROLL DOWN",
                                           command=lambda: self.send_adb_swipe(serial, 'up'), **button_style)
        scroll_down_button.grid(row=1, column=0, padx=button_padx, sticky="ew")

        scroll_up_button = ctk.CTkButton(button_frame, text="SCROLL UP",
                                         command=lambda: self.send_adb_swipe(serial, 'down'), **button_style)
        scroll_up_button.grid(row=1, column=1, padx=button_padx, sticky="ew")

        danger_style = button_style.copy()
        danger_style['fg_color'] = self.COLOR_DANGER
        danger_style['hover_color'] = self.COLOR_DANGER_HOVER

        close_button = ctk.CTkButton(button_frame, text="SCREEN OFF 💡", command=lambda: self.send_adb_keyevent(26),
                                     **danger_style)

        close_button.grid(row=1, column=2, padx=button_padx, sticky="ew")

    def on_canvas_container_resize(self, event):
        if not self.selected_device_serial:
            return

        canvas = self.device_canvases.get(self.selected_device_serial)
        if not canvas:
            return

        container_width = event.width
        container_height = event.height

        aspect_ratio = 9 / 16

        if container_width / container_height > aspect_ratio:
            new_height = container_height
            new_width = int(new_height * aspect_ratio)
        else:
            new_width = container_width
            new_height = int(new_width / aspect_ratio)

        canvas.configure(width=new_width, height=new_height)
        canvas.place(relx=0.5, rely=0.5, anchor=tk.CENTER, width=new_width, height=new_height)

        self.after(10, self.update_image)

    def start_press(self, event, serial):
        self.press_time[serial] = time.time()
        self.press_start_coords[serial] = (event.x, event.y)

    def handle_release(self, event, serial):
        end_time = time.time()
        start_time = self.press_time.get(serial)

        if not start_time:
            return

        duration = end_time - start_time
        start_x, start_y = self.press_start_coords.get(serial, (event.x, event.y))
        end_x, end_y = (event.x, event.y)
        distance = ((end_x - start_x) ** 2 + (end_y - start_y) ** 2) ** 0.5

        if distance > self.drag_threshold:
            self.send_adb_swipe_command(start_x, start_y, end_x, end_y, serial)
        elif duration > self.long_press_duration:
            self.send_adb_long_press(event, serial)
        else:
            self.send_adb_tap(event, serial)

        self.press_time.pop(serial, None)
        self.press_start_coords.pop(serial, None)

    def _get_scaled_coords(self, canvas_x, canvas_y, serial):
        pil_image_info = self.device_images.get(self.selected_device_serial, {})
        pil_image = pil_image_info.get('pil_image')

        if not pil_image:
            return None, None

        img_width, img_height = pil_image.size
        canvas = self.device_canvases[serial]
        canvas_width = canvas.winfo_width()
        canvas_height = canvas.winfo_height()

        canvas_aspect = canvas_width / canvas_height
        image_aspect = img_width / img_height

        if canvas_aspect > image_aspect:
            effective_height = canvas_height
            effective_width = int(effective_height * image_aspect)
        else:
            effective_width = canvas_width
            effective_height = int(effective_width / image_aspect)

        image_x_offset = (canvas_width - effective_width) // 2
        image_y_offset = (canvas_height - effective_height) // 2

        click_x = canvas_x - image_x_offset
        click_y = canvas_y - image_y_offset

        if not (0 <= click_x < effective_width and 0 <= click_y < effective_height):
            return None, None

        try:
            adb_size_output = subprocess.run(['adb', '-s', serial, 'shell', 'wm', 'size'], capture_output=True,
                                             text=True, check=True, timeout=5).stdout.strip()
            adb_width, adb_height = map(int, adb_size_output.split()[-1].split('x'))
        except Exception:
            return None, None

        scaled_x = int(click_x * adb_width / effective_width)
        scaled_y = int(click_y * adb_height / effective_height)

        return scaled_x, scaled_y

    def send_adb_tap(self, event, serial):
        scaled_x, scaled_y = self._get_scaled_coords(event.x, event.y, serial)
        if scaled_x is None:
            self.status_label.configure(text=f"⚠️ Tap ignored (outside screen area).", text_color=self.COLOR_WARNING)
            return

        command = ['shell', 'input', 'tap', str(scaled_x), str(scaled_y)]
        for device_serial in self.devices:
            self.executor.submit(run_adb_command, command, device_serial)
        self.status_label.configure(text=f"✅ TAP command sent.", text_color=self.COLOR_SUCCESS)

    def send_adb_long_press(self, event, serial):
        scaled_x, scaled_y = self._get_scaled_coords(event.x, event.y, serial)
        if scaled_x is None:
            self.status_label.configure(text=f"⚠️ Long press ignored (outside screen area).",
                                        text_color=self.COLOR_WARNING)
            return

        command = ['shell', 'input', 'swipe', str(scaled_x), str(scaled_y), str(scaled_x), str(scaled_y), '1000']
        for device_serial in self.devices:
            self.executor.submit(run_adb_command, command, device_serial)
        self.status_label.configure(text=f"✅ LONG PRESS command sent.", text_color=self.COLOR_SUCCESS)

    def send_adb_swipe_command(self, start_x, start_y, end_x, end_y, serial):
        scaled_start_x, scaled_start_y = self._get_scaled_coords(start_x, start_y, serial)
        scaled_end_x, scaled_end_y = self._get_scaled_coords(end_x, end_y, serial)

        if scaled_start_x is None or scaled_end_x is None:
            self.status_label.configure(text=f"⚠️ Swipe ignored (outside screen area).", text_color=self.COLOR_WARNING)
            return

        command = ['shell', 'input', 'swipe',
                   str(scaled_start_x), str(scaled_start_y),
                   str(scaled_end_x), str(scaled_end_y), '300']

        for device_serial in self.devices:
            self.executor.submit(run_adb_command, command, device_serial)
        self.status_label.configure(text=f"✅ SWIPE command sent.", text_color=self.COLOR_SUCCESS)

    def send_adb_swipe(self, serial, direction):
        try:
            adb_width_str = subprocess.run(['adb', '-s', serial, 'shell', 'wm', 'size'], capture_output=True, text=True,
                                           check=True).stdout.strip().split()[-1]
            adb_width, adb_height = map(int, adb_width_str.split('x'))

            if direction == 'down':
                start_x, start_y = adb_width // 2, adb_height // 4 * 3
                end_x, end_y = start_x, adb_height // 4
            elif direction == 'up':
                start_x, start_y = adb_width // 2, adb_height // 4
                end_x, end_y = start_x, adb_height // 4 * 3

            command = ['shell', 'input', 'swipe',
                       str(start_x), str(start_y), str(end_x), str(end_y), '300']
            for device_serial in self.devices:
                self.executor.submit(run_adb_command, command, device_serial)
            self.status_label.configure(text=f"✅ {direction.upper()} SCROLL command sent.",
                                        text_color=self.COLOR_SUCCESS)
        except Exception as e:
            self.status_label.configure(text=f"❌ ERROR: Failed to send scroll command: {e}",
                                        text_color=self.COLOR_DANGER)

    def send_adb_keyevent(self, keycode):
        command = ['shell', 'input', 'keyevent', str(keycode)]
        for device_serial in self.devices:
            self.executor.submit(run_adb_command, command, device_serial)

        key_name = {3: "HOME", 4: "BACK", 187: "RECENTS", 24: "VOL UP", 25: "VOL DOWN", 26: "POWER/SCREEN OFF"}.get(
            keycode, "KEY EVENT")
        self.status_label.configure(text=f"✅ {key_name} command sent.", text_color=self.COLOR_SUCCESS)

    # --- TikTok Lite Specific Methods (UPDATED PACKAGE & ACTIVITY) ---

    def open_tiktok_lite_deeplink(self):
        post_url = self.tiktok_url_entry.get()
        if not post_url or not self.devices:
            self.status_label.configure(text="⚠️ Check URL and devices.", text_color=self.COLOR_WARNING)
            return

        self.status_label.configure(text=f"[CMD] Opening TikTok video URL...", text_color=self.COLOR_ACCENT)

        command = [
            'shell', 'am', 'start',
            '-a', 'android.intent.action.VIEW',
            '-d', f'"{post_url}"',
            TIKTOK_LITE_PACKAGE
        ]
        for device_serial in self.devices:
            self.executor.submit(run_adb_command, command, device_serial)
        self.status_label.configure(text="✅ Visited TikTok video on all devices.", text_color=self.COLOR_SUCCESS)

    def launch_tiktok_lite(self):
        if not self.devices:
            self.status_label.configure(text="⚠️ No devices detected.", text_color=self.COLOR_WARNING)
            return

        self.status_label.configure(text=f"[CMD] Launching TikTok Lite...", text_color=self.COLOR_ACCENT)

        # Ginamit ang bagong package at activity
        command = ['shell', 'am', 'start', '-n', f'{TIKTOK_LITE_PACKAGE}/{TIKTOK_LITE_ACTIVITY}']
        for device_serial in self.devices:
            self.executor.submit(run_adb_command, command, device_serial)
        self.status_label.configure(text="✅ Launched TikTok Lite on all devices.", text_color=self.COLOR_SUCCESS)

    def force_stop_tiktok_lite(self):
        if not self.devices:
            self.status_label.configure(text="⚠️ No devices detected.", text_color=self.COLOR_WARNING)
            return

        self.status_label.configure(text=f"[CMD] Force stopping TikTok Lite...", text_color=self.COLOR_DANGER)

        command = ['shell', 'am', 'force-stop', TIKTOK_LITE_PACKAGE]
        for device_serial in self.devices:
            self.executor.submit(run_adb_command, command, device_serial)
        self.status_label.configure(text="✅ Force stopped TikTok Lite on all devices.", text_color=self.COLOR_SUCCESS)

    def share_image_to_tiktok_lite(self):
        file_name = self.image_file_name_entry.get()
        if not file_name or not self.devices:
            self.status_label.configure(text="⚠️ Check image filename and devices.", text_color=self.COLOR_WARNING)
            return

        self.status_label.configure(text=f"[CMD] Sending sharing intent for '{file_name}'...",
                                    text_color=self.COLOR_ACCENT)

        device_path = f'/sdcard/Download/{file_name}'
        command = [
            'shell', 'am', 'start',
            '-a', 'android.intent.action.SEND',
            '-t', 'image/jpeg',
            '--eu', 'android.intent.extra.STREAM', f'file://{device_path}',
            TIKTOK_LITE_PACKAGE
        ]
        for device_serial in self.devices:
            self.executor.submit(run_adb_command, command, device_serial)
        self.status_label.configure(text="✅ Image sharing command sent to all devices.", text_color=self.COLOR_SUCCESS)

    def stop_all_commands(self):
        self.status_label.configure(text="⚠️ TERMINATING ALL ACTIVE COMMANDS...", text_color=self.COLOR_WARNING)
        is_stop_requested.set()

        self.stop_auto_type_loop()

        self.executor.shutdown(wait=True)

        self.executor = concurrent.futures.ThreadPoolExecutor(max_workers=multiprocessing.cpu_count() * 4)
        is_stop_requested.clear()

        self.status_label.configure(text="✅ ALL OPERATIONS TERMINATED. Ready.", text_color=self.COLOR_SUCCESS)


if __name__ == '__main__':
    multiprocessing.freeze_support()
    app = AdbControllerApp()
    app.mainloop()
