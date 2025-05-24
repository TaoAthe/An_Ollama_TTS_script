import customtkinter as ctk
import tkinter.messagebox as messagebox
import tkinter as tk
from PIL import Image, ImageTk
import threading
import queue
import json
import os
import io
import base64
import re
import time  # Import time module for sleep
import uuid
import warnings  # For handling FutureWarnings if necessary
import wave  # Import wave module for WAV file writing

# External Libraries
try:
    import ollama
except ImportError:
    messagebox.showerror(
        "Dependency Error",
        "The 'ollama' library is not installed. Please install it with: pip install ollama",
    )
    exit()

try:
    import pyaudio
except ImportError:
    messagebox.showerror(
        "Dependency Error",
        "The 'PyAudio' library is not installed. Please install it with: pip install pyaudio\n(Also ensure PortAudio is installed on your system).",
    )
    exit()

try:
    import pyttsx3
except ImportError:
    messagebox.showerror(
        "Dependency Error",
        "The 'pyttsx3' library is not installed. Please install it with: pip install pyttsx3\n(Also ensure espeak or other TTS backend is installed on your system).",
    )
    exit()

try:
    import keyboard
except ImportError:
    messagebox.showerror(
        "Dependency Error",
        "The 'keyboard' library is not installed. Please install it with: pip install keyboard",
    )
    exit()

# --- Whisper Integration ---
try:
    import whisper
    import numpy as np  # Whisper works well with numpy arrays

    WHISPER_AVAILABLE = True
except ImportError:
    messagebox.showwarning(
        "Dependency Warning",
        "Whisper or NumPy not found. Offline speech recognition will be disabled. Install with: pip install openai-whisper numpy",
    )
    WHISPER_AVAILABLE = False
# --- End Whisper Integration ---

try:
    from pygments import highlight
    from pygments.lexers import get_lexer_by_name, guess_lexer
    from pygments.formatters import (
        TerminalFormatter,
    )  # Not directly used for CTkTextbox, but good for testing
    from pygments.styles import get_style_by_name
    from pygments.token import Token

    PYGMENTS_AVAILABLE = True
except ImportError:
    messagebox.showwarning(
        "Dependency Warning",
        "Pygments not found. Code syntax highlighting will be disabled. Install with: pip install Pygments",
    )
    PYGMENTS_AVAILABLE = False


# Configuration File
CONFIG_FILE = "config.json"

# Default Settings
DEFAULT_SETTINGS = {
    "theme_mode": "System",
    "color_theme": "blue",
    "ollama_model": "llama2",
    "hotkey_str": "alt",
    "tts_voice_id": None,
    "tts_rate": 180,
    "tts_volume": 1.0,
    "font_size": 14,
    "whisper_model_name": "base",  # e.g., "tiny", "base", "small", "medium", "large"
    "selected_mic_index": -1,  # -1 means no specific mic selected, will try default or first available
}

# --- Pygments Style for CTkTextbox ---
# A simple mapping for a dark theme. Extend as needed.
# (foreground_dark_mode, foreground_light_mode)
PYGMENTS_STYLE_MAP = {
    Token.Keyword: ("#ff79c6", "#bd93f9"),  # pink, purple
    Token.Keyword.Constant: ("#ff79c6", "#bd93f9"),
    Token.Keyword.Declaration: ("#ff79c6", "#bd93f9"),
    Token.Keyword.Namespace: ("#ff79c6", "#bd93f9"),
    Token.Keyword.Pseudo: ("#ff79c6", "#bd93f9"),
    Token.Keyword.Reserved: ("#ff79c6", "#bd93f9"),
    Token.Keyword.Type: ("#ff79c6", "#bd93f9"),
    Token.Name.Builtin: ("#8be9fd", "#a4e6ff"),  # cyan
    Token.Name.Builtin.Pseudo: ("#8be9fd", "#a4e6ff"),
    Token.Name.Class: ("#50fa7b", "#77f29a"),  # green
    Token.Name.Function: ("#50fa7b", "#77f29a"),
    Token.Name.Exception: ("#ff5555", "#ff7a7a"),  # red
    Token.Name.Variable: ("#8be9fd", "#a4e6ff"),
    Token.Name.Constant: ("#bd93f9", "#e0caff"),  # purple
    Token.Name.Label: ("#f1fa8c", "#ffe9a4"),  # yellow
    Token.Literal.String: ("#f1fa8c", "#ffe9a4"),  # yellow
    Token.Literal.String.Doc: ("#f1fa8c", "#ffe9a4"),
    Token.Literal.Number: ("#bd93f9", "#e0caff"),  # purple
    Token.Operator: ("#ff79c6", "#bd93f9"),  # pink
    Token.Punctuation: ("#f8f8f2", "#282a36"),  # white/grey
    Token.Comment: ("#6272a4", "#959595"),  # grey/blue
    Token.Comment.Single: ("#6272a4", "#959595"),
    Token.Comment.Multiline: ("#6272a4", "#959595"),
    Token.Generic.Deleted: ("#ff5555", "#ff7a7a"),
    Token.Generic.Inserted: ("#50fa7b", "#77f29a"),
    Token.Generic.Heading: ("#ffb86c", "#ffd2a4"),
    Token.Generic.Subheading: ("#ffb86c", "#ffd2a4"),
    Token.Generic.Emph: ("#f8f8f2", "#282a36"),
    Token.Generic.Strong: ("#f8f8f2", "#282a36"),
    Token.Generic.Error: ("#ff5555", "#ff7a7a"),
    Token.Text: ("#f8f8f2", "#282a36"),  # default text color
    Token.Error: ("#ff5555", "#ff7a7a"),
}

# Global Whisper model reference
whisper_model = None


# --- NEW: Special Character Cleanup Tool ---
import re  # Ensure re is imported at the top of your file


def remove_special_characters(text_input):  # This is the function your TTS worker uses
    """
    Cleans text for TTS:
    1. Removes entire <think>...</think> blocks (tags and content).
    2. Removes any stray <think> or </think> tags.
    3. Cleans other non-essential special characters for smoother speech.
    4. Consolidates whitespace.
    """
    if not isinstance(text_input, str):
        return ""

    # 1. Remove entire <think>...</think> blocks (tags and their content)
    processed_text = re.sub(
        r"<\s*think\s*>.*?<\s*/\s*think\s*>",
        "",
        text_input,
        flags=re.IGNORECASE | re.DOTALL,
    )

    # 2. Remove any standalone <think> or </think> tags
    processed_text = re.sub(
        r"<\s*/?\s*think\s*>", "", processed_text, flags=re.IGNORECASE
    )

    # 3. Further cleanup of other special characters (keeping essential punctuation)
    processed_text = re.sub(r"[^\w\s.,!?\"”-]", "", processed_text)

    # 4. Consolidate multiple spaces and strip.
    processed_text = re.sub(r"\s+", " ", processed_text).strip()

    return processed_text


# --- END NEW ---


# --- Main Application ---
class OllamaSpeechChatApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Ollama Speech Chat")
        self.geometry("1000x750")
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(2, weight=1)

        self.conversation_history = []
        self.MAX_HISTORY_CHARS = 2000

        self.is_recording = False
        self.audio_frames = []
        self.recording_thread = None
        self.transcription_thread = None
        self.ollama_response_thread = None
        self.tts_thread = None
        self.tts_engine = None
        self.hotkey_pressed = False
        self.hotkey_listening_event = threading.Event()

        self.tts_text_queue = queue.Queue()
        self.stop_tts_event = threading.Event()

        self.code_window = None
        self.available_mics_info = []

        self.settings = DEFAULT_SETTINGS.copy()
        self._load_config()
        ctk.set_appearance_mode(self.settings["theme_mode"])
        ctk.set_default_color_theme(self.settings["color_theme"])

        self._create_widgets()
        self._init_tts_engine()
        self._start_tts_playback_thread()

        # Defer tasks that might interact with GUI early or use global hooks
        self.after(100, self._perform_initial_background_tasks)

        self.protocol("WM_DELETE_WINDOW", self._on_closing)

    # --- NEW: Make remove_special_characters a method of the class for easy access ---
    def remove_special_characters(self, text):
        return remove_special_characters(text)  # Calls the global function

    # --- END NEW ---

    def _perform_initial_background_tasks(self):
        """
        Performs initializations that should happen after the main loop is more stable,
        especially those involving threads that might schedule GUI updates or global hooks.
        """
        print("Performing initial background tasks...")
        self._fetch_ollama_models()  # This will start its own thread
        self._fetch_microphones()  # This will start its own thread
        self._load_whisper_model()  # This runs in main thread, GUI updates are safe via self.after
        self._initialize_hotkey_listener_safely()  # This was already correctly deferred

    def _initialize_hotkey_listener_safely(self):
        """Initializes the global hotkey listener."""
        print("Attempting to initialize hotkey listener safely...")
        self._start_hotkey_listener()

    def _load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r") as f:
                    loaded_settings = json.load(f)
                    self.settings.update(loaded_settings)
            except json.JSONDecodeError:
                messagebox.showwarning(
                    "Config Error",
                    f"Could not decode {CONFIG_FILE}. Using default settings.",
                )
        if not isinstance(self.settings.get("hotkey_str"), str):
            self.settings["hotkey_str"] = DEFAULT_SETTINGS["hotkey_str"]
        if "whisper_model_name" not in self.settings:
            self.settings["whisper_model_name"] = DEFAULT_SETTINGS["whisper_model_name"]
        if "selected_mic_index" not in self.settings:
            self.settings["selected_mic_index"] = DEFAULT_SETTINGS["selected_mic_index"]

    def _save_config(self):
        try:
            with open(CONFIG_FILE, "w") as f:
                json.dump(self.settings, f, indent=4)
        except IOError as e:
            messagebox.showerror("Config Error", f"Could not save configuration: {e}")

    def _create_widgets(self):
        # Sidebar Frame
        self.sidebar_frame = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=4, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(7, weight=0)
        self.sidebar_frame.grid_rowconfigure(24, weight=1)

        self.logo_label = ctk.CTkLabel(
            self.sidebar_frame,
            text="Ollama Chat",
            font=ctk.CTkFont(size=24, weight="bold"),
        )
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))

        self.theme_label = ctk.CTkLabel(self.sidebar_frame, text="Appearance Mode:")
        self.theme_label.grid(row=1, column=0, padx=20, pady=(10, 0), sticky="w")
        self.theme_optionmenu = ctk.CTkOptionMenu(
            self.sidebar_frame,
            values=["Light", "Dark", "System"],
            command=self._change_appearance_mode_event,
        )
        self.theme_optionmenu.grid(row=2, column=0, padx=20, pady=(0, 10), sticky="ew")
        self.theme_optionmenu.set(self.settings["theme_mode"])

        self.color_theme_label = ctk.CTkLabel(self.sidebar_frame, text="Color Theme:")
        self.color_theme_label.grid(row=3, column=0, padx=20, pady=(10, 0), sticky="w")
        self.color_theme_optionmenu = ctk.CTkOptionMenu(
            self.sidebar_frame,
            values=["blue", "green", "dark-blue"],
            command=self._change_color_theme_event,
        )
        self.color_theme_optionmenu.grid(
            row=4, column=0, padx=20, pady=(0, 10), sticky="ew"
        )
        self.color_theme_optionmenu.set(self.settings["color_theme"])

        self.font_size_label = ctk.CTkLabel(self.sidebar_frame, text="Font Size:")
        self.font_size_label.grid(row=5, column=0, padx=20, pady=(10, 0), sticky="w")
        self.font_size_slider = ctk.CTkSlider(
            self.sidebar_frame,
            from_=10,
            to=20,
            number_of_steps=10,
            command=self._update_font_size,
        )
        self.font_size_slider.grid(row=6, column=0, padx=20, pady=(0, 10), sticky="ew")
        self.font_size_slider.set(self.settings["font_size"])
        self.font_size_value_label = ctk.CTkLabel(
            self.sidebar_frame, text=str(int(self.settings["font_size"]))
        )
        self.font_size_value_label.grid(
            row=5, column=0, padx=(100, 20), pady=(10, 0), sticky="e"
        )

        self.model_label = ctk.CTkLabel(self.sidebar_frame, text="Select Ollama Model:")
        self.model_label.grid(row=8, column=0, padx=20, pady=(10, 0), sticky="w")
        self.model_combobox = ctk.CTkComboBox(
            self.sidebar_frame,
            values=["Loading models..."],
            command=self._select_ollama_model,
        )
        self.model_combobox.grid(row=9, column=0, padx=20, pady=(0, 10), sticky="ew")
        self.model_combobox.set("No models found")

        self.hotkey_label = ctk.CTkLabel(
            self.sidebar_frame, text="Push-to-Talk Hotkey:"
        )
        self.hotkey_label.grid(row=10, column=0, padx=20, pady=(10, 0), sticky="w")
        self.current_hotkey_label = ctk.CTkLabel(
            self.sidebar_frame, text=f"Current: {self.settings['hotkey_str']}"
        )
        self.current_hotkey_label.grid(
            row=11, column=0, padx=20, pady=(0, 5), sticky="w"
        )
        self.set_hotkey_button = ctk.CTkButton(
            self.sidebar_frame, text="Set Hotkey", command=self._set_hotkey_callback
        )
        self.set_hotkey_button.grid(
            row=12, column=0, padx=20, pady=(0, 10), sticky="ew"
        )

        self.tts_label = ctk.CTkLabel(self.sidebar_frame, text="TTS Settings:")
        self.tts_label.grid(row=13, column=0, padx=20, pady=(10, 0), sticky="w")

        self.tts_voice_label = ctk.CTkLabel(self.sidebar_frame, text="Voice:")
        self.tts_voice_label.grid(row=14, column=0, padx=20, pady=(5, 0), sticky="w")
        self.tts_voice_combobox = ctk.CTkComboBox(
            self.sidebar_frame,
            values=["System Default"],  # Changed default value
            command=self._update_tts_settings,
        )
        self.tts_voice_combobox.grid(
            row=15, column=0, padx=20, pady=(0, 5), sticky="ew"
        )
        self.tts_voice_combobox.set("System Default")  # Set initial display
        self.tts_voice_combobox.configure(state="disabled")  # Disable selection for now

        self.tts_rate_label = ctk.CTkLabel(self.sidebar_frame, text="Rate:")
        self.tts_rate_label.grid(row=16, column=0, padx=20, pady=(5, 0), sticky="w")
        self.tts_rate_slider = ctk.CTkSlider(
            self.sidebar_frame,
            from_=50,
            to=300,
            number_of_steps=25,
            command=self._update_tts_settings,
        )
        self.tts_rate_slider.grid(row=17, column=0, padx=20, pady=(0, 5), sticky="ew")
        self.tts_rate_slider.set(self.settings["tts_rate"])
        self.tts_rate_value_label = ctk.CTkLabel(
            self.sidebar_frame, text=str(int(self.settings["tts_rate"]))
        )
        self.tts_rate_value_label.grid(
            row=16, column=0, padx=(100, 20), pady=(5, 0), sticky="e"
        )

        self.tts_volume_label = ctk.CTkLabel(self.sidebar_frame, text="Volume:")
        self.tts_volume_label.grid(row=18, column=0, padx=20, pady=(5, 0), sticky="w")
        self.tts_volume_slider = ctk.CTkSlider(
            self.sidebar_frame,
            from_=0.0,
            to=1.0,
            number_of_steps=10,
            command=self._update_tts_settings,
        )
        self.tts_volume_slider.grid(
            row=19, column=0, padx=20, pady=(0, 10), sticky="ew"
        )
        self.tts_volume_slider.set(self.settings["tts_volume"])
        self.tts_volume_value_label = ctk.CTkLabel(
            self.sidebar_frame, text=f"{self.settings['tts_volume']:.1f}"
        )
        self.tts_volume_value_label.grid(
            row=18, column=0, padx=(100, 20), pady=(5, 0), sticky="e"
        )

        self.stop_speaking_button = ctk.CTkButton(
            self.sidebar_frame, text="Stop Speaking", command=self._stop_tts_playback
        )
        self.stop_speaking_button.grid(
            row=20, column=0, padx=20, pady=(10, 20), sticky="ew"
        )

        self.mic_label = ctk.CTkLabel(self.sidebar_frame, text="Select Microphone:")
        self.mic_label.grid(row=21, column=0, padx=20, pady=(10, 0), sticky="w")
        self.mic_combobox = ctk.CTkComboBox(
            self.sidebar_frame,
            values=["Loading microphones..."],
            command=self._select_microphone,
        )
        self.mic_combobox.grid(row=22, column=0, padx=20, pady=(0, 10), sticky="ew")
        self.mic_combobox.set("No microphones found")

        self.exit_button = ctk.CTkButton(
            self.sidebar_frame, text="Exit", command=self._on_closing
        )
        self.exit_button.grid(row=25, column=0, padx=20, pady=(10, 20), sticky="sew")

        self.main_frame = ctk.CTkFrame(self, corner_radius=0)
        self.main_frame.grid(
            row=0, column=1, rowspan=4, sticky="nsew", padx=10, pady=10
        )
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(2, weight=1)
        self.main_frame.grid_rowconfigure(4, weight=2)
        self.main_frame.grid_rowconfigure(5, weight=0)

        self.status_label = ctk.CTkLabel(
            self.main_frame,
            text="Press and hold 'Alt' to speak...",
            font=ctk.CTkFont(size=self.settings["font_size"] + 2, weight="bold"),
        )
        self.status_label.grid(row=0, column=0, padx=10, pady=(10, 5), sticky="ew")

        self.user_query_label = ctk.CTkLabel(
            self.main_frame,
            text="Your Query:",
            font=ctk.CTkFont(size=self.settings["font_size"], weight="bold"),
        )
        self.user_query_label.grid(row=1, column=0, padx=10, pady=(5, 0), sticky="w")
        self.user_query_textbox = ctk.CTkTextbox(
            self.main_frame, wrap="word", height=100
        )
        self.user_query_textbox.grid(
            row=2, column=0, padx=10, pady=(0, 10), sticky="nsew"
        )
        self.user_query_textbox.configure(
            font=ctk.CTkFont(size=self.settings["font_size"])
        )

        self.ollama_response_label = ctk.CTkLabel(
            self.main_frame,
            text="Ollama's Response:",
            font=ctk.CTkFont(size=self.settings["font_size"], weight="bold"),
        )
        self.ollama_response_label.grid(
            row=3, column=0, padx=10, pady=(5, 0), sticky="w"
        )
        self.ollama_response_textbox = ctk.CTkTextbox(self.main_frame, wrap="word")
        self.ollama_response_textbox.grid(
            row=4, column=0, padx=10, pady=(0, 10), sticky="nsew"
        )
        self.ollama_response_textbox.configure(
            font=ctk.CTkFont(size=self.settings["font_size"])
        )

        self.image_frame = ctk.CTkFrame(
            self.main_frame, corner_radius=8, border_width=2
        )
        self.image_label = ctk.CTkLabel(self.image_frame, text="")
        self.image_label.pack(padx=10, pady=5, expand=True, fill="both")
        self.image_caption_label = ctk.CTkLabel(
            self.image_frame,
            text="",
            wraplength=400,
            font=ctk.CTkFont(size=self.settings["font_size"] - 2, slant="italic"),
        )
        self.image_caption_label.pack(padx=10, pady=5, expand=True, fill="x")
        self.current_image = None

        self._hide_image_frame()
        self._update_all_fonts()

    def _update_all_fonts(self, *args):
        font_size = int(self.settings["font_size"])
        bold_font = ctk.CTkFont(size=font_size, weight="bold")
        normal_font = ctk.CTkFont(size=font_size)
        italic_font = ctk.CTkFont(size=font_size - 2, slant="italic")

        self.logo_label.configure(font=ctk.CTkFont(size=font_size + 10, weight="bold"))
        self.status_label.configure(font=ctk.CTkFont(size=font_size + 2, weight="bold"))
        self.user_query_label.configure(font=bold_font)
        self.ollama_response_label.configure(font=bold_font)
        self.user_query_textbox.configure(font=normal_font)
        self.ollama_response_textbox.configure(font=normal_font)
        self.image_caption_label.configure(font=italic_font)

        for widget in [
            self.theme_label,
            self.color_theme_label,
            self.font_size_label,
            self.font_size_value_label,
            self.model_label,
            self.hotkey_label,
            self.current_hotkey_label,
            self.tts_label,
            self.tts_voice_label,
            self.tts_rate_label,
            self.tts_rate_value_label,
            self.tts_volume_label,
            self.tts_volume_value_label,
            self.mic_label,
        ]:
            widget.configure(font=normal_font)
        for widget in [
            self.theme_optionmenu,
            self.color_theme_optionmenu,
            self.model_combobox,
            self.tts_voice_combobox,  # This will be disabled, but font still applies
            self.mic_combobox,
        ]:
            widget.configure(font=normal_font)
        for widget in [
            self.set_hotkey_button,
            self.stop_speaking_button,
            self.exit_button,
        ]:
            widget.configure(font=normal_font)

    def _update_font_size(self, value):
        self.settings["font_size"] = int(value)
        self.font_size_value_label.configure(text=str(int(value)))
        self._update_all_fonts()
        self._save_config()

    def _change_appearance_mode_event(self, new_appearance_mode: str):
        ctk.set_appearance_mode(new_appearance_mode)
        self.settings["theme_mode"] = new_appearance_mode
        self._save_config()
        if self.code_window and self.code_window.winfo_exists():
            print(
                "Appearance mode changed. Re-open code window for updated syntax highlighting if needed."
            )

    def _change_color_theme_event(self, new_color_theme: str):
        ctk.set_default_color_theme(new_color_theme)
        self.settings["color_theme"] = new_color_theme
        self._save_config()

    def _update_status_label(self, message, color="white"):
        self.after(
            0, lambda: self.status_label.configure(text=message, text_color=color)
        )

    def _show_error_message(self, title, message):
        self.after(0, lambda: messagebox.showerror(title, message))

    def _fetch_ollama_models(self):
        def _fetch():
            model_names = ["No models found"]
            selected_model = "No models found"
            status_color = "red"
            status_message = "No Ollama models found."
            error_msg = None
            try:
                self._update_status_label("Fetching Ollama models...", "orange")
                models_info = ollama.list()
                # print(f"Ollama list response: {models_info}") # Debug print: See the raw response

                extracted_model_names = []
                # Access the 'models' attribute of the response object, which is a list of Model objects
                # The ollama.list() returns a ListResponse object, which has a 'models' attribute
                if hasattr(models_info, "models") and isinstance(
                    models_info.models, list
                ):
                    for m_obj in models_info.models:
                        # Each item in models_info.models is an ollama.Model object
                        # The model name is in the 'model' attribute of this object
                        if hasattr(m_obj, "model"):  # Check for 'model' attribute
                            model_name_val = (
                                m_obj.model
                            )  # CORRECTED: Access .model attribute
                            if model_name_val:
                                extracted_model_names.append(model_name_val)

                model_names = extracted_model_names

                if not model_names:
                    error_msg = "No Ollama models found. Download one (e.g., 'ollama run llama2') and ensure Ollama server is running."
                else:
                    if self.settings["ollama_model"] in model_names:
                        selected_model = self.settings["ollama_model"]
                    else:
                        selected_model = model_names[0]
                        self.settings["ollama_model"] = model_names[0]
                        self._save_config()
                    status_message = f"Ready. Selected: {selected_model}"
                    status_color = "green"
            except ollama.ResponseError as e:
                error_msg = f"Error connecting to Ollama server: {e}\nPlease ensure Ollama server is running."
                status_message = "Ollama server connection failed."
                status_color = "red"
            except Exception as e:
                error_msg = f"Unexpected error fetching Ollama models: {e}"
                status_message = "Failed to load models."
                status_color = "red"
            finally:
                self.after(
                    0,
                    self._complete_ollama_models_fetch_gui_update,
                    model_names,
                    selected_model,
                    status_message,
                    status_color,
                    error_msg,
                )

        threading.Thread(target=_fetch, daemon=True).start()

    def _complete_ollama_models_fetch_gui_update(
        self, model_names, selected_model, status_message, status_color, error_msg
    ):
        self.model_combobox.configure(
            values=model_names if model_names else ["No models found"]
        )
        self.model_combobox.set(selected_model)
        self._update_status_label(status_message, status_color)
        if error_msg:
            self._show_error_message("Ollama Error", error_msg)

    def _select_ollama_model(self, model_name):
        self.settings["ollama_model"] = model_name
        self._save_config()
        self._update_status_label(f"Model selected: {model_name}", "green")

    def _send_query_to_ollama_threaded(self, query):
        if (
            not self.settings["ollama_model"]
            or self.settings["ollama_model"] == "No models found"
        ):
            self._show_error_message("Ollama Error", "No Ollama model selected.")
            self._update_status_label("No model selected. Cannot converse.", "red")
            return

        def _send():
            self._update_status_label("Thinking...", "orange")
            self._hide_image_frame()
            self._destroy_code_window()
            self.after(0, lambda: self.ollama_response_textbox.delete("1.0", "end"))

            # Corrected: Empty the queue safely
            while not self.tts_text_queue.empty():
                try:
                    self.tts_text_queue.get_nowait()
                except queue.Empty:
                    break
            print("TTS queue cleared before new response.")  # Debug print

            # Crucial: Clear the stop event so the TTS worker can start processing new data
            self.stop_tts_event.clear()
            print("TTS stop event cleared.")  # Debug print

            self.conversation_history.append({"role": "user", "content": query})
            messages_for_ollama = []
            current_char_count = 0
            messages_for_ollama.append({"role": "user", "content": query})
            current_char_count += len(query)
            temp_prev_messages = []
            for i in range(len(self.conversation_history) - 2, -1, -1):
                msg = self.conversation_history[i]
                msg_len = len(msg["content"])
                if current_char_count + msg_len + 50 < self.MAX_HISTORY_CHARS:
                    temp_prev_messages.append(msg)
                    current_char_count += msg_len
                else:
                    break
            messages_for_ollama = (
                list(reversed(temp_prev_messages)) + messages_for_ollama
            )
            full_response_content = ""
            current_code_block = ""
            in_code_block = False
            language_for_code_block = "text"
            try:
                # self.tts_text_queue.put("This is a test sentence.") # Removed test sentence
                response_generator = ollama.chat(
                    model=self.settings["ollama_model"],
                    messages=messages_for_ollama,
                    stream=True,
                )
                for chunk in response_generator:
                    if chunk["message"]["content"]:
                        content_chunk = chunk["message"]["content"]
                        full_response_content += content_chunk
                        self.after(
                            0,
                            lambda c=content_chunk: self.ollama_response_textbox.insert(
                                "end", c
                            ),
                        )
                        self.after(0, self.ollama_response_textbox.see("end"))
                        self.tts_text_queue.put(content_chunk)

                        temp_buffer = current_code_block + content_chunk
                        parts = temp_buffer.split("```")
                        processed_up_to_content_chunk = 0
                        for i, part in enumerate(parts):
                            if i == 0 and not in_code_block:
                                processed_up_to_content_chunk += len(part)
                                continue
                            if not in_code_block:
                                in_code_block = True
                                lang_match = re.match(r"^\s*([a-zA-Z0-9]+)\s*\n", part)
                                if lang_match:
                                    language_for_code_block = lang_match.group(
                                        1
                                    ).lower()
                                    current_code_block = part[
                                        len(lang_match.group(0)) :
                                    ]
                                else:
                                    language_for_code_block = "text"
                                    current_code_block = part
                                processed_up_to_content_chunk += len("```") + len(part)
                            else:
                                if (
                                    i == len(parts) - 1
                                    and "```"
                                    not in content_chunk[processed_up_to_content_chunk:]
                                ):
                                    current_code_block += part
                                    processed_up_to_content_chunk += len(part)
                                else:
                                    current_code_block += part
                                    self.after(
                                        0,
                                        lambda code=current_code_block, lang=language_for_code_block: self._show_code_window(
                                            code, lang
                                        ),
                                    )
                                    current_code_block = ""
                                    in_code_block = False
                                    language_for_code_block = "text"
                                    processed_up_to_content_chunk += len(part) + len(
                                        "```"
                                    )
                        if in_code_block and len(parts) % 2 == 0:
                            pass
                        elif not in_code_block:
                            current_code_block = ""
                if in_code_block and current_code_block.strip():
                    self.after(
                        0,
                        lambda code=current_code_block, lang=language_for_code_block: self._show_code_window(
                            code, lang
                        ),
                    )
                try:
                    json_data_match = re.search(
                        r"\{.*?\"type\":\s*\"text/image\".*?\}",
                        full_response_content,
                        re.DOTALL | re.IGNORECASE,
                    )
                    if json_data_match:
                        json_str = json_data_match.group(0)
                        first_brace = json_str.find("{")
                        last_brace = json_str.rfind("}")
                        if (
                            first_brace != -1
                            and last_brace != -1
                            and last_brace > first_brace
                        ):
                            json_str_cleaned = json_str[first_brace : last_brace + 1]
                            try:
                                response_obj = json.loads(json_str_cleaned)
                                if (
                                    isinstance(response_obj, dict)
                                    and response_obj.get("type", "").lower()
                                    == "text/image"
                                    and "data" in response_obj
                                ):
                                    image_data = response_obj["data"]
                                    image_caption = response_obj.get("caption", "")
                                    self.after(
                                        0,
                                        lambda: self._show_image_frame(
                                            image_data, image_caption
                                        ),
                                    )
                            except json.JSONDecodeError as jde_inner:
                                print(
                                    f"JSON Decode Error for potential image: {jde_inner} on string: {json_str_cleaned}"
                                )
                except Exception as e:
                    print(f"Error processing potential image JSON: {e}")
                self.conversation_history.append(
                    {"role": "assistant", "content": full_response_content}
                )
                print(
                    f"Ollama full response content: '{full_response_content[:100]}...'"
                )  # Debug print
                self._update_status_label(
                    f"Ready. Press and hold '{self.settings['hotkey_str']}' to speak...",
                    "green",
                )
            except ollama.ResponseError as e:
                self._show_error_message("Ollama Error", f"Model response error: {e}")
                self._update_status_label("Ollama error.", "red")
                if (
                    self.conversation_history
                    and self.conversation_history[-1]["role"] == "user"
                ):
                    self.conversation_history.pop()
            except Exception as e:
                self._show_error_message(
                    "Ollama Error", f"Unexpected error during Ollama interaction: {e}"
                )
                self._update_status_label("Ollama error.", "red")
                if (
                    self.conversation_history
                    and self.conversation_history[-1]["role"] == "user"
                ):
                    self.conversation_history.pop()

        self.ollama_response_thread = threading.Thread(target=_send, daemon=True)
        self.ollama_response_thread.start()

    def _load_whisper_model(self):
        global whisper_model
        if not WHISPER_AVAILABLE:
            self._update_status_label(
                "Whisper not installed. Offline SR disabled.", "red"
            )
            return False
        if whisper_model is None:
            model_name = self.settings.get("whisper_model_name", "base")
            print(f"Attempting to load Whisper model: {model_name}")
            try:
                self._update_status_label(
                    f"Loading Whisper model ({model_name})...", "orange"
                )
                # The FutureWarning from Whisper regarding torch.load(weights_only=False)
                # originates from within whisper.load_model. We cannot pass weights_only=True
                # to it directly. This warning is for developers of libraries using torch.load
                # or users loading untrusted models. It's safe to ignore for now with official Whisper models.
                with warnings.catch_warnings():  # Context manager to temporarily ignore specific warning
                    warnings.filterwarnings(
                        "ignore", category=FutureWarning, module="torch.serialization"
                    )
                    whisper_model = whisper.load_model(model_name)
                self._update_status_label("Whisper model loaded.", "green")
                print(f"Whisper model '{model_name}' loaded successfully.")
                return True
            except Exception as e:
                self._show_error_message(
                    "Whisper Model Error",
                    f"Failed to load Whisper model '{model_name}': {e}\nEnsure model name is correct, internet for first download, and PyTorch is installed.",
                )
                self._update_status_label("Whisper model load failed.", "red")
                whisper_model = None
                return False
        print("Whisper model already loaded.")
        return True

    def _start_recording(self):
        print("Attempting to start recording...")  # Debug print
        if self.is_recording:
            print("Already recording, ignoring start request.")  # Debug print
            return
        if not WHISPER_AVAILABLE or whisper_model is None:
            self._show_error_message(
                "Speech Recognition Error", "Whisper not available/loaded."
            )
            self._update_status_label("Offline SR not ready.", "red")
            print("Whisper not ready, cannot start recording.")  # Debug print
            return
        if self.settings.get("selected_mic_index", -1) == -1:
            self._show_error_message("Microphone Error", "No microphone selected.")
            self._update_status_label("No microphone selected.", "red")
            print("No microphone selected, cannot start recording.")  # Debug print
            return
        self.is_recording = True
        self.audio_frames = []
        self._update_status_label(
            f"Recording... Release '{self.settings['hotkey_str']}' to stop.", "red"
        )
        print("Recording initiated.")  # Debug print
        self.recording_thread = threading.Thread(
            target=self._record_audio_threaded, daemon=True
        )
        self.recording_thread.start()

    def _stop_recording(self):
        print("Attempting to stop recording...")  # Debug print
        if not self.is_recording:
            print("Not currently recording, ignoring stop request.")  # Debug print
            return
        self.is_recording = False
        self._update_status_label("Processing speech...", "orange")
        if self.recording_thread and self.recording_thread.is_alive():
            print("Waiting for recording thread to finish...")  # Debug print
            self.recording_thread.join(timeout=2.0)  # Increased timeout slightly
            if self.recording_thread.is_alive():
                print(
                    "Warning: Recording thread did not finish in time."
                )  # Debug print
        else:
            print("Recording thread was not active or already finished.")  # Debug print

        print("Recording stopped. Starting transcription.")  # Debug print
        self.transcription_thread = threading.Thread(
            target=self._transcribe_audio_threaded, daemon=True
        )
        self.transcription_thread.start()

    def _record_audio_threaded(self):
        CHUNK, FORMAT, CHANNELS, RATE = 1024, pyaudio.paInt16, 1, 16000
        p, stream = None, None
        input_device_index = self.settings.get("selected_mic_index", -1)
        print(
            f"Recording thread started. Using microphone index: {input_device_index}"
        )  # Debug print
        try:
            p = pyaudio.PyAudio()
            # If no specific mic selected, try to find default input device
            if input_device_index == -1:
                try:
                    default_input_device = p.get_default_input_device_info()
                    input_device_index = default_input_device["index"]
                    self.settings["selected_mic_index"] = (
                        input_device_index  # Save it for next time
                    )
                    self._save_config()
                    # Update GUI for mic selection on main thread
                    self.after(
                        0, lambda: self.mic_combobox.set(default_input_device["name"])
                    )
                    print(
                        f"No specific mic selected, defaulting to: {default_input_device['name']} (index: {input_device_index})"
                    )  # Debug print
                except Exception as e:
                    self.after(
                        0,
                        lambda: self._show_error_message(
                            "Microphone Error",
                            f"No default microphone found. Please select one manually from the sidebar. Error: {e}",
                        ),
                    )
                    self.after(
                        0,
                        lambda: self._update_status_label(
                            "No microphone selected.", "red"
                        ),
                    )
                    print(f"Failed to find default microphone: {e}")  # Debug print
                    return

            stream = p.open(
                format=FORMAT,
                channels=CHANNELS,
                rate=RATE,
                input=True,
                frames_per_buffer=CHUNK,
                input_device_index=input_device_index,
            )
            print(
                f"Microphone stream opened successfully for index {input_device_index}."
            )  # Debug print
            print(
                f"Initial audio_frames length: {len(self.audio_frames)}"
            )  # Debug print

            frames_read = 0
            while self.is_recording:
                try:
                    data = stream.read(CHUNK, exception_on_overflow=False)
                    self.audio_frames.append(data)
                    frames_read += 1
                    if frames_read % 10 == 0:  # Print every 10 chunks to avoid spamming
                        print(
                            f"  Recording: Captured {frames_read} chunks. Last chunk size: {len(data)} bytes. Total frames: {len(self.audio_frames)}"
                        )  # Debug print
                        # Optional: Add a subtle visual pulse to the status label
                        self.after(
                            0,
                            lambda: self.status_label.configure(
                                text="Recording... (Active)",
                                text_color=(
                                    "dark red"
                                    if self.status_label.cget("text_color") == "red"
                                    else "red"
                                ),
                            ),
                        )
                    if len(data) == 0:
                        print(
                            "Warning: Received empty audio data chunk from stream. Microphone might be disconnected or faulty."
                        )  # Debug print

                except IOError as e:
                    print(f"IOError during recording: {e}")
                    self.is_recording = False
                    self.after(
                        0,
                        lambda e=e: self._show_error_message(
                            "Mic Error", f"Audio error: {e}"
                        ),
                    )
                    self.after(
                        0, lambda: self._update_status_label("Recording error.", "red")
                    )
                    break
        except Exception as e:
            self.is_recording = False
            self.after(
                0,
                lambda e=e: self._show_error_message(
                    "Mic Error",
                    f"Could not open mic (index {input_device_index}): {e}\nPlease check your microphone connection, OS permissions, and PyAudio installation.",
                ),
            )
            self.after(
                0,
                lambda: self._update_status_label(
                    "Microphone error. Try restarting.", "red"
                ),
            )
            print(f"Failed to open microphone stream: {e}")  # Debug print
        finally:
            if stream:
                print("Closing audio stream.")  # Debug print
                stream.stop_stream()
                stream.close()
            if p:
                p.terminate()
                print("Terminating PyAudio.")  # Debug print

            # --- NEW: Add a small delay after PyAudio terminates ---
            time.sleep(1.0)  # Increased delay to 1000ms
            print("Added 1000ms delay after PyAudio termination.")  # Debug print
            # --- END NEW ---

            print(
                f"Recording thread finished. Final audio_frames length: {len(self.audio_frames)}"
            )  # Debug print

    def _transcribe_audio_threaded(self):
        print("Transcription thread started.")  # Debug print
        print(
            f"Audio frames available for transcription: {len(self.audio_frames)}"
        )  # Debug print
        if not self.audio_frames:
            self.after(
                0, lambda: self._update_status_label("No audio recorded.", "yellow")
            )
            self.after(
                0,
                lambda: self._update_status_label(
                    f"Ready. Press and hold '{self.settings['hotkey_str']}' to speak...",
                    "green",
                ),
            )
            print("No audio frames to transcribe.")  # Debug print
            return
        if not WHISPER_AVAILABLE or whisper_model is None:
            self.after(
                0,
                lambda: self._show_error_message(
                    "SR Error", "Whisper not available/loaded."
                ),
            )
            self.after(
                0, lambda: self._update_status_label("Offline SR not ready.", "red")
            )
            print("Whisper not ready, cannot transcribe.")  # Debug print
            return

        audio_data_bytes = b"".join(self.audio_frames)
        self.audio_frames = []  # Clear frames after joining
        print(
            f"Total audio data length for transcription: {len(audio_data_bytes)} bytes."
        )  # Debug print
        if len(audio_data_bytes) == 0:
            self.after(
                0,
                lambda: self._update_status_label(
                    "Recorded silence or no audio.", "yellow"
                ),
            )
            self.after(0, lambda: self.user_query_textbox.delete("1.0", "end"))
            self.after(
                0,
                lambda: self.user_query_textbox.insert(
                    "end", "Recorded silence or no audio."
                ),
            )
            self.after(
                0,
                lambda: self._update_status_label(
                    f"Ready. Press and hold '{self.settings['hotkey_str']}' to speak...",
                    "green",
                ),
            )
            print("Recorded audio data is empty.")  # Debug print
            return

        # --- Save audio to WAV for debugging ---
        wav_filename = "recorded_audio.wav"
        try:
            with wave.open(wav_filename, "wb") as wf:
                wf.setnchannels(1)  # Mono
                wf.setsampwidth(pyaudio.get_sample_size(pyaudio.paInt16))  # 16-bit
                wf.setframerate(16000)  # 16kHz
                wf.writeframes(audio_data_bytes)
            print(f"Recorded audio saved to {wav_filename}")
            self.after(
                0,
                lambda: self._update_status_label(
                    f"Recorded audio saved to {wav_filename}. Processing...", "orange"
                ),
            )
        except Exception as e:
            print(f"Error saving WAV file: {e}")
            self.after(
                0,
                lambda: self._show_error_message(
                    "Audio Save Error", f"Could not save recorded audio to WAV: {e}"
                ),
            )
        # --- END WAV SAVE ---

        try:
            audio_np = (
                np.frombuffer(audio_data_bytes, dtype=np.int16).astype(np.float32)
                / 32768.0
            )
            self._update_status_label("Transcribing with Whisper...", "orange")
            # Suppress the specific FutureWarning from Whisper's internal torch.load here as well
            with warnings.catch_warnings():
                warnings.filterwarnings(
                    "ignore", category=FutureWarning, module="torch.serialization"
                )
                result = whisper_model.transcribe(
                    audio_np, fp16=False
                )  # fp16=False for CPU
            text = result["text"].strip()
            print(
                f"Whisper raw result: {result}"
            )  # Debug print: See the full Whisper result
            print(f"Whisper transcription: '{text}'")  # Debug print
            if text:
                self.after(
                    0,
                    lambda t=text: (
                        self.user_query_textbox.delete("1.0", "end"),
                        self.user_query_textbox.insert("end", t),
                    ),
                )
                self.after(0, lambda t=text: self._send_query_to_ollama_threaded(t))
            else:
                self.after(
                    0,
                    lambda: self._update_status_label(
                        "Could not understand (Whisper: no speech detected or unclear)",
                        "yellow",
                    ),
                )
                self.after(
                    0,
                    lambda: self.user_query_textbox.insert(
                        "end", "Could not understand (Whisper)."
                    ),
                )
                self.after(
                    0,
                    lambda: self._update_status_label(
                        f"Ready. Press and hold '{self.settings['hotkey_str']}' to speak...",
                        "green",
                    ),
                )
        except Exception as e:
            self.after(
                0,
                lambda e=e: self._show_error_message(
                    "Whisper Error",
                    f"Transcription error: {e}\nEnsure Whisper model is correctly loaded and audio format is compatible.",
                ),
            )
            self.after(
                0,
                lambda: self._update_status_label(
                    "Whisper transcription error.", "red"
                ),
            )
            self.after(
                0,
                lambda: self._update_status_label(
                    f"Ready. Press and hold '{self.settings['hotkey_str']}' to speak...",
                    "green",
                ),
            )
            print(f"Whisper transcription error: {e}")  # Debug print
        finally:
            print("Transcription thread finished.")

    def _start_hotkey_listener(self):
        try:
            keyboard.unhook_all()
        except Exception:
            pass
        hotkey_to_bind = self.settings["hotkey_str"]
        print(f"Attempting to bind hotkey: '{hotkey_to_bind}'")  # Debug print
        if not hotkey_to_bind or hotkey_to_bind.lower() == "none":
            self._update_status_label("No hotkey set.", "yellow")
            print("No hotkey configured.")  # Debug print
            return
        try:
            keyboard.on_press_key(hotkey_to_bind, self._on_hotkey_press, suppress=True)
            keyboard.on_release_key(
                hotkey_to_bind, self._on_hotkey_release, suppress=True
            )
            self._update_status_label(
                f"Ready. Hold '{hotkey_to_bind}' to speak...", "green"
            )
            self.after(
                0,
                lambda: self.current_hotkey_label.configure(
                    text=f"Current: {hotkey_to_bind}"
                ),
            )
            print(f"Hotkey '{hotkey_to_bind}' bound successfully.")  # Debug print
        except Exception as e:
            self._show_error_message(
                "Hotkey Error",
                f"Failed to set hotkey '{hotkey_to_bind}': {e}\nOn macOS, ensure Accessibility permissions are granted for your terminal/IDE or Python app. On Linux, you might need 'sudo' or specific X server configuration.",
            )
            self._update_status_label("Hotkey setup failed.", "red")
            print(f"Hotkey binding failed: {e}")  # Debug print

    def _on_hotkey_press(self, event):
        if not self.hotkey_pressed:
            self.hotkey_pressed = True
            # The tts_engine instance is now local to the worker thread.
            # _stop_tts_playback will signal the worker to stop its engine.
            print(
                "Hotkey pressed. Signaling TTS to stop if active, then starting recording."
            )
            self._stop_tts_playback()  # Signals the worker thread
            self.after(0, self._start_recording)  # Start recording after a brief moment

    def _on_hotkey_release(self, event):

        if self.hotkey_pressed:
            self.hotkey_pressed = False
            self.after(0, self._stop_recording)

    def _set_hotkey_callback(self):
        self._update_status_label("Press key combination for new hotkey...", "orange")
        self.set_hotkey_button.configure(state="disabled")
        print("Listening for new hotkey...")  # Debug print
        threading.Thread(target=self._listen_for_hotkey, daemon=True).start()

    def _listen_for_hotkey(self):
        try:
            keyboard.unhook_all()
            recorded_hotkey = keyboard.read_hotkey(suppress=False)
            new_hotkey_string = (
                str(recorded_hotkey)
                .lower()
                .replace("right ", "")
                .replace("left ", "")
                .replace("control", "ctrl")
            )
            if not new_hotkey_string:
                raise ValueError("No hotkey detected.")
            self.settings["hotkey_str"] = new_hotkey_string
            self._save_config()
            self.after(0, self._start_hotkey_listener)
            self.after(0, lambda: self.set_hotkey_button.configure(state="normal"))
            self.after(
                0,
                lambda s=new_hotkey_string: self.current_hotkey_label.configure(
                    text=f"Current: {s}"
                ),
            )
            print(f"New hotkey detected and set: '{new_hotkey_string}'")  # Debug print
        except Exception as e:
            self._show_error_message("Hotkey Error", f"Failed to set hotkey: {e}")
            self.after(0, lambda: self.set_hotkey_button.configure(state="normal"))
            self.after(0, self._start_hotkey_listener)  # Restore previous listener
            print(f"Error setting new hotkey: {e}")  # Debug print

    def _init_tts_engine(self):
        # This function now primarily sets up GUI elements based on settings,
        # and no longer initializes the pyttsx3 engine directly here.
        # The engine will be initialized in the _tts_playback_worker thread.
        self.tts_engine = (
            None  # No longer a pyttsx3 engine instance at class level for now
        )

        # Configure GUI elements based on settings
        self.after(
            0,
            lambda: self.tts_voice_combobox.configure(values=["System Default"]),
        )
        self.after(0, lambda: self.tts_voice_combobox.set("System Default"))
        self.after(0, lambda: self.tts_voice_combobox.configure(state="disabled"))

        self.after(0, lambda: self.tts_rate_slider.set(self.settings["tts_rate"]))
        self.after(0, lambda: self.tts_volume_slider.set(self.settings["tts_volume"]))

        self.after(
            0,
            lambda: self.tts_rate_value_label.configure(
                text=str(int(self.settings["tts_rate"]))
            ),
        )
        self.after(
            0,
            lambda: self.tts_volume_value_label.configure(
                text=f"{self.settings['tts_volume']:.1f}"
            ),
        )
        print("TTS GUI configured. Engine will be initialized in its worker thread.")

    def _update_tts_settings(self, *args):
        # This function now only updates self.settings.
        # The _tts_playback_worker thread will apply these to its engine instance.
        new_rate = int(self.tts_rate_slider.get())
        new_volume = float(self.tts_volume_slider.get())

        self.settings.update({"tts_rate": new_rate, "tts_volume": new_volume})

        self.tts_rate_value_label.configure(text=str(new_rate))
        self.tts_volume_value_label.configure(text=f"{new_volume:.1f}")
        self._save_config()
        print(
            f"TTS settings updated in config: Rate={new_rate}, Volume={new_volume}. Worker will apply."
        )

    def _start_tts_playback_thread(self):
        self.tts_thread = threading.Thread(
            target=self._tts_playback_worker, daemon=True
        )
        self.tts_thread.start()
        print("TTS playback thread started.")  # Debug print

    def _tts_playback_worker(self):
        engine = None  # Local engine instance for this thread
        try:
            print("TTS Playback Worker: Initializing pyttsx3 engine...")
            engine = pyttsx3.init()
            if not engine:
                print(
                    "TTS Playback Worker: Failed to initialize pyttsx3 engine. Thread stopping."
                )
                return  # Cannot proceed without an engine

            print("TTS Playback Worker: pyttsx3 engine initialized successfully.")
            # Initial properties based on app settings
            engine.setProperty("rate", self.settings["tts_rate"])
            engine.setProperty("volume", self.settings["tts_volume"])
            # Voice selection is using system default as per previous changes

        except Exception as e:
            print(f"TTS Playback Worker: Critical error initializing pyttsx3: {e}")
            # Optionally notify the main thread/GUI about this failure
            self.after(
                0,
                lambda: self._show_error_message(
                    "TTS Critical Error",
                    f"Failed to initialize TTS engine in its thread: {e}",
                ),
            )
            return  # Stop the worker if engine init fails

        current_sentence_buffer = ""
        last_known_settings = self.settings.copy()  # To detect changes

        while True:
            try:
                # Check for settings changes (rate, volume)
                if (
                    self.settings["tts_rate"] != last_known_settings["tts_rate"]
                    or self.settings["tts_volume"] != last_known_settings["tts_volume"]
                ):

                    print("TTS Playback Worker: Detected settings change. Applying...")
                    engine.setProperty("rate", self.settings["tts_rate"])
                    engine.setProperty("volume", self.settings["tts_volume"])
                    last_known_settings["tts_rate"] = self.settings["tts_rate"]
                    last_known_settings["tts_volume"] = self.settings["tts_volume"]
                    print(
                        f"TTS Playback Worker: Rate set to {self.settings['tts_rate']}, Volume to {self.settings['tts_volume']}"
                    )

                if self.stop_tts_event.is_set():
                    print("TTS Playback Worker: Stop signal received.")
                    self.stop_tts_event.clear()
                    if engine.isBusy():
                        print(
                            "TTS Playback Worker: Engine is busy, stopping current speech."
                        )
                        engine.stop()

                    # Clear the queue
                    while not self.tts_text_queue.empty():
                        try:
                            self.tts_text_queue.get_nowait()
                        except queue.Empty:
                            break
                    print("TTS Playback Worker: Text queue cleared.")
                    current_sentence_buffer = ""
                    time.sleep(0.1)  # Brief pause after stop
                    continue

                if self.tts_text_queue.empty() and not current_sentence_buffer:
                    time.sleep(0.05)
                    continue

                while not self.tts_text_queue.empty():
                    chunk = self.tts_text_queue.get()
                    current_sentence_buffer += chunk

                if not current_sentence_buffer:
                    time.sleep(0.05)
                    continue

                # Sentence splitting logic (simplified for brevity, use your existing refined logic)
                # Make sure to use re.split(r'(?<=[.!?\"”])\s+|\n\n+', current_sentence_buffer)
                sentences_to_process = []
                if (
                    re.search(r"[.!?\"”\n]\s*$", current_sentence_buffer)
                    or "\n\n" in current_sentence_buffer
                ):  # Basic check for end of potential sentence
                    split_parts = re.split(
                        r"(?<=[.!?\"”])\s+|\n\n+", current_sentence_buffer
                    )

                    # This logic needs to be robust as in your original code to handle partial sentences
                    if (
                        len(split_parts) > 1 and split_parts[-1] == ""
                    ):  # Ends with delimiter
                        sentences_to_process = [
                            s for s in split_parts[:-1] if s.strip()
                        ]
                        current_sentence_buffer = ""
                    elif (
                        len(split_parts) == 1 and split_parts[0].strip()
                    ):  # Single complete sentence
                        sentences_to_process = [split_parts[0].strip()]
                        current_sentence_buffer = ""
                    elif len(split_parts) > 1:  # Last part is incomplete
                        sentences_to_process = [
                            s for s in split_parts[:-1] if s.strip()
                        ]
                        current_sentence_buffer = split_parts[-1]
                    else:  # No complete sentence yet or only partial
                        pass  # Keep buffering

                for text_to_speak_now in sentences_to_process:
                    if (
                        text_to_speak_now and not self.stop_tts_event.is_set()
                    ):  # Re-check stop event
                        cleaned_text_to_speak = self.remove_special_characters(
                            text_to_speak_now
                        )  # Uses the class method
                        print(
                            f"TTS Playback Worker: Speaking: '{cleaned_text_to_speak[:100]}...'"
                        )
                        if cleaned_text_to_speak:
                            engine.say(cleaned_text_to_speak)
                            engine.runAndWait()  # This blocks this worker thread, not the GUI
                        if (
                            engine.isBusy()
                        ):  # Should not be busy after runAndWait, but good for sanity
                            print(
                                "TTS Playback Worker: WARNING - Engine still busy after runAndWait."
                            )
                            engine.stop()  # Force stop if stuck
                    else:
                        # If stop event was set during sentence processing
                        break

                if (
                    self.stop_tts_event.is_set()
                ):  # Handle stop if it occurred during sentence loop
                    print(
                        "TTS Playback Worker: Stop signal handled during sentence processing."
                    )
                    # (Stop event is cleared and queue is cleared at the beginning of the main loop)
                    current_sentence_buffer = ""  # Clear buffer as well
                    continue

                # If buffer is too large without forming sentences (e.g., no punctuation), speak it to prevent overflow
                if len(current_sentence_buffer) > 500:  # Arbitrary threshold
                    print(
                        f"TTS Playback Worker: Buffer limit reached, speaking: '{current_sentence_buffer[:100]}...'"
                    )
                    cleaned_text_to_speak = self.remove_special_characters(
                        current_sentence_buffer
                    )
                    if cleaned_text_to_speak:
                        engine.say(cleaned_text_to_speak)
                        engine.runAndWait()
                    current_sentence_buffer = ""

            except Exception as e:
                print(f"TTS Playback Worker: Error during playback loop: {e}")
                # Decide if the error is fatal for this worker
                current_sentence_buffer = ""  # Clear buffer on error
                time.sleep(0.1)  # Avoid rapid error loops

        # This part of the code (cleanup) will only be reached if the while loop exits,
        # which it currently doesn't. For a clean shutdown, you'd need a way to signal
        # this thread to exit its loop. The daemon=True helps, but explicit cleanup is better.
        # If you add an exit condition to the while loop:
        # print("TTS Playback Worker: Exiting and cleaning up engine.")
        # if engine and engine._inLoop: # Check if engine is still running its own loop
        #     engine.endLoop() # Should not be necessary with runAndWait
        # engine = None

    def _stop_tts_playback(self):
        # This method signals the _tts_playback_worker to stop.
        self.stop_tts_event.set()

        # It's also good to clear the queue here so the worker doesn't
        # immediately pick up old items after stopping and before the event is processed.
        cleared_count = 0
        while not self.tts_text_queue.empty():
            try:
                self.tts_text_queue.get_nowait()
                cleared_count += 1
            except queue.Empty:
                break
        if cleared_count > 0:
            print(f"Stop TTS Playback: Cleared {cleared_count} items from TTS queue.")

        print("Stop TTS Playback: Stop event set for TTS worker.")

    def _fetch_microphones(self):
        def _fetch():
            mic_names, selected_mic_name, error_msg = ["No mics"], "No mics", None
            temp_mics_info = []
            p = None
            try:
                p = pyaudio.PyAudio()
                for i in range(p.get_device_count()):
                    dev_info = p.get_device_info_by_index(i)
                    if dev_info.get("maxInputChannels") > 0:
                        temp_mics_info.append(
                            {"name": dev_info.get("name"), "index": i}
                        )
                mic_names = [mic["name"] for mic in temp_mics_info]
                if not mic_names:
                    error_msg = "No input microphones found."
                    self.settings["selected_mic_index"] = -1
                else:
                    idx = self.settings.get("selected_mic_index", -1)
                    found = any(mic["index"] == idx for mic in temp_mics_info)
                    if found:
                        selected_mic_name = next(
                            mic["name"] for mic in temp_mics_info if mic["index"] == idx
                        )
                    else:
                        selected_mic_name = mic_names[0]
                        self.settings["selected_mic_index"] = temp_mics_info[0]["index"]
            except Exception as e:
                error_msg = f"Failed to list mics: {e}"
                self.settings["selected_mic_index"] = -1
            finally:
                if p:
                    p.terminate()
                self.after(
                    0,
                    self._complete_microphone_fetch_gui_update,
                    mic_names,
                    selected_mic_name,
                    error_msg,
                    temp_mics_info,
                )

        threading.Thread(target=_fetch, daemon=True).start()

    def _complete_microphone_fetch_gui_update(self, names, sel_name, err, infos):
        self.available_mics_info = infos
        self.mic_combobox.configure(values=names if names else ["No mics"])
        self.mic_combobox.set(sel_name)
        if err:
            self._show_error_message("Mic Error", err)
            self._update_status_label("Error loading mics.", "red")
        elif names and names[0] != "No mics":
            # Only update status if it's not currently showing an error from another part
            if (
                "error" not in self.status_label.cget("text").lower()
                and "failed" not in self.status_label.cget("text").lower()
            ):
                self._update_status_label(f"Mic ready: {sel_name}", "green")
        else:
            self._update_status_label("No mics found.", "yellow")

    def _select_microphone(self, mic_name):
        for mic in self.available_mics_info:
            if mic["name"] == mic_name:
                self.settings["selected_mic_index"] = mic["index"]
                self._save_config()
                self._update_status_label(f"Mic selected: {mic_name}", "green")
                return
        self._show_error_message("Mic Error", f"Mic '{mic_name}' not found.")
        self.settings["selected_mic_index"] = -1
        self._save_config()

    def _show_image_frame(self, base64_data, caption):
        try:
            img_data = base64.b64decode(base64_data)
            img = Image.open(io.BytesIO(img_data))
            max_w, max_h = self.main_frame.winfo_width() - 40 or 400, 300
            orig_w, orig_h = img.size
            ratio = orig_w / orig_h
            disp_w = min(orig_w, max_w)
            disp_h = int(disp_w / ratio)
            if disp_h > max_h:
                disp_h = max_h
                disp_w = int(disp_h * ratio)
            disp_w, disp_h = max(1, disp_w), max(1, disp_h)
            resized_img = img.resize((disp_w, disp_h), Image.LANCZOS)
            ctk_img = ctk.CTkImage(
                light_image=resized_img, dark_image=resized_img, size=(disp_w, disp_h)
            )
            self.image_label.configure(image=ctk_img, text="")
            self.current_image = ctk_img
            self.image_caption_label.configure(text=caption, wraplength=disp_w - 20)
            self.image_frame.grid(row=5, column=0, padx=10, pady=(0, 10), sticky="ew")
            self.main_frame.grid_rowconfigure(5, weight=0)
            self.main_frame.grid_rowconfigure(4, weight=1)
        except Exception as e:
            self._show_error_message("Image Error", f"Failed to display image: {e}")
            self._hide_image_frame()

    def _hide_image_frame(self):
        self.image_frame.grid_remove()
        self.current_image = None
        self.main_frame.grid_rowconfigure(5, weight=0)
        self.main_frame.grid_rowconfigure(4, weight=2)

    def _show_code_window(self, code_text, language="text"):
        if self.code_window and self.code_window.winfo_exists():
            self.code_window.destroy()
        self.code_window = ctk.CTkToplevel(self)
        self.code_window.title(f"Generated Code ({language})")
        self.code_window.geometry("700x500")
        self.code_window.transient(self)
        self.code_window.protocol("WM_DELETE_WINDOW", self._destroy_code_window)
        self.code_window.grid_columnconfigure(0, weight=1)
        self.code_window.grid_rowconfigure(0, weight=1)
        display_code = code_text.strip()
        tb = ctk.CTkTextbox(
            self.code_window,
            wrap="none",
            font=("Fira Code", self.settings["font_size"] - 1),
        )
        tb.grid(row=0, column=0, padx=10, pady=(10, 5), sticky="nsew")
        if PYGMENTS_AVAILABLE:
            try:
                lexer = get_lexer_by_name(language)
            except:
                lexer = (
                    guess_lexer(display_code)
                    if PYGMENTS_AVAILABLE
                    else get_lexer_by_name("text")
                )  # Fallback
            dark = ctk.get_appearance_mode() == "Dark"
            for tt, val in lexer.get_tokens(display_code):
                tag, parent = str(tt), tt.parent
                colors = PYGMENTS_STYLE_MAP.get(tt)
                while parent and not colors:
                    colors = PYGMENTS_STYLE_MAP.get(parent)
                    parent = parent.parent
                fg = (
                    colors[0]
                    if dark
                    else colors[1] if colors else ("#f8f8f2" if dark else "#282a36")
                )
                tb.tag_config(tag, foreground=fg)
                tb.insert("end", val, tag)
        else:
            tb.insert("end", display_code)
        tb.configure(state="disabled")
        ctk.CTkButton(
            self.code_window,
            text="Copy Code",
            command=lambda: self._copy_code_to_clipboard(display_code),
        ).grid(row=1, column=0, padx=10, pady=(5, 10), sticky="ew")

    def _copy_code_to_clipboard(self, code_text):
        try:
            self.clipboard_clear()
            self.clipboard_append(code_text)
            messagebox.showinfo("Copied", "Code copied!", parent=self.code_window)
        except Exception as e:
            self._show_error_message("Copy Error", f"Failed to copy: {e}")

    def _destroy_code_window(self):
        if self.code_window:
            self.code_window.destroy()
            self.code_window = None

    def _on_closing(self):
        print("Closing application...")
        self.is_recording = False  # Ensure recording stops
        self.stop_tts_event.set()  # Signal TTS thread to stop and clear queue

        # Unhook keyboard listener
        try:
            print("Unhooking all keyboard events...")
            keyboard.unhook_all()
            print("Keyboard events unhooked.")
        except Exception as e:
            print(f"Error unhooking keyboard: {e}")
            pass  # Continue closing even if unhooking fails

        # Wait for threads to finish
        threads_to_join = []
        if self.recording_thread and self.recording_thread.is_alive():
            threads_to_join.append(("Recording", self.recording_thread))
        if self.transcription_thread and self.transcription_thread.is_alive():
            threads_to_join.append(("Transcription", self.transcription_thread))
        if self.ollama_response_thread and self.ollama_response_thread.is_alive():
            threads_to_join.append(("Ollama Response", self.ollama_response_thread))
        if self.tts_thread and self.tts_thread.is_alive():
            threads_to_join.append(("TTS Playback", self.tts_thread))

        for name, t in threads_to_join:
            print(f"Attempting to join {name} thread...")
            try:
                t.join(timeout=1.0)  # Increased timeout for join
                if t.is_alive():
                    print(f"Warning: {name} thread did not join in time.")
                else:
                    print(f"{name} thread joined successfully.")
            except Exception as e:
                print(f"Error joining {name} thread: {e}")

        # Stop TTS engine if busy
        if self.tts_engine:
            print("Checking TTS engine status...")
            if self.tts_engine.isBusy():
                print("TTS engine is busy, attempting to stop...")
                try:
                    self.tts_engine.stop()
                    print("TTS engine stopped.")
                except Exception as e:
                    print(f"Error stopping TTS engine: {e}")
            # Consider tts_engine.endLoop() if issues persist, but usually stop() is enough.
            # self.tts_engine = None # Optionally release the engine instance

        self._save_config()
        print("Configuration saved. Destroying main window.")
        self.destroy()
        print("Application closed.")


if __name__ == "__main__":
    app = OllamaSpeechChatApp()
    app.mainloop()
