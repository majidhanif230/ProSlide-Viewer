import os
import random
import tempfile
import time
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

try:
    from PIL import Image, ImageTk
except ImportError:
    Image = None
    ImageTk = None

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None


APP_TITLE = "PDF/PPTX Fullscreen Viewer"


class DeckLoader:
    def __init__(self):
        self.temp_dir = None

    def cleanup(self):
        if self.temp_dir and os.path.isdir(self.temp_dir):
            for name in os.listdir(self.temp_dir):
                path = os.path.join(self.temp_dir, name)
                try:
                    os.remove(path)
                except OSError:
                    pass
            try:
                os.rmdir(self.temp_dir)
            except OSError:
                pass
            self.temp_dir = None

    def load(self, path):
        ext = os.path.splitext(path)[1].lower()
        if ext == ".pdf":
            return self._load_pdf(path)
        if ext == ".pptx":
            return self._load_pptx(path)
        raise ValueError("Only .pdf and .pptx files are supported")

    def _load_pdf(self, path):
        if fitz is None or Image is None:
            raise RuntimeError(
                "PDF support requires Pillow and PyMuPDF. Run: pip install -r requirements.txt"
            )
        pages = []
        doc = fitz.open(path)
        try:
            for page in doc:
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
                mode = "RGB"
                image = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
                pages.append(image)
        finally:
            doc.close()

        if not pages:
            raise ValueError("No pages found in PDF")
        return pages

    def _load_pptx(self, path):
        if Image is None:
            raise RuntimeError(
                "PPTX support requires Pillow. Run: pip install -r requirements.txt"
            )
        try:
            import win32com.client  # type: ignore
        except ImportError as exc:
            raise RuntimeError(
                "PPTX support needs pywin32 and Microsoft PowerPoint installed"
            ) from exc

        self.cleanup()
        self.temp_dir = tempfile.mkdtemp(prefix="pptx_export_")

        app = win32com.client.Dispatch("PowerPoint.Application")
        app.Visible = 1
        presentation = None
        try:
            presentation = app.Presentations.Open(path, WithWindow=False)
            # 18 is ppSaveAsPNG
            presentation.SaveAs(self.temp_dir, 18)
        finally:
            if presentation is not None:
                presentation.Close()
            app.Quit()

        image_files = sorted(
            [
                os.path.join(self.temp_dir, f)
                for f in os.listdir(self.temp_dir)
                if f.lower().endswith(".png")
            ]
        )

        pages = []
        for image_path in image_files:
            pages.append(Image.open(image_path).convert("RGB"))

        if not pages:
            raise ValueError("Could not export slides from PPTX")
        return pages


class ViewerWindow:
    def __init__(
        self,
        root,
        images,
        timings,
        transition,
        start_index=0,
        loop_mode=True,
        show_countdown=True,
        pure_fullscreen=True,
        source_name="Presentation",
    ):
        if Image is None or ImageTk is None:
            raise RuntimeError(
                "Viewer requires Pillow. Run: pip install -r requirements.txt"
            )
        self.root = root
        self.images = images
        self.timings = timings
        self.transition = transition
        self.loop_mode = loop_mode
        self.show_countdown = show_countdown
        self.pure_fullscreen = pure_fullscreen
        self.source_name = source_name

        self.index = max(0, min(start_index, len(self.images) - 1))
        self.running = True
        self.after_id = None
        self.countdown_id = None
        self.tk_image = None
        self.slide_deadline = None
        self.blackout = False
        self.animation_token = 0
        self.overlay_visible = True
        self.overlay_hide_id = None
        self.direction = 1
        self.toast_label = None
        self.toast_hide_id = None
        self.export_dir = os.path.join(os.getcwd(), "exports")

        self.win = tk.Toplevel(root)
        self.win.title("Presentation Mode")
        self.win.configure(bg="#000000")
        self.win.attributes("-fullscreen", True)

        self.canvas = tk.Canvas(self.win, bg="black", highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)

        self.top_bar = tk.Frame(self.win, bg="#0f172a", height=42)
        self.top_bar.pack_propagate(False)

        self.title_label = tk.Label(
            self.top_bar,
            text=f"Now Presenting: {self.source_name}",
            bg="#0f172a",
            fg="#e2e8f0",
            font=("Segoe UI Semibold", 11),
            padx=10,
        )
        self.title_label.pack(side="left")

        self.clock_label = tk.Label(
            self.top_bar,
            text="",
            bg="#0f172a",
            fg="#cbd5e1",
            font=("Segoe UI", 10),
            padx=10,
        )
        self.clock_label.pack(side="right")

        self.exit_button = tk.Button(
            self.top_bar,
            text="Exit (Esc)",
            command=self.close,
            bg="#dc2626",
            fg="white",
            activebackground="#b91c1c",
            activeforeground="white",
            relief="flat",
            padx=10,
            pady=4,
            cursor="hand2",
        )
        self.exit_button.pack(side="right", padx=8, pady=6)

        self.controls = tk.Frame(self.win, bg="#111827")

        self.info_label = tk.Label(
            self.controls,
            text="",
            bg="#111827",
            fg="#f8fafc",
            padx=10,
            pady=8,
            font=("Segoe UI", 10),
        )
        self.info_label.pack(side="left")

        self.countdown_label = tk.Label(
            self.controls,
            text="",
            bg="#111827",
            fg="#93c5fd",
            padx=4,
            pady=8,
            font=("Segoe UI", 10),
        )
        self.countdown_label.pack(side="left")

        self.guide_label = tk.Label(
            self.controls,
            text=(
                "Guide: Left/Right Next-Prev | Space Pause | B Blackout | F Fullscreen | "
                "G Guides | J Jump | S Save | +/- Timing | R Direction | Esc Exit"
            ),
            bg="#111827",
            fg="#a5b4fc",
            padx=4,
            pady=8,
            font=("Segoe UI", 9),
        )
        self.guide_label.pack(side="left", padx=(8, 0))

        self.progress = ttk.Progressbar(
            self.controls,
            orient="horizontal",
            mode="determinate",
            maximum=max(1, len(self.images)),
            length=180,
        )
        self.progress.pack(side="right", padx=8, pady=9)

        tk.Button(self.controls, text="Prev", command=self.prev_slide, width=8).pack(
            side="right", padx=4, pady=6
        )
        tk.Button(self.controls, text="Next", command=self.next_slide, width=8).pack(
            side="right", padx=4, pady=6
        )
        tk.Button(
            self.controls,
            text="Pause/Resume",
            command=self.toggle_running,
            width=12,
        ).pack(side="right", padx=4, pady=6)
        tk.Button(
            self.controls,
            text="Blackout (B)",
            command=self.toggle_blackout,
            width=12,
        ).pack(side="right", padx=4, pady=6)
        tk.Button(self.controls, text="Stop", command=self.close, width=8).pack(
            side="right", padx=4, pady=6
        )

        self.top_bar.place(relx=0.0, rely=0.0, relwidth=1.0)
        self.controls.place(relx=0.0, rely=1.0, relwidth=1.0, anchor="sw")

        self.win.bind("<Escape>", lambda _e: self.close())
        self.win.bind("<Right>", lambda _e: self.next_slide())
        self.win.bind("<Left>", lambda _e: self.prev_slide())
        self.win.bind("<space>", lambda _e: self.toggle_running())
        self.win.bind("<b>", lambda _e: self.toggle_blackout())
        self.win.bind("<B>", lambda _e: self.toggle_blackout())
        self.win.bind("<f>", lambda _e: self.toggle_fullscreen())
        self.win.bind("<F>", lambda _e: self.toggle_fullscreen())
        self.win.bind("<g>", lambda _e: self.toggle_overlay())
        self.win.bind("<G>", lambda _e: self.toggle_overlay())
        self.win.bind("<j>", lambda _e: self.jump_to_slide_dialog())
        self.win.bind("<J>", lambda _e: self.jump_to_slide_dialog())
        self.win.bind("<s>", lambda _e: self.save_current_slide())
        self.win.bind("<S>", lambda _e: self.save_current_slide())
        self.win.bind("<r>", lambda _e: self.toggle_direction())
        self.win.bind("<R>", lambda _e: self.toggle_direction())
        self.win.bind("<plus>", lambda _e: self.adjust_timing(1.0))
        self.win.bind("<KP_Add>", lambda _e: self.adjust_timing(1.0))
        self.win.bind("<minus>", lambda _e: self.adjust_timing(-1.0))
        self.win.bind("<KP_Subtract>", lambda _e: self.adjust_timing(-1.0))
        self.win.bind("<Home>", lambda _e: self.go_to_slide(0))
        self.win.bind("<End>", lambda _e: self.go_to_slide(len(self.images) - 1))
        self.win.bind("<Configure>", lambda _e: self.show_slide(with_transition=False))

        self.show_slide(with_transition=False)
        self.schedule_next()
        self.update_clock()
        if self.pure_fullscreen:
            self._hide_overlay()
        else:
            self._show_overlay(temp=False)

    def _current_timing_ms(self):
        seconds = self.timings[self.index] if self.index < len(self.timings) else self.timings[-1]
        return max(1, int(seconds * 1000))

    def _canvas_size(self):
        width = max(1, self.canvas.winfo_width())
        height = max(1, self.canvas.winfo_height())
        return width, height

    def _fit_to_canvas(self, image):
        width, height = self._canvas_size()

        img_w, img_h = image.size
        scale = min(width / img_w, height / img_h)
        new_size = (max(1, int(img_w * scale)), max(1, int(img_h * scale)))
        return image.resize(new_size, Image.Resampling.LANCZOS)

    def _draw_image_center(self, image):
        self.canvas.delete("all")
        self.tk_image = ImageTk.PhotoImage(image)
        cw, ch = self._canvas_size()
        self.canvas.create_image(cw // 2, ch // 2, image=self.tk_image)

    def _fade_in(self, image, token):
        fitted = self._fit_to_canvas(image)
        black = Image.new("RGB", fitted.size, "black")
        steps = 10

        def step(i):
            if token != self.animation_token:
                return
            blended = Image.blend(black, fitted, i / steps)
            self.tk_image = ImageTk.PhotoImage(blended)
            self.canvas.delete("all")
            cw, ch = self._canvas_size()
            self.canvas.create_image(cw // 2, ch // 2, image=self.tk_image)
            if i < steps:
                self.win.after(25, lambda: step(i + 1))

        step(0)

    def _slide_left(self, image, token):
        fitted = self._fit_to_canvas(image)
        cw, ch = self._canvas_size()
        iw, ih = fitted.size
        steps = 12
        start_x = cw
        end_x = (cw - iw) // 2

        def step(i):
            if token != self.animation_token:
                return
            x = int(start_x + (end_x - start_x) * (i / steps))
            frame = Image.new("RGB", (cw, ch), "black")
            frame.paste(fitted, (x, (ch - ih) // 2))
            self._draw_image_center(frame)
            if i < steps:
                self.win.after(20, lambda: step(i + 1))

        step(0)

    def _zoom_in(self, image, token):
        fitted = self._fit_to_canvas(image)
        cw, ch = self._canvas_size()
        fw, fh = fitted.size
        steps = 12

        def step(i):
            if token != self.animation_token:
                return
            zoom = 0.85 + (0.15 * (i / steps))
            zw = max(1, int(fw * zoom))
            zh = max(1, int(fh * zoom))
            scaled = fitted.resize((zw, zh), Image.Resampling.LANCZOS)
            frame = Image.new("RGB", (cw, ch), "black")
            frame.paste(scaled, ((cw - zw) // 2, (ch - zh) // 2))
            self._draw_image_center(frame)
            if i < steps:
                self.win.after(20, lambda: step(i + 1))

        step(0)

    def _draw_black(self):
        self.canvas.delete("all")
        self.canvas.configure(bg="black")

    def show_slide(self, with_transition=True):
        if not self.images:
            return

        self.animation_token += 1
        token = self.animation_token

        if self.blackout:
            self._draw_black()
            self._update_status()
            return

        image = self.images[self.index]
        if with_transition and self.transition == "Fade":
            self._fade_in(image, token)
        elif with_transition and self.transition == "Slide Left":
            self._slide_left(image, token)
        elif with_transition and self.transition == "Zoom In":
            self._zoom_in(image, token)
        else:
            fitted = self._fit_to_canvas(image)
            self._draw_image_center(fitted)

        self._update_status()

    def _update_status(self):
        seconds = self.timings[self.index] if self.index < len(self.timings) else self.timings[-1]
        status = "Running" if self.running else "Paused"
        if self.blackout:
            status = "Blackout"
        direction_label = "Forward" if self.direction > 0 else "Reverse"
        self.info_label.config(
            text=(
                f"Slide {self.index + 1}/{len(self.images)} | "
                f"{seconds:.1f}s | {status} | {direction_label} | Transition: {self.transition}"
            )
        )
        self.progress.configure(value=self.index + 1)

    def schedule_next(self):
        if self.after_id is not None:
            self.win.after_cancel(self.after_id)
            self.after_id = None
        if self.countdown_id is not None:
            self.win.after_cancel(self.countdown_id)
            self.countdown_id = None

        if self.running:
            self.slide_deadline = time.time() + (self._current_timing_ms() / 1000.0)
            self.after_id = self.win.after(self._current_timing_ms(), self.auto_next)
            self.update_countdown()
        else:
            self.slide_deadline = None
            self.update_countdown(force_text="")

    def update_countdown(self, force_text=None):
        if force_text is not None:
            self.countdown_label.config(text=force_text)
            return

        if not self.show_countdown:
            self.countdown_label.config(text="")
            return

        if self.slide_deadline is None:
            self.countdown_label.config(text="")
            return

        remaining = max(0.0, self.slide_deadline - time.time())
        self.countdown_label.config(text=f"Next in: {remaining:.1f}s")
        if self.running:
            self.countdown_id = self.win.after(100, self.update_countdown)

    def update_clock(self):
        now = time.strftime("%H:%M:%S")
        self.clock_label.config(text=f"{now}  |  J: Jump  S: Save  +/-: Time  R: Direction")
        self.win.after(1000, self.update_clock)

    def _show_overlay(self, temp=False):
        if self.pure_fullscreen and not self.overlay_visible:
            temp = False

        if self.overlay_hide_id is not None:
            self.win.after_cancel(self.overlay_hide_id)
            self.overlay_hide_id = None

        self.overlay_visible = True
        self.top_bar.place(relx=0.0, rely=0.0, relwidth=1.0)
        self.controls.place(relx=0.0, rely=1.0, relwidth=1.0, anchor="sw")

        if temp:
            self.overlay_hide_id = self.win.after(2500, self._hide_overlay)

    def _hide_overlay(self):
        self.overlay_visible = False
        self.top_bar.place_forget()
        self.controls.place_forget()

    def toggle_overlay(self):
        if self.overlay_visible:
            self._hide_overlay()
        else:
            self._show_overlay(temp=self.pure_fullscreen)
            if self.pure_fullscreen:
                self.show_toast("Guides shown. Press G to hide for pure fullscreen.")

    def auto_next(self):
        if not self.running:
            return

        if not self._advance(self.direction):
            self.running = False
            self._update_status()
            self.schedule_next()
            self.show_toast("Reached the end in current direction")
            return

        self.show_slide(with_transition=True)
        self.schedule_next()

    def next_slide(self):
        self._advance(1)
        self.show_slide(with_transition=True)
        self.schedule_next()

    def prev_slide(self):
        self._advance(-1)
        self.show_slide(with_transition=True)
        self.schedule_next()

    def _advance(self, step):
        target = self.index + step
        if target < 0:
            if self.loop_mode:
                self.index = len(self.images) - 1
                return True
            self.index = 0
            return False
        if target >= len(self.images):
            if self.loop_mode:
                self.index = 0
                return True
            self.index = len(self.images) - 1
            return False
        self.index = target
        return True

    def go_to_slide(self, index):
        self.index = max(0, min(index, len(self.images) - 1))
        self.show_slide(with_transition=False)
        self.schedule_next()

    def toggle_running(self):
        self.running = not self.running
        self.show_slide(with_transition=False)
        self.schedule_next()

    def toggle_blackout(self):
        self.blackout = not self.blackout
        self.show_slide(with_transition=False)

    def toggle_fullscreen(self):
        current = bool(self.win.attributes("-fullscreen"))
        self.win.attributes("-fullscreen", not current)

    def adjust_timing(self, delta):
        current = self.timings[self.index]
        updated = max(0.5, round(current + delta, 2))
        self.timings[self.index] = updated
        self.show_slide(with_transition=False)
        self.schedule_next()
        self.show_toast(f"Slide {self.index + 1} timing set to {updated:.2f}s")

    def jump_to_slide_dialog(self):
        target = simpledialog.askinteger(
            "Jump to Slide",
            f"Enter slide number (1-{len(self.images)}):",
            parent=self.win,
            minvalue=1,
            maxvalue=len(self.images),
        )
        if target is None:
            return
        self.go_to_slide(target - 1)
        self.show_toast(f"Jumped to slide {target}")

    def toggle_direction(self):
        self.direction *= -1
        self._update_status()
        direction_label = "forward" if self.direction > 0 else "reverse"
        self.show_toast(f"Autoplay direction: {direction_label}")

    def save_current_slide(self):
        os.makedirs(self.export_dir, exist_ok=True)
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        path = os.path.join(self.export_dir, f"slide_{self.index + 1}_{timestamp}.png")
        self.images[self.index].save(path, "PNG")
        self.show_toast(f"Saved snapshot: {os.path.basename(path)}")

    def show_toast(self, text):
        if self.toast_label is None:
            self.toast_label = tk.Label(
                self.win,
                text="",
                bg="#0b1020",
                fg="#e2e8f0",
                font=("Segoe UI", 10),
                padx=12,
                pady=6,
            )

        self.toast_label.config(text=text)
        self.toast_label.place(relx=0.5, rely=0.94, anchor="s")

        if self.toast_hide_id is not None:
            self.win.after_cancel(self.toast_hide_id)
        self.toast_hide_id = self.win.after(1700, self.toast_label.place_forget)

    def close(self):
        if self.after_id is not None:
            self.win.after_cancel(self.after_id)
            self.after_id = None
        if self.countdown_id is not None:
            self.win.after_cancel(self.countdown_id)
            self.countdown_id = None
        if self.overlay_hide_id is not None:
            self.win.after_cancel(self.overlay_hide_id)
            self.overlay_hide_id = None
        if self.toast_hide_id is not None:
            self.win.after_cancel(self.toast_hide_id)
            self.toast_hide_id = None
        self.win.destroy()


class App:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("900x540")
        self.root.minsize(840, 500)

        self.loader = DeckLoader()
        self.selected_file = tk.StringVar(value="")
        self.default_time = tk.StringVar(value="10")
        self.custom_times = tk.StringVar(value="")
        self.random_min = tk.StringVar(value="8")
        self.random_max = tk.StringVar(value="20")
        self.start_slide = tk.StringVar(value="1")
        self.timing_mode = tk.StringVar(value="default")
        self.transition = tk.StringVar(value="Fade")
        self.loop_mode = tk.BooleanVar(value=True)
        self.shuffle_slides = tk.BooleanVar(value=False)
        self.show_countdown = tk.BooleanVar(value=True)
        self.pure_fullscreen = tk.BooleanVar(value=True)

        self._build_ui()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self):
        style = ttk.Style(self.root)
        style.configure("Title.TLabel", font=("Segoe UI Semibold", 16))
        style.configure("Sub.TLabel", foreground="#475569")

        frame = ttk.Frame(self.root, padding=18)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Professional Presentation Viewer", style="Title.TLabel").pack(
            anchor="w"
        )
        ttk.Label(
            frame,
            text="Load PDF/PPTX, control autoplay timing, and present with fullscreen controls.",
            style="Sub.TLabel",
        ).pack(anchor="w", pady=(4, 14))

        file_group = ttk.LabelFrame(frame, text="File")
        file_group.pack(fill="x", pady=(0, 10))
        file_row = ttk.Frame(file_group, padding=10)
        file_row.pack(fill="x")
        ttk.Button(file_row, text="Select PDF/PPTX", command=self.pick_file).pack(side="left")
        ttk.Label(file_row, textvariable=self.selected_file, wraplength=640).pack(
            side="left", padx=12
        )

        timing_group = ttk.LabelFrame(frame, text="Timing")
        timing_group.pack(fill="x", pady=(0, 10))
        timing_inner = ttk.Frame(timing_group, padding=10)
        timing_inner.pack(fill="x")

        ttk.Radiobutton(
            timing_inner,
            text="Fixed (all slides)",
            variable=self.timing_mode,
            value="default",
        ).grid(row=0, column=0, sticky="w")
        ttk.Entry(timing_inner, textvariable=self.default_time, width=10).grid(
            row=0, column=1, padx=(8, 18), sticky="w"
        )
        ttk.Label(timing_inner, text="seconds").grid(row=0, column=2, sticky="w")

        ttk.Radiobutton(
            timing_inner,
            text="Custom list",
            variable=self.timing_mode,
            value="custom",
        ).grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(timing_inner, textvariable=self.custom_times, width=48).grid(
            row=1, column=1, columnspan=3, padx=(8, 0), pady=(8, 0), sticky="w"
        )

        ttk.Radiobutton(
            timing_inner,
            text="Random duration per slide",
            variable=self.timing_mode,
            value="random",
        ).grid(row=2, column=0, sticky="w", pady=(8, 0))
        ttk.Label(timing_inner, text="Min").grid(row=2, column=1, sticky="e", pady=(8, 0))
        ttk.Entry(timing_inner, textvariable=self.random_min, width=8).grid(
            row=2, column=2, sticky="w", pady=(8, 0)
        )
        ttk.Label(timing_inner, text="Max").grid(row=2, column=3, sticky="e", pady=(8, 0))
        ttk.Entry(timing_inner, textvariable=self.random_max, width=8).grid(
            row=2, column=4, sticky="w", pady=(8, 0)
        )
        ttk.Label(timing_inner, text="seconds").grid(row=2, column=5, sticky="w", pady=(8, 0))

        ttk.Label(
            timing_inner,
            text=(
                "Custom example: 10,20,15. If list is shorter than slides, "
                "the last value is reused."
            ),
            style="Sub.TLabel",
        ).grid(row=3, column=0, columnspan=6, sticky="w", pady=(8, 0))

        options_group = ttk.LabelFrame(frame, text="Playback Options")
        options_group.pack(fill="x", pady=(0, 10))
        options_inner = ttk.Frame(options_group, padding=10)
        options_inner.pack(fill="x")

        ttk.Label(options_inner, text="Transition:").grid(row=0, column=0, sticky="w")
        ttk.Combobox(
            options_inner,
            textvariable=self.transition,
            values=["None", "Fade", "Slide Left", "Zoom In"],
            state="readonly",
            width=14,
        ).grid(row=0, column=1, sticky="w", padx=(8, 20))

        ttk.Label(options_inner, text="Start from slide:").grid(row=0, column=2, sticky="w")
        ttk.Entry(options_inner, textvariable=self.start_slide, width=8).grid(
            row=0, column=3, sticky="w", padx=(8, 12)
        )

        ttk.Checkbutton(options_inner, text="Loop playback", variable=self.loop_mode).grid(
            row=1, column=0, sticky="w", pady=(8, 0)
        )
        ttk.Checkbutton(
            options_inner,
            text="Shuffle slides/pages",
            variable=self.shuffle_slides,
        ).grid(row=1, column=1, sticky="w", pady=(8, 0))
        ttk.Checkbutton(
            options_inner,
            text="Show countdown timer",
            variable=self.show_countdown,
        ).grid(row=1, column=2, sticky="w", pady=(8, 0))
        ttk.Checkbutton(
            options_inner,
            text="Pure fullscreen (guides hidden by default)",
            variable=self.pure_fullscreen,
        ).grid(row=1, column=3, sticky="w", pady=(8, 0))

        footer = ttk.Frame(frame)
        footer.pack(fill="x", pady=(4, 0))

        ttk.Label(
            footer,
            text=(
                "Fullscreen keys: Left/Right = prev/next, Space = pause/resume, "
                "B = blackout, F = fullscreen, G = guides, J = jump, S = snapshot, "
                "+/- = timing, R = direction, Esc = exit"
            ),
            style="Sub.TLabel",
            wraplength=820,
        ).pack(side="left")

        ttk.Button(
            frame,
            text="Start Fullscreen Presentation",
            command=self.start_show,
        ).pack(anchor="w", pady=(12, 0))

    def pick_file(self):
        file_path = filedialog.askopenfilename(
            title="Select PDF or PPTX",
            filetypes=[("Supported", "*.pdf *.pptx"), ("PDF", "*.pdf"), ("PPTX", "*.pptx")],
        )
        if file_path:
            self.selected_file.set(file_path)

    def _parse_timings(self, slide_count):
        mode = self.timing_mode.get()

        if mode == "custom":
            custom = self.custom_times.get().strip()
            raw_parts = [p.strip() for p in custom.split(",") if p.strip()]
            if not raw_parts:
                raise ValueError("Custom timing list is empty")
            values = [float(v) for v in raw_parts]
            if any(v <= 0 for v in values):
                raise ValueError("Custom timings must be positive")
            if len(values) < slide_count:
                values.extend([values[-1]] * (slide_count - len(values)))
            return values[:slide_count]

        if mode == "random":
            min_val = float(self.random_min.get().strip())
            max_val = float(self.random_max.get().strip())
            if min_val <= 0 or max_val <= 0:
                raise ValueError("Random min and max must be positive")
            if min_val > max_val:
                raise ValueError("Random min cannot be greater than random max")
            return [round(random.uniform(min_val, max_val), 2) for _ in range(slide_count)]

        default_value = float(self.default_time.get().strip())
        if default_value <= 0:
            raise ValueError("Default timing must be positive")
        return [default_value] * slide_count

    def _parse_start_index(self, slide_count):
        value = int(self.start_slide.get().strip())
        if value < 1 or value > slide_count:
            raise ValueError(f"Start slide must be between 1 and {slide_count}")
        return value - 1

    def start_show(self):
        path = self.selected_file.get().strip()
        if not path:
            messagebox.showerror("No File", "Please select a PDF or PPTX file first.")
            return

        try:
            images = self.loader.load(path)
            timings = self._parse_timings(len(images))
            start_index = self._parse_start_index(len(images))
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return

        if self.shuffle_slides.get():
            pairs = list(zip(images, timings))
            random.shuffle(pairs)
            images = [p[0] for p in pairs]
            timings = [p[1] for p in pairs]
            start_index = 0

        ViewerWindow(
            self.root,
            images,
            timings,
            self.transition.get(),
            start_index=start_index,
            loop_mode=self.loop_mode.get(),
            show_countdown=self.show_countdown.get(),
            pure_fullscreen=self.pure_fullscreen.get(),
            source_name=os.path.basename(path),
        )

    def _on_close(self):
        self.loader.cleanup()
        self.root.destroy()


def main():
    root = tk.Tk()
    style = ttk.Style(root)
    if "vista" in style.theme_names():
        style.theme_use("vista")
    app = App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
