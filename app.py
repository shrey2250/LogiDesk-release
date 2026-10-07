import customtkinter as ctk
from tkinter import filedialog, messagebox, PhotoImage
import pandas as pd
import os
import math
import re
import uuid
from datetime import date, datetime
import xml.etree.ElementTree as ET
from xml.dom import minidom

from bill_generator import (
    BillGeneratorService,
    BILL_MARKER_OPTIONS,
    DEFAULT_BILL_SAC_CODE,
    parse_date,
)
from updater import check_for_updates_async, APP_VERSION


# ============================================================
# APPEARANCE
# ============================================================

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


class CalendarDateEntry(ctk.CTkFrame):
    def __init__(self, master, callback=None, width=150, height=30, **kwargs):
        super().__init__(master, fg_color="transparent")
        self.callback = callback
        self.selected_date = None
        self.current_month = date.today().replace(day=1)
        self.popup = None

        self.button = ctk.CTkButton(
            self,
            text="Select date",
            width=width,
            height=height,
            corner_radius=8,
            fg_color=("#FFFFFF", "#1C1C1E"),
            text_color=("#3C3C43", "#EBEBF5"),
            border_width=1,
            border_color=("#D1D1D6", "#38383A"),
            hover_color=("#F2F2F7", "#2C2C2E"),
            anchor="w",
            command=self._open_calendar,
            **kwargs,
        )
        self.button.grid(sticky="ew")

    def set(self, value):
        parsed = parse_date(value)
        if parsed is None:
            self.selected_date = None
            self.button.configure(text="Select date")
            if self.callback:
                self.callback()
            return

        if isinstance(parsed, datetime):
            self.selected_date = parsed.date()
        else:
            self.selected_date = parsed.date()
        self.button.configure(text=self.selected_date.strftime("%d.%m.%y"))
        if self.callback:
            self.callback()

    def get(self):
        if self.selected_date is None:
            return ""
        return self.selected_date.strftime("%d.%m.%y")

    def get_date(self):
        return self.selected_date

    def _open_calendar(self):
        if self.popup is not None and self.popup.winfo_exists():
            self.popup.focus_force()
            return

        popup = ctk.CTkToplevel(self)
        popup.title("Select Date")
        popup.geometry("280x300")
        popup.transient(self.winfo_toplevel())
        popup.grab_set()
        popup.configure(fg_color=("#F2F2F7", "#1C1C1E"))
        self.popup = popup

        popup.grid_columnconfigure(0, weight=1)
        popup.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(popup, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=12, pady=(12, 8))
        header.grid_columnconfigure(0, weight=1)
        header.grid_columnconfigure(1, weight=1)
        header.grid_columnconfigure(2, weight=1)

        prev = ctk.CTkButton(header, text="‹", width=32, height=30, command=lambda: self._shift_month(-1),
                             corner_radius=8,
                             fg_color=("#FFFFFF", "#2C2C2E"), text_color=("#007AFF", "#0A84FF"),
                             border_width=1, border_color=("#D1D1D6", "#38383A"),
                             hover_color=("#F2F2F7", "#1C1C1E"),
                             font=ctk.CTkFont(family="Segoe UI", size=16))
        prev.grid(row=0, column=0, sticky="w")

        month_label = ctk.CTkLabel(header, text="", font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"), text_color=("#000000", "#FFFFFF"))
        month_label.grid(row=0, column=1)

        next_btn = ctk.CTkButton(header, text="›", width=32, height=30, command=lambda: self._shift_month(1),
                                  corner_radius=8,
                                  fg_color=("#FFFFFF", "#2C2C2E"), text_color=("#007AFF", "#0A84FF"),
                                  border_width=1, border_color=("#D1D1D6", "#38383A"),
                                  hover_color=("#F2F2F7", "#1C1C1E"),
                                  font=ctk.CTkFont(family="Segoe UI", size=16))
        next_btn.grid(row=0, column=2, sticky="e")

        calendar = ctk.CTkFrame(popup, fg_color="transparent")
        calendar.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 12))
        calendar.grid_columnconfigure(tuple(range(7)), weight=1)

        self._calendar_month = self.selected_date.replace(day=1) if self.selected_date else date.today().replace(day=1)
        self._calendar_label = month_label
        self._calendar_grid = calendar
        self._render_calendar(month_label, calendar)

        popup.protocol("WM_DELETE_WINDOW", lambda: self._close_popup(popup))

    def _close_popup(self, popup):
        popup.destroy()
        self.popup = None

    def _shift_month(self, delta):
        year = self._calendar_month.year + (self._calendar_month.month - 1 + delta) // 12
        month = (self._calendar_month.month - 1 + delta) % 12 + 1
        self._calendar_month = date(year, month, 1)
        self._render_calendar(self._calendar_label, self._calendar_grid)

    def _render_calendar(self, month_label, calendar):
        for child in calendar.winfo_children():
            child.destroy()

        month_label.configure(text=self._calendar_month.strftime("%B %Y"))
        days = ["Su", "Mo", "Tu", "We", "Th", "Fr", "Sa"]
        for idx, label in enumerate(days):
            ctk.CTkLabel(calendar, text=label, font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"), text_color=("#3C3C43", "#EBEBF5")).grid(row=0, column=idx, padx=2, pady=(0, 6))

        first_weekday = self._calendar_month.weekday() + 1
        first_day = 1
        total_days = 31
        while True:
            try:
                date(self._calendar_month.year, self._calendar_month.month, total_days)
                break
            except ValueError:
                total_days -= 1

        row = 1
        col = first_weekday % 7
        for day in range(1, total_days + 1):
            cell = date(self._calendar_month.year, self._calendar_month.month, day)
            is_selected = self.selected_date == cell
            btn = ctk.CTkButton(
                calendar,
                text=str(day),
                width=28,
                height=28,
                corner_radius=8,
                fg_color=("#FFFFFF", "#2C2C2E") if not is_selected else ("#007AFF", "#0A84FF"),
                text_color=("#000000", "#FFFFFF") if not is_selected else ("#FFFFFF", "#FFFFFF"),
                border_width=1 if not is_selected else 0,
                border_color=("#E5E5EA", "#38383A"),
                hover_color=("#F2F2F7", "#3A3A3C") if not is_selected else ("#0062CC", "#3395FF"),
                font=ctk.CTkFont(family="Segoe UI", size=11),
                command=lambda chosen=cell: self._select_date(chosen),
            )
            btn.grid(row=row, column=col, padx=2, pady=2, sticky="nsew")
            col += 1
            if col > 6:
                col = 0
                row += 1

    def _select_date(self, selected):
        self.selected_date = selected
        self.button.configure(text=self.selected_date.strftime("%d.%m.%y"))
        if self.callback:
            self.callback()
        if self.popup is not None and self.popup.winfo_exists():
            self.popup.destroy()
            self.popup = None


# ============================================================
# APPLICATION
# ============================================================

class TallyConverterApp(ctk.CTk):

    def __init__(self):
        super().__init__()

        self.title("Shrey Logistics — Business Automation Suite")
        self.geometry("1180x820")
        self.minsize(1020, 680)

        # Excel
        self.excel_path = ""
        self.workbook = None
        self.data = None

        # Sheet/header
        self.current_sheet = ""
        self.header_row = None

        # Mapping
        self.mapping_path = ""
        self.mapping_data = None
        self.mapping_dict = {}

        # Columns
        self.date_column = None
        self.lr_column = None
        self.broker_column = None
        self.amount_column = None

        # States
        self.data_valid = False
        self.mapping_valid = False

        self.bill_service = BillGeneratorService()
        self.bill_excel_path = ""
        self.bill_mapping_path = ""
        self.bill_template_path = ""

        self.xml_frame_root = None
        self.bill_frame_root = None
        self.status = None
        self.bill_status = None
        self.active_tool = None
        self.nav_buttons = {}

        self.create_ui()
        self.build_shell()

        # Check for updates silently in background on startup
        self.after(1500, lambda: check_for_updates_async(self, silent=True))

    # ============================================================
    # DESIGN TOKENS  —  Apple Human Interface Guidelines palette
    # ============================================================

    # Backgrounds — layered depth system
    BG          = ("#F2F2F7", "#1C1C1E")   # system background (iOS/macOS primary)
    SURFACE     = ("#FFFFFF", "#2C2C2E")   # grouped secondary background
    SURFACE2    = ("#F2F2F7", "#3A3A3C")   # elevated tertiary surface / hover state
    SEP         = ("#E5E5EA", "#38383A")   # opaque separator

    # Typography — Apple text style hierarchy
    T_PRIMARY   = ("#000000", "#FFFFFF")   # label / primary text
    T_SECONDARY = ("#3C3C43", "#EBEBF5")   # secondary label
    T_MUTED     = ("#8E8E93", "#636366")   # tertiary label / placeholder

    # Interactive — Apple Blue system accent
    BLUE        = ("#007AFF", "#0A84FF")   # system blue — primary CTA
    BLUE_HOVER  = ("#0062CC", "#3395FF")   # blue darkened / lightened for hover

    # Semantic — Apple system colors
    GREEN       = ("#34C759", "#30D158")   # system green
    AMBER       = ("#FF9500", "#FF9F0A")   # system orange / warning
    RED         = ("#FF3B30", "#FF453A")   # system red / destructive
    PURPLE      = ("#AF52DE", "#BF5AF2")   # system purple — second tool accent

    # Per-tool accent (icon badges, active nav state, header dot)
    ACCENT = {"xml": BLUE, "bill": PURPLE}

    # Input / combobox — clean elevated inputs
    INPUT_BG    = ("#FFFFFF", "#1C1C1E")
    INPUT_BDR   = ("#D1D1D6", "#38383A")

    # ============================================================
    # UI ENTRY POINT
    # ============================================================

    def create_ui(self):
        self.configure(fg_color=self.BG)

        # ── Fonts — Apple Dynamic Type scale
        # SF Pro Display for large text, SF Pro Text for body/UI
        UI_FONT   = "SF Pro Display" if os.name == "posix" else "Segoe UI"
        UI_TEXT   = "SF Pro Text"    if os.name == "posix" else "Segoe UI"
        MONO_FONT = "SF Mono"        if os.name == "posix" else "Cascadia Code"

        # Title: Large Title — 28pt bold, tight tracking
        self.F_TITLE    = ctk.CTkFont(family=UI_FONT,  size=22, weight="bold")
        # Subtitle: Callout
        self.F_SUBTITLE = ctk.CTkFont(family=UI_TEXT,  size=13)
        # Section headers: Caption 2 — all-caps, semibold, wide tracking
        self.F_SECTION  = ctk.CTkFont(family=UI_TEXT,  size=11, weight="bold")
        # Step number: Footnote
        self.F_STEP_NUM = ctk.CTkFont(family=UI_TEXT,  size=12)
        # Label: Body
        self.F_LABEL    = ctk.CTkFont(family=UI_TEXT,  size=14)
        # Label semibold: Body Emphasized
        self.F_LABEL_SB = ctk.CTkFont(family=UI_TEXT,  size=14, weight="bold")
        # Body: Callout
        self.F_BODY     = ctk.CTkFont(family=UI_TEXT,  size=13)
        # Small: Footnote
        self.F_SMALL    = ctk.CTkFont(family=UI_TEXT,  size=12)
        # Tiny: Caption 2
        self.F_TINY     = ctk.CTkFont(family=UI_TEXT,  size=11)
        # Button: Callout medium — not all-caps
        self.F_BTN      = ctk.CTkFont(family=UI_TEXT,  size=13)
        # Primary button: Body semibold
        self.F_BTN_PRI  = ctk.CTkFont(family=UI_TEXT,  size=14, weight="bold")
        # Monospace: code/data preview
        self.F_MONO     = ctk.CTkFont(family=MONO_FONT, size=12)
        # Status bar: Caption 1
        self.F_STATUS   = ctk.CTkFont(family=UI_TEXT,  size=12)

    # ============================================================
    # APP SHELL  —  persistent sidebar + swappable content pane
    # ============================================================

    def build_shell(self):
        self.configure(fg_color=self.BG)
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        self._build_sidebar()

        self.content_area = ctk.CTkFrame(self, fg_color=self.BG, corner_radius=0)
        self.content_area.grid(row=0, column=1, sticky="nsew")
        self.content_area.grid_rowconfigure(0, weight=1)
        self.content_area.grid_columnconfigure(0, weight=1)

        # Land on the first tool by default — no gate screen to click through.
        self.select_tool("xml")

    def _build_sidebar(self):
        sidebar = ctk.CTkFrame(self, fg_color=self.SURFACE, corner_radius=0, width=236)
        sidebar.grid(row=0, column=0, sticky="nsw")
        sidebar.grid_propagate(False)
        sidebar.grid_rowconfigure(3, weight=1)

        # Hairline separator on the right edge of the sidebar.
        ctk.CTkFrame(sidebar, fg_color=self.SEP, width=1, corner_radius=0).place(
            relx=1.0, rely=0, relheight=1, anchor="ne"
        )

        # ── Brand block
        brand = ctk.CTkFrame(sidebar, fg_color="transparent")
        brand.grid(row=0, column=0, sticky="ew", padx=20, pady=(28, 22))

        badge_row = ctk.CTkFrame(brand, fg_color="transparent")
        badge_row.pack(anchor="w")
        badge = ctk.CTkFrame(badge_row, fg_color=self.BLUE, corner_radius=10, width=36, height=36)
        badge.pack(side="left")
        badge.pack_propagate(False)
        ctk.CTkLabel(
            badge, text="SL", font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            text_color="#FFFFFF", fg_color="transparent",
        ).place(relx=0.5, rely=0.5, anchor="center")

        brand_text = ctk.CTkFrame(badge_row, fg_color="transparent")
        brand_text.pack(side="left", padx=(10, 0))
        ctk.CTkLabel(
            brand_text, text="Shrey Logistics", font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            text_color=self.T_PRIMARY, anchor="w",
        ).pack(anchor="w")

        ctk.CTkLabel(
            brand, text="Business Automation Suite", font=self.F_TINY,
            text_color=self.T_MUTED, anchor="w",
        ).pack(anchor="w", pady=(6, 0))

        ctk.CTkFrame(sidebar, fg_color=self.SEP, height=1, corner_radius=0).grid(
            row=1, column=0, sticky="ew", padx=20, pady=(0, 18)
        )

        # ── Nav section
        nav = ctk.CTkFrame(sidebar, fg_color="transparent")
        nav.grid(row=2, column=0, sticky="ew", padx=14)

        ctk.CTkLabel(
            nav, text="TOOLS", font=self.F_SECTION, text_color=self.T_MUTED, anchor="w",
        ).pack(fill="x", padx=6, pady=(0, 10))

        self.nav_buttons["xml"] = self._make_nav_button(
            nav, "X", "Tally XML Converter", "Excel → Tally vouchers", "xml"
        )
        self.nav_buttons["bill"] = self._make_nav_button(
            nav, "B", "Bill Generator", "Excel → Word invoice", "bill"
        )

        # ── Footer
        footer = ctk.CTkFrame(sidebar, fg_color="transparent")
        footer.grid(row=4, column=0, sticky="ew", pady=(0, 20))
        ctk.CTkFrame(sidebar, fg_color=self.SEP, height=1, corner_radius=0).grid(
            row=3, column=0, sticky="ew", padx=20, pady=(0, 12)
        )
        ctk.CTkLabel(
            footer, text=f"v{APP_VERSION}  ·  Shrey Logistics", font=self.F_TINY, text_color=self.T_MUTED,
        ).pack()

        update_link = ctk.CTkButton(
            footer,
            text="Check for updates",
            font=self.F_TINY,
            fg_color="transparent",
            text_color=self.BLUE,
            hover_color=self.SURFACE2,
            height=20,
            command=lambda: check_for_updates_async(self, silent=False)
        )
        update_link.pack(pady=(4, 0))

    def _make_nav_button(self, parent, icon, title, subtitle, key):
        """A clickable sidebar row: accent bar + icon badge + title/subtitle.

        Built from plain frames/labels (rather than a CTkButton) so the two
        text lines can have different weights/colors and a left accent bar
        can indicate the active tool — closer to a native sidebar control.
        """
        accent_color = self.ACCENT.get(key, self.BLUE)

        wrap = ctk.CTkFrame(parent, fg_color="transparent", corner_radius=10, height=58)
        wrap.pack(fill="x", pady=3)
        wrap.pack_propagate(False)
        wrap.grid_columnconfigure(2, weight=1)
        wrap.grid_rowconfigure(0, weight=1)

        accent = ctk.CTkFrame(wrap, fg_color="transparent", width=3, corner_radius=2)
        accent.grid(row=0, column=0, sticky="ns", padx=(0, 0), pady=8)

        badge = ctk.CTkFrame(wrap, fg_color=self.SURFACE2, corner_radius=9, width=32, height=32)
        badge.grid(row=0, column=1, padx=(12, 10), pady=0)
        badge.pack_propagate(False)
        icon_lbl = ctk.CTkLabel(
            badge, text=icon, font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            text_color=self.T_SECONDARY, fg_color="transparent",
        )
        icon_lbl.place(relx=0.5, rely=0.5, anchor="center")

        text_col = ctk.CTkFrame(wrap, fg_color="transparent")
        text_col.grid(row=0, column=2, sticky="ew", pady=9)
        title_lbl = ctk.CTkLabel(
            text_col, text=title, font=self.F_LABEL_SB, text_color=self.T_PRIMARY, anchor="w",
        )
        title_lbl.pack(anchor="w")
        sub_lbl = ctk.CTkLabel(
            text_col, text=subtitle, font=self.F_TINY, text_color=self.T_MUTED, anchor="w",
        )
        sub_lbl.pack(anchor="w")

        record = {
            "wrap": wrap, "accent": accent, "badge": badge,
            "icon": icon_lbl, "title": title_lbl, "sub": sub_lbl,
            "accent_color": accent_color,
        }

        def on_click(_event=None):
            self.select_tool(key)

        def on_enter(_event=None):
            if self.active_tool != key:
                wrap.configure(fg_color=self.SURFACE2)

        def on_leave(_event=None):
            if self.active_tool != key:
                wrap.configure(fg_color="transparent")

        for widget in (wrap, badge, icon_lbl, text_col, title_lbl, sub_lbl):
            widget.bind("<Button-1>", on_click)
            widget.bind("<Enter>", on_enter)
            widget.bind("<Leave>", on_leave)
            try:
                widget.configure(cursor="hand2")
            except Exception:
                pass

        return record

    def select_tool(self, key):
        """Switch the visible tool in-place — no separate window or gate screen."""
        if key == "xml":
            if self.xml_frame_root is None:
                self._build_xml_tool()
            self.xml_frame_root.grid(row=0, column=0, sticky="nsew")
            self.xml_frame_root.tkraise()
        else:
            if self.bill_frame_root is None:
                self._build_bill_tool()
            self.bill_frame_root.grid(row=0, column=0, sticky="nsew")
            self.bill_frame_root.tkraise()

        self.active_tool = key
        self._update_nav_highlight()

    def _update_nav_highlight(self):
        for key, rec in self.nav_buttons.items():
            active = key == self.active_tool
            accent_color = rec["accent_color"]
            rec["wrap"].configure(fg_color=self.SURFACE2 if active else "transparent")
            rec["accent"].configure(fg_color=accent_color if active else "transparent")
            rec["badge"].configure(fg_color=accent_color if active else self.SURFACE2)
            rec["icon"].configure(text_color="#FFFFFF" if active else self.T_SECONDARY)
            rec["title"].configure(text_color=self.T_PRIMARY)

    def _tool_header(self, parent, icon, title, subtitle, key):
        accent_color = self.ACCENT.get(key, self.BLUE)
        header = ctk.CTkFrame(parent, fg_color=self.SURFACE, corner_radius=0, border_width=0)
        header.grid(row=0, column=0, sticky="ew")
        ctk.CTkFrame(header, fg_color=self.SEP, height=1, corner_radius=0).pack(side="bottom", fill="x")
        inner = ctk.CTkFrame(header, fg_color="transparent")
        inner.pack(fill="x", padx=28, pady=18)

        badge = ctk.CTkFrame(inner, fg_color=accent_color, corner_radius=10, width=38, height=38)
        badge.pack(side="left", padx=(0, 14))
        badge.pack_propagate(False)
        ctk.CTkLabel(
            badge, text=icon, font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            text_color="#FFFFFF", fg_color="transparent",
        ).place(relx=0.5, rely=0.5, anchor="center")

        title_col = ctk.CTkFrame(inner, fg_color="transparent")
        title_col.pack(side="left")
        ctk.CTkLabel(title_col, text=title, font=self.F_TITLE, text_color=self.T_PRIMARY, anchor="w").pack(anchor="w")
        ctk.CTkLabel(title_col, text=subtitle, font=self.F_SUBTITLE, text_color=self.T_SECONDARY, anchor="w").pack(anchor="w")

    # ============================================================
    # XML TOOL  —  built once, kept alive, swapped via tkraise
    # ============================================================

    def _build_xml_tool(self):
        self.xml_frame_root = ctk.CTkFrame(self.content_area, fg_color=self.BG, corner_radius=0)
        self.xml_frame_root.grid_rowconfigure(1, weight=1)
        self.xml_frame_root.grid_columnconfigure(0, weight=1)

        self._tool_header(
            self.xml_frame_root,
            "X",
            "Tally XML Converter",
            "Convert Excel accounting data into Tally-compatible XML vouchers",
            "xml",
        )

        scroll = ctk.CTkScrollableFrame(self.xml_frame_root, fg_color=self.BG, corner_radius=0)
        scroll.grid(row=1, column=0, sticky="nsew")
        scroll.grid_columnconfigure(0, weight=1)
        self.xml_frame = scroll

        # Status bar (chips + message) for this tool.
        status_bar = ctk.CTkFrame(self.xml_frame_root, fg_color=self.SURFACE, corner_radius=0)
        status_bar.grid(row=2, column=0, sticky="ew")
        ctk.CTkFrame(status_bar, fg_color=self.SEP, height=1, corner_radius=0).pack(side="top", fill="x")
        status_inner = ctk.CTkFrame(status_bar, fg_color="transparent")
        status_inner.pack(fill="x", padx=24, pady=9)

        self._status_dot = ctk.CTkLabel(status_inner, text="●", font=self.F_TINY, text_color=self.T_MUTED)
        self._status_dot.pack(side="left", padx=(0, 6))
        self.status = ctk.CTkLabel(status_inner, text="Ready", font=self.F_STATUS, text_color=self.T_SECONDARY, anchor="w")
        self.status.pack(side="left")

        chip_frame = ctk.CTkFrame(status_inner, fg_color="transparent")
        chip_frame.pack(side="right")
        self._chip_file = self._make_chip(chip_frame, "No file")
        self._chip_rows = self._make_chip(chip_frame, "—")
        self._chip_map = self._make_chip(chip_frame, "No mapping")

        # These legacy hooks are superseded by the shell above; keep as no-ops
        # in case any older call site still references them.
        self._build_titlebar = lambda: None
        self._build_scroll_workspace = lambda: None
        self._build_statusbar = lambda: None

        self._xml_build_ui(scroll)

    def _xml_build_ui(self, parent):
        W = parent
        H = 24
        self._build_config_row(W, H)
        self._hsep(W, H, top=8, bottom=0)
        self._build_step_excel(W, H)
        self._hsep(W, H)
        self._build_step_sheet(W, H)
        self._hsep(W, H)
        self._build_step_columns(W, H)
        self._hsep(W, H)
        self._build_step_mapping_file(W, H)
        self._hsep(W, H)
        self._build_step_actions(W, H)
        self._hsep(W, H)
        self._build_output_panels(W, H)
        ctk.CTkFrame(W, fg_color="transparent", height=20).pack(fill="x")

    # ============================================================
    # BILL GENERATOR TOOL  —  built once, kept alive, swapped via tkraise
    # ============================================================

    def _build_bill_tool(self):
        self.bill_frame_root = ctk.CTkFrame(self.content_area, fg_color=self.BG, corner_radius=0)
        self.bill_frame_root.grid_rowconfigure(1, weight=1)
        self.bill_frame_root.grid_columnconfigure(0, weight=1)

        self._tool_header(
            self.bill_frame_root,
            "B",
            "Bill Generator",
            "Generate formatted Word invoices from Excel source data",
            "bill",
        )

        body = ctk.CTkScrollableFrame(self.bill_frame_root, fg_color=self.BG, corner_radius=0)
        body.grid(row=1, column=0, sticky="nsew")
        body.grid_columnconfigure(0, weight=1)

        status_bar = ctk.CTkFrame(self.bill_frame_root, fg_color=self.SURFACE, corner_radius=0)
        status_bar.grid(row=2, column=0, sticky="ew")
        ctk.CTkFrame(status_bar, fg_color=self.SEP, height=1, corner_radius=0).pack(side="top", fill="x")
        status_inner = ctk.CTkFrame(status_bar, fg_color="transparent")
        status_inner.pack(fill="x", padx=24, pady=9)
        self._bill_status_dot = ctk.CTkLabel(status_inner, text="●", font=self.F_TINY, text_color=self.T_MUTED)
        self._bill_status_dot.pack(side="left", padx=(0, 6))
        self.bill_status = ctk.CTkLabel(
            status_inner, text="Ready", font=self.F_STATUS, text_color=self.T_SECONDARY, anchor="w",
        )
        self.bill_status.pack(side="left")

        self._build_bill_generator_section(body, 24)

    # ============================================================
    # TITLE BAR
    # ============================================================

    def _build_titlebar(self):
        bar = ctk.CTkFrame(self, fg_color=self.SURFACE, corner_radius=0,
                           border_width=0)
        bar.grid(row=0, column=0, sticky="ew")

        # bottom separator
        sep = ctk.CTkFrame(bar, fg_color=self.SEP, height=1, corner_radius=0)
        sep.pack(side="bottom", fill="x")

        inner = ctk.CTkFrame(bar, fg_color="transparent")
        inner.pack(fill="x", padx=24, pady=12)

        # Company logo — uses the supplied Shrey Logistics logo.
        # PhotoImage keeps this dependency-free (no Pillow required).
        logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "shrey_logistics_logo.png")
        self._logo_image = None
        if os.path.exists(logo_path):
            try:
                logo = PhotoImage(file=logo_path)
                # Original logo is wide; 2x downsample keeps it crisp and compact.
                self._logo_image = logo.subsample(2, 2)
                icon = ctk.CTkLabel(
                    inner,
                    text="",
                    image=self._logo_image,
                    width=221,
                    height=27,
                    fg_color="transparent",
                )
            except Exception:
                icon = ctk.CTkLabel(
                    inner, text="SL", width=42, height=32, corner_radius=8,
                    fg_color=self.BLUE, text_color="#FFFFFF",
                    font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
                )
        else:
            icon = ctk.CTkLabel(
                inner, text="SL", width=42, height=32, corner_radius=8,
                fg_color=self.BLUE, text_color="#FFFFFF",
                font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            )
        icon.pack(side="left", padx=(0, 16))

        title_col = ctk.CTkFrame(inner, fg_color="transparent")
        title_col.pack(side="left")

        ctk.CTkLabel(
            title_col,
            text="Tally Excel → XML Converter",
            font=self.F_TITLE,
            text_color=self.T_PRIMARY,
            anchor="w",
        ).pack(anchor="w")

        ctk.CTkLabel(
            title_col,
            text="Convert Excel accounting data into Tally-compatible XML vouchers",
            font=self.F_SUBTITLE,
            text_color=self.T_SECONDARY,
            anchor="w",
        ).pack(anchor="w")

    # ============================================================
    # SCROLLABLE WORKSPACE
    # ============================================================

    def _build_scroll_workspace(self):
        self.scroll_frame = ctk.CTkScrollableFrame(
            self,
            fg_color=self.BG,
            corner_radius=0,
            scrollbar_button_color=self.SEP,
            scrollbar_button_hover_color=self.T_MUTED,
        )
        self.scroll_frame.grid(row=1, column=0, sticky="nsew")
        self.scroll_frame.grid_columnconfigure(0, weight=1)

        # Faster, smoother Windows mouse-wheel scrolling.
        self._bind_fast_scroll()

        W = self.scroll_frame   # workspace
        H = 24                  # horizontal margin

        # ── Configuration row (compact, not a card)
        self._build_config_row(W, H)
        self._hsep(W, H, top=8, bottom=0)

        # ── Steps
        self._build_step_excel(W, H)
        self._hsep(W, H)

        self._build_step_sheet(W, H)
        self._hsep(W, H)

        self._build_step_columns(W, H)
        self._hsep(W, H)

        self._build_step_mapping_file(W, H)
        self._hsep(W, H)

        self._build_step_actions(W, H)
        self._hsep(W, H)

        # ── Output panels  (side by side)
        self._build_output_panels(W, H)

        # breathing room at bottom
        ctk.CTkFrame(W, fg_color="transparent", height=20).pack(fill="x")

    # ── Faster mouse-wheel scrolling
    def _bind_fast_scroll(self):
        canvas = getattr(self.scroll_frame, "_parent_canvas", None)
        if canvas is None:
            return

        def on_mousewheel(event):
            delta = event.delta
            if delta == 0:
                return "break"
            # Windows normally sends +/-120 per wheel notch.  4 units gives
            # a noticeably faster scroll without making the page jump.
            units = max(1, int(abs(delta) / 30))
            units *= -1 if delta > 0 else 1
            canvas.yview_scroll(units, "units")
            return "break"

        canvas.bind("<MouseWheel>", on_mousewheel, add="+")

    # ── Thin horizontal separator
    def _hsep(self, parent, h_pad, top=0, bottom=0):
        ctk.CTkFrame(
            parent,
            fg_color=self.SEP,
            height=1,
            corner_radius=0,
        ).pack(fill="x", padx=h_pad, pady=(top, bottom))

    # ============================================================
    # CONFIG ROW  —  Company + Ledger inline
    # ============================================================

    def _build_config_row(self, parent, h_pad):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", padx=h_pad, pady=(16, 12))
        row.grid_columnconfigure(1, weight=2)
        row.grid_columnconfigure(3, weight=1)

        ctk.CTkLabel(
            row, text="Company",
            font=self.F_SMALL, text_color=self.T_SECONDARY,
        ).grid(row=0, column=0, padx=(0, 8), sticky="w")

        self.company_entry = ctk.CTkEntry(
            row,
            height=34,
            fg_color=self.INPUT_BG,
            border_color=self.INPUT_BDR,
            border_width=1,
            corner_radius=8,
            font=self.F_BODY,
            text_color=self.T_PRIMARY,
        )
        self.company_entry.insert(0, "SHREY LOGISTICS [FINAL FY : 24-25]")
        self.company_entry.grid(row=0, column=1, padx=(0, 20), sticky="ew")

        ctk.CTkLabel(
            row, text="Debit Ledger",
            font=self.F_SMALL, text_color=self.T_SECONDARY,
        ).grid(row=0, column=2, padx=(0, 8), sticky="w")

        self.ledger_entry = ctk.CTkEntry(
            row,
            height=34,
            fg_color=self.INPUT_BG,
            border_color=self.INPUT_BDR,
            border_width=1,
            corner_radius=8,
            font=self.F_BODY,
            text_color=self.T_PRIMARY,
        )
        self.ledger_entry.insert(0, "Transport Expense")
        self.ledger_entry.grid(row=0, column=3, sticky="ew")

    # ============================================================
    # STEP HELPERS
    # ============================================================

    def _step_row(self, parent, h_pad, num, title, top_pad=16):
        """Returns (outer_frame) — lays out step number + title + content area."""
        outer = ctk.CTkFrame(parent, fg_color="transparent")
        outer.pack(fill="x", padx=h_pad, pady=(top_pad, 10))
        outer.grid_columnconfigure(1, weight=1)

        # Step badge — filled circle with number
        badge = ctk.CTkFrame(
            outer,
            fg_color=self.SEP,
            corner_radius=11,
            width=22,
            height=22,
        )
        badge.grid(row=0, column=0, sticky="nw", padx=(0, 14), pady=(2, 0))
        badge.grid_propagate(False)
        ctk.CTkLabel(
            badge,
            text=num,
            font=self.F_TINY,
            text_color=self.T_SECONDARY,
            fg_color="transparent",
        ).place(relx=0.5, rely=0.5, anchor="center")

        # Title — headline weight
        ctk.CTkLabel(
            outer,
            text=title,
            font=self.F_LABEL_SB,
            text_color=self.T_PRIMARY,
            anchor="w",
        ).grid(row=0, column=1, sticky="w")

        # Content frame below title, indented to align with title
        content = ctk.CTkFrame(outer, fg_color="transparent")
        content.grid(row=1, column=1, sticky="ew", pady=(8, 0))
        content.grid_columnconfigure(0, weight=1)

        return content

    def _ghost_label(self, parent, text):
        """Pill-style read-only value display — inset grouped style."""
        return ctk.CTkLabel(
            parent,
            text=text,
            anchor="w",
            font=self.F_BODY,
            text_color=self.T_MUTED,
            fg_color=self.INPUT_BG,
            corner_radius=8,
            height=34,
            padx=12,
        )

    def _secondary_btn(self, parent, text, command, width=110):
        return ctk.CTkButton(
            parent,
            text=text,
            command=command,
            width=width,
            height=34,
            corner_radius=8,
            font=self.F_BTN,
            fg_color=self.SURFACE,
            text_color=self.BLUE,
            border_width=1,
            border_color=self.SEP,
            hover_color=self.BG,
        )

    def _primary_btn(self, parent, text, command, width=140, state="normal"):
        return ctk.CTkButton(
            parent,
            text=text,
            command=command,
            width=width,
            height=36,
            corner_radius=10,
            font=self.F_BTN_PRI,
            fg_color=self.BLUE,
            hover_color=self.BLUE_HOVER,
            text_color="#FFFFFF",
            state=state,
        )

    def _combo(self, parent, values, command=None, width=280):
        kw = dict(
            values=values,
            width=width,
            height=34,
            corner_radius=8,
            fg_color=self.INPUT_BG,
            border_color=self.INPUT_BDR,
            border_width=1,
            button_color=self.INPUT_BDR,
            button_hover_color=self.T_MUTED,
            dropdown_fg_color=self.SURFACE,
            dropdown_hover_color=self.BG,
            dropdown_text_color=self.T_PRIMARY,
            text_color=self.T_PRIMARY,
            font=self.F_BODY,
        )
        if command:
            kw["command"] = command
        return ctk.CTkComboBox(parent, **kw)

    # ============================================================
    # STEP 01 — EXCEL FILE
    # ============================================================

    def _build_step_excel(self, parent, h_pad):
        content = self._step_row(parent, h_pad, "01", "Excel File")
        content.grid_columnconfigure(0, weight=1)

        row = ctk.CTkFrame(content, fg_color="transparent")
        row.pack(fill="x")
        row.grid_columnconfigure(0, weight=1)

        self.excel_label = self._ghost_label(row, "No file selected")
        self.excel_label.grid(row=0, column=0, sticky="ew", padx=(0, 10))

        self._secondary_btn(row, "Choose…", self.select_excel, width=90
                            ).grid(row=0, column=1)

    # ============================================================
    # STEP 02 — SHEET
    # ============================================================

    def _build_step_sheet(self, parent, h_pad):
        content = self._step_row(parent, h_pad, "02", "Sheet")

        row = ctk.CTkFrame(content, fg_color="transparent")
        row.pack(fill="x")

        self.sheet_menu = self._combo(row, ["Select Excel first"],
                                      command=self.load_sheet, width=300)
        self.sheet_menu.pack(side="left", padx=(0, 16))

        self.sheet_status = ctk.CTkLabel(
            row, text="",
            font=self.F_SMALL, text_color=self.T_MUTED, anchor="w")
        self.sheet_status.pack(side="left", anchor="w")

    # ============================================================
    # STEP 03 — COLUMN MAPPING
    # ============================================================

    def _build_step_columns(self, parent, h_pad):
        content = self._step_row(parent, h_pad, "03", "Column Mapping")

        grid = ctk.CTkFrame(content, fg_color="transparent")
        grid.pack(anchor="w")

        FIELDS = [
            ("Date",           "date_menu"),
            ("L.R.NO",         "lr_menu"),
            ("Broker",         "broker_menu"),
            ("Amount / Total", "amount_menu"),
        ]

        for i, (label_text, attr) in enumerate(FIELDS):
            col = i % 2
            row = i // 2

            cell = ctk.CTkFrame(grid, fg_color="transparent")
            cell.grid(row=row, column=col, padx=(0, 32), pady=(0, 8), sticky="w")

            ctk.CTkLabel(
                cell, text=label_text,
                font=self.F_SMALL, text_color=self.T_SECONDARY,
                anchor="w", width=110,
            ).pack(anchor="w")

            combo = self._combo(cell, ["Load Excel first"], width=220)
            combo.pack(anchor="w", pady=(2, 0))
            setattr(self, attr, combo)

    # ============================================================
    # STEP 04 — BROKER MAPPING FILE
    # ============================================================

    def _build_step_mapping_file(self, parent, h_pad):
        content = self._step_row(parent, h_pad, "04", "Broker Mapping File")
        content.grid_columnconfigure(0, weight=1)

        row = ctk.CTkFrame(content, fg_color="transparent")
        row.pack(fill="x")
        row.grid_columnconfigure(0, weight=1)

        self.mapping_label = self._ghost_label(row, "No file selected")
        self.mapping_label.grid(row=0, column=0, sticky="ew", padx=(0, 10))

        self._secondary_btn(row, "Choose…", self.select_mapping, width=90
                            ).grid(row=0, column=1)

    # ============================================================
    # STEP 05 — VALIDATE, CHECK & GENERATE
    # ============================================================

    def _build_step_actions(self, parent, h_pad):
        content = self._step_row(parent, h_pad, "05", "Validate & Generate")

        # Test mode toggle row
        toggle_section = ctk.CTkFrame(content, fg_color="transparent")
        toggle_section.pack(fill="x", pady=(0, 14))

        self.test_first_var = ctk.BooleanVar(value=True)
        ctk.CTkSwitch(
            toggle_section,
            text="Test mode  —  generate first voucher only",
            variable=self.test_first_var,
            font=self.F_SMALL,
            text_color=self.T_SECONDARY,
            progress_color=self.BLUE,
            button_color=("#FFFFFF", "#FFFFFF"),
            button_hover_color=("#F0F0F0", "#F0F0F0"),
            fg_color=self.SEP,
            switch_width=36,
            switch_height=20,
        ).pack(anchor="w")

        # Action buttons row
        btn_row = ctk.CTkFrame(content, fg_color="transparent")
        btn_row.pack(fill="x", pady=(0, 4))

        self._secondary_btn(btn_row, "Preview", self.preview_data, width=90
                            ).pack(side="left", padx=(0, 8))

        self._secondary_btn(btn_row, "Validate Data", self.validate_data, width=112
                            ).pack(side="left", padx=(0, 8))

        self.mapping_button = self._secondary_btn(
            btn_row, "Check Mapping", self.check_mapping, width=120)
        self.mapping_button.configure(state="disabled")
        self.mapping_button.pack(side="left", padx=(0, 20))

        # Generate — prominent, separated from secondary actions
        self.generate_button = ctk.CTkButton(
            btn_row,
            text="Generate XML  →",
            command=self.generate_xml,
            width=158,
            height=36,
            corner_radius=10,
            font=self.F_BTN_PRI,
            fg_color=self.BLUE,
            hover_color=self.BLUE_HOVER,
            text_color="#FFFFFF",
            state="disabled",
        )
        self.generate_button.pack(side="left")

        # Tally import hint — footnote style
        ctk.CTkLabel(
            content,
            text="To import: Gateway of Tally → Import Data → Vouchers",
            font=self.F_TINY,
            text_color=self.T_MUTED,
            anchor="w",
        ).pack(anchor="w", pady=(10, 0))

    # ============================================================
    # OUTPUT PANELS  (mapping status + data preview)
    # ============================================================

    # ============================================================
    # BILL GENERATOR WORKFLOW
    # ============================================================

    def _build_bill_generator_section(self, parent, h_pad):
        content = self._step_row(parent, h_pad, "01", "Source Excel")
        content.grid_columnconfigure(0, weight=1)

        row = ctk.CTkFrame(content, fg_color="transparent")
        row.pack(fill="x")
        row.grid_columnconfigure(0, weight=1)
        self.bill_excel_label = self._ghost_label(row, "No file selected")
        self.bill_excel_label.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        self._secondary_btn(row, "Choose Excel…", self.select_bill_excel, width=110).grid(row=0, column=1)

        content2 = self._step_row(parent, h_pad, "01", "Billing Period")
        content2.grid_columnconfigure(0, weight=1)

        period = ctk.CTkFrame(content2, fg_color="transparent")
        period.pack(fill="x")
        period.grid_columnconfigure(1, weight=1)
        period.grid_columnconfigure(3, weight=1)

        ctk.CTkLabel(period, text="From", font=self.F_SMALL, text_color=self.T_SECONDARY).grid(row=0, column=0, sticky="w", padx=(0, 10))
        self.bill_from_entry = CalendarDateEntry(period, callback=self._refresh_bill_company_options, width=140, height=30)
        self.bill_from_entry.grid(row=0, column=1, sticky="w", padx=(0, 20))
        ctk.CTkLabel(period, text="To", font=self.F_SMALL, text_color=self.T_SECONDARY).grid(row=0, column=2, sticky="w", padx=(0, 10))
        self.bill_to_entry = CalendarDateEntry(period, callback=self._refresh_bill_company_options, width=140, height=30)
        self.bill_to_entry.grid(row=0, column=3, sticky="w")

        marker_pad = ctk.CTkFrame(content2, fg_color="transparent")
        marker_pad.pack(fill="x", pady=(10, 0))
        ctk.CTkLabel(marker_pad, text="Excel BILL marker", font=self.F_SMALL, text_color=self.T_SECONDARY).pack(anchor="w")
        self.bill_marker_combo = self._combo(marker_pad, BILL_MARKER_OPTIONS, width=120)
        self.bill_marker_combo.set("+")
        self.bill_marker_combo.pack(anchor="w", pady=(4, 0))
        self.bill_marker_combo.bind("<<ComboboxSelected>>", lambda _event: self._refresh_bill_company_options())
        self.bill_marker_combo.bind("<FocusOut>", lambda _event: self._refresh_bill_company_options())

        content3 = self._step_row(parent, h_pad, "02", "Company / SEN.UNIT")
        content3.grid_columnconfigure(0, weight=1)
        self.bill_company_combo = self._combo(content3, ["Select billing period first"], width=320)
        self.bill_company_combo.configure(state="disabled")
        self.bill_company_combo.pack(anchor="w")

        content4 = self._step_row(parent, h_pad, "03", "Bill Details")
        content4.grid_columnconfigure(0, weight=1)
        detail_grid = ctk.CTkFrame(content4, fg_color="transparent")
        detail_grid.pack(fill="x")
        detail_grid.grid_columnconfigure(1, weight=1)
        detail_grid.grid_columnconfigure(3, weight=1)

        ctk.CTkLabel(detail_grid, text="Bill No.", font=self.F_SMALL, text_color=self.T_SECONDARY).grid(row=0, column=0, sticky="w", padx=(0, 10))
        self.bill_no_entry = ctk.CTkEntry(detail_grid, width=180, height=34, corner_radius=8, fg_color=self.INPUT_BG, border_color=self.INPUT_BDR, border_width=1, font=self.F_BODY, text_color=self.T_PRIMARY)
        self.bill_no_entry.grid(row=0, column=1, sticky="w", padx=(0, 18))
        ctk.CTkLabel(detail_grid, text="Bill Date", font=self.F_SMALL, text_color=self.T_SECONDARY).grid(row=0, column=2, sticky="w", padx=(0, 10))
        self.bill_date_entry = ctk.CTkEntry(detail_grid, width=150, height=34, corner_radius=8, fg_color=self.INPUT_BG, border_color=self.INPUT_BDR, border_width=1, font=self.F_BODY, text_color=self.T_PRIMARY)
        self.bill_date_entry.grid(row=0, column=3, sticky="w")
        ctk.CTkLabel(detail_grid, text="SAC Code", font=self.F_SMALL, text_color=self.T_SECONDARY).grid(row=1, column=0, sticky="w", pady=(12, 0), padx=(0, 10))
        self.bill_sac_entry = ctk.CTkEntry(detail_grid, width=180, height=34, corner_radius=8, fg_color=self.INPUT_BG, border_color=self.INPUT_BDR, border_width=1, font=self.F_BODY, text_color=self.T_PRIMARY)
        self.bill_sac_entry.insert(0, DEFAULT_BILL_SAC_CODE)
        self.bill_sac_entry.grid(row=1, column=1, sticky="w", pady=(12, 0), padx=(0, 18))

        content5 = self._step_row(parent, h_pad, "04", "Company / Origin Mapping")
        content5.grid_columnconfigure(0, weight=1)
        mapping_row = ctk.CTkFrame(content5, fg_color="transparent")
        mapping_row.pack(fill="x")
        mapping_row.grid_columnconfigure(0, weight=1)
        self.bill_mapping_label = self._ghost_label(mapping_row, "No file selected")
        self.bill_mapping_label.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        self._secondary_btn(mapping_row, "Choose Mapping…", self.select_bill_mapping, width=140).grid(row=0, column=1)

        content6 = self._step_row(parent, h_pad, "05", "Word Template")
        content6.grid_columnconfigure(0, weight=1)
        template_row = ctk.CTkFrame(content6, fg_color="transparent")
        template_row.pack(fill="x")
        template_row.grid_columnconfigure(0, weight=1)
        self.bill_template_label = self._ghost_label(template_row, "No template selected")
        self.bill_template_label.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        self._secondary_btn(template_row, "Choose Word Template...", self.select_bill_template, width=140).grid(row=0, column=1)

        content7 = self._step_row(parent, h_pad, "06", "Generate")
        content7.grid_columnconfigure(0, weight=1)
        self.generate_bill_button = self._primary_btn(content7, "Generate Word Bill →", self.generate_bill_from_ui, width=190)
        self.generate_bill_button.pack(anchor="w")

    def _refresh_bill_company_options(self, event=None):
        if not self.bill_excel_path:
            self.bill_company_combo.configure(values=["Select Excel first"])
            self.bill_company_combo.set("Select Excel first")
            self.bill_company_combo.configure(state="disabled")
            return

        start_date = self.bill_from_entry.get()
        end_date = self.bill_to_entry.get()
        if not start_date or not end_date:
            self.bill_company_combo.configure(values=["Select billing period first"])
            self.bill_company_combo.set("Select billing period first")
            self.bill_company_combo.configure(state="disabled")
            return

        start_dt = parse_date(start_date)
        end_dt = parse_date(end_date)
        if start_dt is None or end_dt is None:
            self.bill_company_combo.configure(values=["Invalid billing period"])
            self.bill_company_combo.set("Invalid billing period")
            self.bill_company_combo.configure(state="disabled")
            return

        if start_dt > end_dt:
            self.bill_company_combo.configure(values=["From date must be before To date"])
            self.bill_company_combo.set("From date must be before To date")
            self.bill_company_combo.configure(state="disabled")
            return

        names = self.bill_service.get_company_names_for_period(
            self.bill_excel_path,
            start_date,
            end_date,
            self.bill_marker_combo.get().strip() or "+",
        )

        if not names:
            self.bill_company_combo.configure(values=["No company data for period"])
            self.bill_company_combo.set("No company data for period")
            self.bill_company_combo.configure(state="disabled")
            return

        self.bill_company_combo.configure(values=names)
        self.bill_company_combo.set(names[0] if len(names) == 1 else "Select company")
        self.bill_company_combo.configure(state="normal")

    def select_bill_excel(self):
        path = filedialog.askopenfilename(title="Select Excel Source", filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")])
        if not path:
            return
        self.bill_excel_path = path
        self.bill_excel_label.configure(text=os.path.basename(path), text_color=self.T_PRIMARY)
        self._refresh_bill_company_options()

    def select_bill_mapping(self):
        path = filedialog.askopenfilename(title="Select Company Mapping File", filetypes=[("Excel Files", "*.xlsx *.xls *.csv"), ("All Files", "*.*")])
        if not path:
            return
        self.bill_mapping_path = path
        self.bill_mapping_label.configure(text=os.path.basename(path), text_color=self.T_PRIMARY)

    def select_bill_template(self):
        path = filedialog.askopenfilename(
            title="Select Word Bill Template",
            filetypes=[("Word Documents", "*.docx *.dotx"), ("All Files", "*.*")],
        )
        if not path:
            return
        self.bill_template_path = path
        self.bill_template_label.configure(text=os.path.basename(path), text_color=self.T_PRIMARY)
        self._set_status("✓ Template selected", self.GREEN)

    def generate_bill_from_ui(self):
        excel_path = self.bill_excel_path.strip()
        if not excel_path:
            messagebox.showwarning("Validation", "Select a source Excel file first.")
            return

        company_name = self.bill_company_combo.get().strip()
        if not company_name or company_name == "Select company" or company_name == "No company names detected":
            messagebox.showwarning("Validation", "Select a company / SEN.UNIT.")
            return

        start_date = self.bill_from_entry.get().strip()
        end_date = self.bill_to_entry.get().strip()
        if not start_date or not end_date:
            messagebox.showwarning("Validation", "Enter a valid billing date range.")
            return
        if parse_date(start_date) is None or parse_date(end_date) is None:
            messagebox.showwarning("Validation", "Billing dates are invalid. Use formats like 25.05.26.")
            return
        if parse_date(start_date) > parse_date(end_date):
            messagebox.showwarning("Validation", "From date must be on or before To date.")
            return

        marker = self.bill_marker_combo.get().strip()
        if not marker:
            messagebox.showwarning("Validation", "Select an Excel bill marker.")
            return

        bill_no = self.bill_no_entry.get().strip()
        if not bill_no:
            messagebox.showwarning("Validation", "Enter the Bill No.")
            return

        bill_date = self.bill_date_entry.get().strip()
        if not bill_date or parse_date(bill_date) is None:
            messagebox.showwarning("Validation", "Enter a valid Bill Date.")
            return

        mapping_path = self.bill_mapping_path or self.bill_service.ensure_default_mapping()
        if not os.path.exists(mapping_path):
            messagebox.showwarning("Validation", "Company mapping file is missing.")
            return

        template_path = self.bill_template_path.strip() if self.bill_template_path else ""
        if not template_path:
            messagebox.showwarning("Validation", "Please choose a Word template first.")
            return
        if not os.path.exists(template_path):
            messagebox.showwarning("Validation", "Selected template file is missing. Please choose a valid Word template.")
            return

        mapping = self.bill_service.load_mapping_file(mapping_path)
        if self.bill_service.normalize_company_name(company_name) not in mapping:
            messagebox.showwarning("Validation", f"No company mapping found for '{company_name}'.")
            return

        rows, debug = self.bill_service.get_bill_rows(excel_path, company_name, start_date, end_date, marker, return_debug=True)
        if not rows:
            messagebox.showwarning(
                "Validation",
                "No transactions matched the selected company, date range, and BILL marker.\n\n"
                f"Debug: company rows={debug['company_matches']}, date rows={debug['date_matches']}, bill rows={debug['bill_matches']}"
            )
            return

        output_path = filedialog.asksaveasfilename(
            title="Save Bill as Word Document",
            defaultextension=".docx",
            initialfile=f"{company_name}_{bill_no}.docx",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")],
        )
        if not output_path:
            return

        try:
            result = self.bill_service.generate_bill_document(
                excel_path=excel_path,
                company_name=company_name,
                start_date=start_date,
                end_date=end_date,
                bill_marker=marker,
                bill_no=bill_no,
                bill_date=bill_date,
                sac_code=self.bill_sac_entry.get().strip() or DEFAULT_BILL_SAC_CODE,
                mapping_path=mapping_path,
                output_path=output_path,
                template_path=template_path,
            )
            messagebox.showinfo("Bill Generated", f"Bill saved successfully.\n\nRows: {result['rows']}\nRoutes: {result['routes']}\nGrand Total: ₹{result['grand_total']:,.2f}\n\nFile:\n{output_path}")
            self._set_status("✓ Bill generated", self.GREEN)
        except Exception as exc:
            messagebox.showerror("Bill Generation Error", f"Could not generate the bill:\n\n{exc}")
            self._set_status("⚠ Bill generation failed", self.RED)

    def _build_output_panels(self, parent, h_pad):
        outer = ctk.CTkFrame(parent, fg_color="transparent")
        outer.pack(fill="x", padx=h_pad, pady=(0, 0))
        outer.grid_columnconfigure(0, weight=1)
        outer.grid_columnconfigure(1, weight=1)

        # ── Left: Broker Mapping Status
        left_card = ctk.CTkFrame(outer, fg_color=self.SURFACE, corner_radius=12,
                                  border_width=1, border_color=self.SEP)
        left_card.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        self._panel_header(left_card, "Broker Mapping")

        self.mapping_box = ctk.CTkTextbox(
            left_card,
            height=210,
            corner_radius=8,
            fg_color=self.BG,
            border_width=0,
            font=self.F_MONO,
            text_color=self.T_SECONDARY,
            wrap="none",
        )
        self.mapping_box.pack(fill="x", padx=12, pady=(0, 12))
        self.mapping_box.insert("end", "Load a mapping file and click Check Mapping.")

        # ── Right: Data Preview
        right_card = ctk.CTkFrame(outer, fg_color=self.SURFACE, corner_radius=12,
                                   border_width=1, border_color=self.SEP)
        right_card.grid(row=0, column=1, sticky="nsew", padx=(8, 0))

        self._panel_header(right_card, "Data Preview")

        self.preview_box = ctk.CTkTextbox(
            right_card,
            height=210,
            corner_radius=8,
            fg_color=self.BG,
            border_width=0,
            font=self.F_MONO,
            text_color=self.T_SECONDARY,
            wrap="none",
        )
        self.preview_box.pack(fill="x", padx=12, pady=(0, 12))
        self.preview_box.insert("end", "Select a sheet to preview transactions.")

    def _panel_header(self, parent, title):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", padx=12, pady=(12, 6))

        ctk.CTkLabel(
            row,
            text=title.upper(),
            font=self.F_SECTION,
            text_color=self.T_MUTED,
            anchor="w",
        ).pack(anchor="w")

    # ============================================================
    # STATUS BAR
    # ============================================================

    def _build_statusbar(self):
        bar_wrap = ctk.CTkFrame(self, fg_color=self.SURFACE, corner_radius=0)
        bar_wrap.grid(row=2, column=0, sticky="ew")

        # Hairline top separator
        ctk.CTkFrame(bar_wrap, fg_color=self.SEP, height=1, corner_radius=0
                     ).pack(side="top", fill="x")

        inner = ctk.CTkFrame(bar_wrap, fg_color="transparent")
        inner.pack(fill="x", padx=24, pady=9)

        # Dot + main status
        self._status_dot = ctk.CTkLabel(
            inner, text="●",
            font=self.F_TINY, text_color=self.T_MUTED)
        self._status_dot.pack(side="left", padx=(0, 6))

        self.status = ctk.CTkLabel(
            inner, text="Ready",
            font=self.F_STATUS, text_color=self.T_SECONDARY, anchor="w")
        self.status.pack(side="left")

        # Right: info chips — subtle pill badges
        chip_frame = ctk.CTkFrame(inner, fg_color="transparent")
        chip_frame.pack(side="right")

        self._chip_file = self._make_chip(chip_frame, "No file")
        self._chip_rows = self._make_chip(chip_frame, "—")
        self._chip_map  = self._make_chip(chip_frame, "No mapping")

    def _make_chip(self, parent, text):
        chip = ctk.CTkLabel(
            parent,
            text=text,
            font=self.F_TINY,
            text_color=self.T_SECONDARY,
            fg_color=self.BG,
            corner_radius=6,
            padx=10, pady=3,
        )
        chip.pack(side="left", padx=3)
        return chip

    def _update_chips(self):
        if self.excel_path:
            name = os.path.basename(self.excel_path)
            self._chip_file.configure(text=(name[:22] + "…") if len(name) > 22 else name)
        else:
            self._chip_file.configure(text="No file")

        if self.data is not None:
            tx = self.get_transaction_indices()
            self._chip_rows.configure(text=f"{len(tx)} rows")
        else:
            self._chip_rows.configure(text="—")

        if self.mapping_data is not None:
            self._chip_map.configure(text=f"{len(self.mapping_data)} ledgers")
        else:
            self._chip_map.configure(text="No mapping")

    def _set_status(self, text, colour=None):
        c = colour if colour else self.T_SECONDARY
        if getattr(self, "status", None) is not None and self.status.winfo_exists():
            self.status.configure(text=text, text_color=c)
        if getattr(self, "bill_status", None) is not None and self.bill_status.winfo_exists():
            self.bill_status.configure(text=text, text_color=c)

    # ============================================================
    # EXCEL SELECTION
    # ============================================================

    def select_excel(self):

        path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[
                ("Excel Files", "*.xlsx *.xls"),
                ("All Files", "*.*")
            ]
        )

        if not path:
            return

        try:

            self.excel_path = path

            self.workbook = pd.ExcelFile(
                path
            )

            sheets = self.workbook.sheet_names

            if not sheets:
                raise ValueError(
                    "No worksheets found."
                )

            self.sheet_menu.configure(
                values=sheets
            )

            # IMPORTANT:
            # Do NOT automatically choose a sheet.

            self.sheet_menu.set(
                "Select a sheet"
            )

            self.data = None
            self.current_sheet = ""
            self.header_row = None

            self.mapping_button.configure(
                state="disabled"
            )

            self.generate_button.configure(
                state="disabled"
            )

            self.data_valid = False
            self.mapping_valid = False

            self.excel_label.configure(
                text=os.path.basename(path),
                text_color=self.T_PRIMARY,
            )

            self.sheet_status.configure(
                text=f"{len(sheets)} sheet(s) available"
            )

            self.preview_box.delete(
                "1.0",
                "end"
            )

            self.preview_box.insert(
                "end",
                "Excel loaded.\n\n"
                "Select the sheet you want to process."
            )

            self.mapping_box.delete(
                "1.0",
                "end"
            )

            self.mapping_box.insert(
                "end",
                "Select a sheet first."
            )

            self.status.configure(
                text="Select an Excel sheet"
            )

            self._update_chips()

        except Exception as e:

            messagebox.showerror(
                "Excel Error",
                f"Could not open Excel file:\n\n{e}"
            )

    # ============================================================
    # LOAD SELECTED SHEET
    # ============================================================

    def load_sheet(self, sheet_name):

        if (
            not self.excel_path
            or sheet_name == "Select a sheet"
        ):
            return

        try:

            raw = pd.read_excel(
                self.excel_path,
                sheet_name=sheet_name,
                header=None,
                dtype=object
            )

            structure = (
                self.detect_sheet_structure(
                    raw
                )
            )

            if structure is None:

                self.data = None
                self.current_sheet = sheet_name

                self.clear_column_menus()

                self.mapping_button.configure(
                    state="disabled"
                )

                self.generate_button.configure(
                    state="disabled"
                )

                self.sheet_status.configure(
                    text="No accounting header detected"
                )

                self.preview_box.delete(
                    "1.0",
                    "end"
                )

                self.preview_box.insert(
                    "end",
                    f"Sheet: {sheet_name}\n\n"
                    "Could not detect DATE, L.R.NO, "
                    "BROKER and TOTAL/AMOUNT columns."
                )

                return

            header_row = structure[
                "header_row"
            ]

            headers = structure[
                "headers"
            ]

            if (
                header_row + 1
                < len(raw)
            ):

                data = raw.iloc[
                    header_row + 1:
                ].copy()

                data.columns = headers

            else:

                data = pd.DataFrame(
                    columns=headers
                )

            data = self.clean_data(
                data
            )

            self.data = data
            self.current_sheet = sheet_name
            self.header_row = header_row

            date_col = headers[
                structure["date_index"]
            ]

            lr_col = headers[
                structure["lr_index"]
            ]

            broker_col = headers[
                structure["broker_index"]
            ]

            amount_col = headers[
                structure["amount_index"]
            ]

            columns = [
                str(x)
                for x in data.columns
            ]

            self.date_menu.configure(
                values=columns
            )

            self.lr_menu.configure(
                values=columns
            )

            self.broker_menu.configure(
                values=columns
            )

            self.amount_menu.configure(
                values=columns
            )

            self.date_menu.set(
                date_col
            )

            self.lr_menu.set(
                lr_col
            )

            self.broker_menu.set(
                broker_col
            )

            self.amount_menu.set(
                amount_col
            )

            self.date_column = date_col
            self.lr_column = lr_col
            self.broker_column = broker_col
            self.amount_column = amount_col

            self.data_valid = False
            self.mapping_valid = False

            self.generate_button.configure(
                state="disabled"
            )

            if self.mapping_data is not None:

                self.mapping_button.configure(
                    state="normal"
                )

            else:

                self.mapping_button.configure(
                    state="disabled"
                )

            transactions = (
                self.get_transaction_indices()
            )

            self.sheet_status.configure(
                text=(
                    f"Header row {header_row + 1} | "
                    f"{len(transactions)} "
                    f"transaction(s)"
                )
            )

            self.preview_data()

            self.mapping_box.delete(
                "1.0",
                "end"
            )

            self.mapping_box.insert(
                "end",
                (
                    "Sheet loaded successfully.\n"
                    "Click Validate Data, then "
                    "Check Mapping."
                )
            )

            self.status.configure(
                text=(
                    f"{sheet_name} loaded — "
                    f"{len(transactions)} "
                    f"transaction(s)"
                )
            )

            self._update_chips()

        except Exception as e:

            messagebox.showerror(
                "Sheet Error",
                f"Could not load sheet:\n\n{e}"
            )

    # ============================================================
    # SMART HEADER DETECTION
    # ============================================================

    def detect_sheet_structure(
        self,
        raw
    ):

        if raw is None or raw.empty:
            return None

        scan_rows = min(
            40,
            len(raw)
        )

        best = None

        for row_index in range(
            scan_rows
        ):

            row = raw.iloc[
                row_index
            ].tolist()

            score_info = (
                self.score_header_row(
                    row
                )
            )

            if score_info is None:
                continue

            (
                header_score,
                date_index,
                lr_index,
                broker_index,
                amount_index
            ) = score_info

            headers = self.make_unique_headers(
                row
            )

            start = row_index + 1

            transaction_count = 0

            if start < len(raw):

                temp = raw.iloc[
                    start:
                ].copy()

                temp.columns = headers

                transaction_count = (
                    self.count_transactions(
                        temp,
                        headers[date_index],
                        headers[lr_index],
                        headers[broker_index],
                        headers[amount_index]
                    )
                )

            score = (
                header_score * 100000
                + transaction_count * 100
                - row_index
            )

            candidate = {
                "score": score,
                "header_row": row_index,
                "headers": headers,
                "date_index": date_index,
                "lr_index": lr_index,
                "broker_index": broker_index,
                "amount_index": amount_index,
                "transaction_count": transaction_count
            }

            if (
                best is None
                or candidate["score"]
                > best["score"]
            ):

                best = candidate

        return best

    def score_header_row(
        self,
        row
    ):

        normalized = [
            self.normalize_header(v)
            for v in row
        ]

        # ====================================================
        # IMPORTANT FIX
        #
        # Use the exact DATE column FIRST.
        #
        # In your sheet:
        #   BILL/DATE
        #   DATE
        #
        # The old code searched ["date", "billdate"]
        # and therefore selected BILL/DATE because it
        # appeared earlier.
        #
        # Now DATE always wins.
        # ====================================================

        date_index = self.find_header_index(
            normalized,
            [
                "date"
            ]
        )

        if date_index is None:

            date_index = self.find_header_index(
                normalized,
                [
                    "billdate"
                ]
            )

        lr_index = self.find_header_index(
            normalized,
            [
                "lrno",
                "lrnumber"
            ]
        )

        broker_index = self.find_header_index(
            normalized,
            [
                "broker"
            ]
        )

        amount_index = self.find_header_index(
            normalized,
            [
                "pay",
                "amount",
                "total",
                "totalamount",
                "nettotal",
                "amt",
                "invoiceamount",
                "billamount"
            ]
        )

        # July-style numeric header fallback.
        if amount_index is None:

            amount_index = (
                self.guess_numeric_amount_column(
                    row,
                    broker_index
                )
            )

        matches = sum(
            x is not None
            for x in [
                date_index,
                lr_index,
                broker_index,
                amount_index
            ]
        )

        if matches < 3:
            return None

        score = matches

        if amount_index is not None:
            score += 1

        return (
            score,
            date_index,
            lr_index,
            broker_index,
            amount_index
        )

    # ============================================================
    # HEADER HELPERS
    # ============================================================

    def normalize_header(
        self,
        value
    ):

        if value is None:
            return ""

        try:

            if pd.isna(value):
                return ""

        except Exception:
            pass

        text = str(
            value
        ).strip().lower()

        return re.sub(
            r"[^a-z0-9]+",
            "",
            text
        )

    def find_header_index(
        self,
        normalized,
        possible
    ):

        targets = [
            self.normalize_header(x)
            for x in possible
        ]

        for i, value in enumerate(
            normalized
        ):

            if value in targets:
                return i

        return None

    def guess_numeric_amount_column(
        self,
        row,
        broker_index
    ):

        candidates = []

        ignored = {
            "date",
            "billdate",
            "lrno",
            "broker",
            "advance",
            "rate",
            "state",
            "town",
            "truckno",
            "senunit",
            "no"
        }

        for index, value in enumerate(
            row
        ):

            if value is None:
                continue

            text = str(value).strip()

            if not text:
                continue

            normalized = (
                self.normalize_header(value)
            )

            if normalized in ignored:
                continue

            try:

                float(
                    text.replace(",", "")
                )

            except Exception:

                continue

            if broker_index is None:

                distance = 100

            else:

                distance = abs(
                    index - broker_index
                )

            before = (
                broker_index is not None
                and index < broker_index
            )

            score = 0

            if before:
                score += 100

            score -= distance

            candidates.append(
                (
                    score,
                    index
                )
            )

        if not candidates:
            return None

        candidates.sort(
            reverse=True
        )

        return candidates[0][1]

    def make_unique_headers(
        self,
        row
    ):

        headers = []
        counts = {}

        for index, value in enumerate(
            row
        ):

            if value is None:

                name = f"Unnamed_{index + 1}"

            else:

                text = str(
                    value
                ).strip()

                if not text:

                    name = f"Unnamed_{index + 1}"

                else:

                    name = text

            if name in counts:

                counts[name] += 1

                name = (
                    f"{name}_{counts[name]}"
                )

            else:

                counts[name] = 1

            headers.append(
                name
            )

        return headers

    # ============================================================
    # DATA CLEANING
    # ============================================================

    def clean_data(
        self,
        df
    ):

        df = df.copy()

        df = df.dropna(
            how="all"
        ).copy()

        for column in df.columns:

            if df[column].dtype == object:

                df[column] = df[column].apply(
                    lambda x:
                    x.strip()
                    if isinstance(x, str)
                    else x
                )

        return df.dropna(
            how="all"
        ).copy()

    # ============================================================
    # VALUE HELPERS
    # ============================================================

    def clean_text(
        self,
        value
    ):

        if value is None:
            return ""

        try:

            if pd.isna(value):
                return ""

        except Exception:
            pass

        return str(value).strip()

    # ============================================================
    # FIX:
    # Convert 15044.0 → 15044
    # ============================================================

    def format_lr(
        self,
        value
    ):

        if value is None:
            return ""

        try:

            if pd.isna(value):
                return ""

        except Exception:
            pass

        if isinstance(
            value,
            (int, float)
        ):

            try:

                number = float(
                    value
                )

                if math.isnan(number):
                    return ""

                if number.is_integer():

                    return str(
                        int(number)
                    )

                return str(
                    number
                )

            except Exception:

                return str(
                    value
                ).strip()

        text = str(
            value
        ).strip()

        if not text:
            return ""

        # Excel can turn an integer LR into
        # a string such as "15044.0".

        try:

            number = float(
                text
            )

            if number.is_integer():

                return str(
                    int(number)
                )

        except Exception:
            pass

        return text

    def parse_date(
        self,
        value
    ):

        if value is None:
            return None

        try:

            if pd.isna(value):
                return None

        except Exception:
            pass

        # Excel serial date.

        if isinstance(
            value,
            (int, float)
        ):

            try:

                number = float(
                    value
                )

                if (
                    not math.isnan(number)
                    and 20000
                    <= number
                    <= 100000
                ):

                    return (
                        pd.Timestamp(
                            "1899-12-30"
                        )
                        +
                        pd.to_timedelta(
                            number,
                            unit="D"
                        )
                    )

            except Exception:
                pass

        # Explicit date formats first.
        # This makes 02.03.26 unambiguous.

        if isinstance(
            value,
            str
        ):

            text = value.strip()

            formats = [
                "%d.%m.%y",
                "%d.%m.%Y",
                "%d-%m-%y",
                "%d-%m-%Y",
                "%d/%m/%y",
                "%d/%m/%Y"
            ]

            for fmt in formats:

                try:

                    return pd.Timestamp(
                        pd.to_datetime(
                            text,
                            format=fmt
                        )
                    )

                except Exception:
                    pass

        try:

            result = pd.to_datetime(
                value,
                dayfirst=True,
                errors="coerce"
            )

            if pd.isna(result):
                return None

            return result

        except Exception:

            return None

    def parse_amount(
        self,
        value
    ):

        if value is None:
            return None

        try:

            if pd.isna(value):
                return None

        except Exception:
            pass

        if isinstance(
            value,
            (int, float)
        ):

            try:

                number = float(
                    value
                )

                if math.isnan(number):
                    return None

                return number

            except Exception:

                return None

        text = str(
            value
        ).strip()

        if not text:
            return None

        text = (
            text
            .replace(",", "")
            .replace("₹", "")
            .replace("Rs.", "")
            .replace("Rs", "")
            .replace("INR", "")
            .strip()
        )

        if (
            text.startswith("(")
            and text.endswith(")")
        ):

            text = (
                "-"
                + text[1:-1]
            )

        try:

            return float(
                text
            )

        except Exception:

            return None

    # ============================================================
    # TRANSACTIONS
    # ============================================================

    def get_transaction_indices(
        self
    ):

        if self.data is None:
            return []

        date_col = self.date_menu.get()
        lr_col = self.lr_menu.get()
        broker_col = self.broker_menu.get()
        amount_col = self.amount_menu.get()

        required = [
            date_col,
            lr_col,
            broker_col,
            amount_col
        ]

        if not all(
            column in self.data.columns
            for column in required
        ):

            return []

        transactions = []

        for index in self.data.index:

            date_value = self.parse_date(
                self.data.loc[
                    index,
                    date_col
                ]
            )

            # FIX:
            # Use format_lr instead of clean_text.
            lr = self.format_lr(
                self.data.loc[
                    index,
                    lr_col
                ]
            )

            broker = self.clean_text(
                self.data.loc[
                    index,
                    broker_col
                ]
            )

            amount = self.parse_amount(
                self.data.loc[
                    index,
                    amount_col
                ]
            )

            if (
                date_value is not None
                and lr != ""
                and broker != ""
                and amount is not None
                and amount != 0
            ):

                transactions.append(
                    index
                )

        return transactions

    def count_transactions(
        self,
        df,
        date_col,
        lr_col,
        broker_col,
        amount_col
    ):

        if df.empty:
            return 0

        count = 0

        for index in df.index:

            date_value = self.parse_date(
                df.loc[
                    index,
                    date_col
                ]
            )

            # FIX:
            # Keep LR values clean here too.
            lr = self.format_lr(
                df.loc[
                    index,
                    lr_col
                ]
            )

            broker = self.clean_text(
                df.loc[
                    index,
                    broker_col
                ]
            )

            amount = self.parse_amount(
                df.loc[
                    index,
                    amount_col
                ]
            )

            if (
                date_value is not None
                and lr != ""
                and broker != ""
                and amount is not None
                and amount != 0
            ):

                count += 1

        return count

    # ============================================================
    # PREVIEW
    # ============================================================

    def preview_data(
        self
    ):

        if self.data is None:
            return

        date_col = self.date_menu.get()
        lr_col = self.lr_menu.get()
        broker_col = self.broker_menu.get()
        amount_col = self.amount_menu.get()

        transactions = (
            self.get_transaction_indices()
        )

        self.preview_box.delete(
            "1.0",
            "end"
        )

        self.preview_box.insert(
            "end",
            "DATA PREVIEW\n"
        )

        self.preview_box.insert(
            "end",
            "=" * 100
            + "\n\n"
        )

        self.preview_box.insert(
            "end",
            f"File: "
            f"{os.path.basename(self.excel_path)}\n"
        )

        self.preview_box.insert(
            "end",
            f"Sheet: "
            f"{self.current_sheet}\n"
        )

        self.preview_box.insert(
            "end",
            f"Header row: "
            f"{self.header_row + 1}\n"
        )

        self.preview_box.insert(
            "end",
            f"Transactions: "
            f"{len(transactions)}\n\n"
        )

        self.preview_box.insert(
            "end",
            "COLUMNS\n"
        )

        self.preview_box.insert(
            "end",
            "-" * 100
            + "\n"
        )

        self.preview_box.insert(
            "end",
            f"Date:       "
            f"{date_col}\n"
        )

        self.preview_box.insert(
            "end",
            f"L.R.NO:     "
            f"{lr_col}\n"
        )

        self.preview_box.insert(
            "end",
            f"Broker:     "
            f"{broker_col}\n"
        )

        self.preview_box.insert(
            "end",
            f"Amount:     "
            f"{amount_col}\n\n"
        )

        self.preview_box.insert(
            "end",
            "FIRST 10 TRANSACTIONS\n"
        )

        self.preview_box.insert(
            "end",
            "-" * 100
            + "\n"
        )

        for index in transactions[:10]:

            date_value = self.parse_date(
                self.data.loc[
                    index,
                    date_col
                ]
            )

            # FIX:
            # Preview now also shows 15044,
            # not 15044.0.
            lr = self.format_lr(
                self.data.loc[
                    index,
                    lr_col
                ]
            )

            broker = self.clean_text(
                self.data.loc[
                    index,
                    broker_col
                ]
            )

            amount = self.parse_amount(
                self.data.loc[
                    index,
                    amount_col
                ]
            )

            self.preview_box.insert(
                "end",
                f"{date_value.strftime('%d-%m-%Y'):<12} | "
                f"LR: {lr:<12} | "
                f"{broker:<35} | "
                f"₹{amount:,.2f}\n"
            )

        self._update_chips()

    # ============================================================
    # VALIDATION
    # ============================================================

    def validate_data(
        self
    ):

        if self.data is None:

            messagebox.showwarning(
                "No Sheet",
                "Select an Excel sheet first."
            )

            return

        transactions = (
            self.get_transaction_indices()
        )

        self.data_valid = (
            len(transactions) > 0
        )

        self.preview_box.insert(
            "end",
            "\n\nVALIDATION\n"
        )

        self.preview_box.insert(
            "end",
            "-" * 100
            + "\n"
        )

        self.preview_box.insert(
            "end",
            f"Complete transactions: "
            f"{len(transactions)}\n"
        )

        if self.data_valid:

            self.preview_box.insert(
                "end",
                "✓ DATA LOOKS GOOD"
            )

            self.status.configure(
                text=(
                    f"✓ {len(transactions)} "
                    f"transaction(s) ready"
                )
            )

        else:

            self.preview_box.insert(
                "end",
                "⚠ NO COMPLETE TRANSACTIONS"
            )

            self.status.configure(
                text="⚠ No usable transactions"
            )

        if self.mapping_data is not None:

            self.mapping_button.configure(
                state="normal"
            )

        self.update_generate_state()

    # ============================================================
    # MAPPING FILE
    # ============================================================

    def select_mapping(
        self
    ):

        path = filedialog.askopenfilename(
            title="Select Broker Mapping File",
            filetypes=[
                ("Excel Files", "*.xlsx *.xls"),
                ("All Files", "*.*")
            ]
        )

        if not path:
            return

        try:

            mapping = pd.read_excel(
                path,
                dtype=object
            )

            if mapping.empty:

                raise ValueError(
                    "Mapping file is empty."
                )

            if len(mapping.columns) < 2:

                raise ValueError(
                    "Mapping file needs at least "
                    "two columns."
                )

            columns = [
                str(x).strip()
                for x in mapping.columns
            ]

            mapping.columns = columns

            excel_col = None
            tally_col = None

            for column in columns:

                normalized = (
                    self.normalize_header(
                        column
                    )
                )

                if (
                    "excel" in normalized
                    and "broker" in normalized
                ):

                    excel_col = column

                if (
                    "tally" in normalized
                    and "ledger" in normalized
                ):

                    tally_col = column

            if excel_col is None:

                excel_col = columns[0]

            if tally_col is None:

                tally_col = columns[1]

            mapping = mapping[
                [
                    excel_col,
                    tally_col
                ]
            ].copy()

            mapping.columns = [
                "Excel Broker",
                "Tally Ledger"
            ]

            mapping["Excel Broker"] = (
                mapping["Excel Broker"]
                .apply(
                    self.clean_text
                )
            )

            mapping["Tally Ledger"] = (
                mapping["Tally Ledger"]
                .apply(
                    self.clean_text
                )
            )

            mapping = mapping[
                (
                    mapping["Excel Broker"]
                    != ""
                )
                &
                (
                    mapping["Tally Ledger"]
                    != ""
                )
            ].copy()

            self.mapping_data = mapping
            self.mapping_path = path

            self.mapping_dict = {}

            for _, row in mapping.iterrows():

                self.mapping_dict[
                    row["Excel Broker"]
                    .casefold()
                ] = row[
                    "Tally Ledger"
                ]

            self.mapping_valid = False

            self.mapping_label.configure(
                text=os.path.basename(path),
                text_color=self.T_PRIMARY,
            )

            self.mapping_box.delete(
                "1.0",
                "end"
            )

            self.mapping_box.insert(
                "end",
                "MAPPING FILE LOADED\n"
            )

            self.mapping_box.insert(
                "end",
                "=" * 100
                + "\n\n"
            )

            self.mapping_box.insert(
                "end",
                f"Mappings: "
                f"{len(mapping)}\n\n"
            )

            self.mapping_box.insert(
                "end",
                f"Excel Broker → "
                f"{excel_col}\n"
            )

            self.mapping_box.insert(
                "end",
                f"Tally Ledger → "
                f"{tally_col}\n\n"
            )

            self.mapping_box.insert(
                "end",
                "✓ Ready for mapping check."
            )

            if self.data is not None:

                self.mapping_button.configure(
                    state="normal"
                )

            self.update_generate_state()
            self._update_chips()

        except Exception as e:

            messagebox.showerror(
                "Mapping Error",
                f"Could not load mapping file:\n\n{e}"
            )

    # ============================================================
    # CHECK MAPPING
    # ============================================================

    def check_mapping(
        self
    ):

        if self.data is None:

            messagebox.showwarning(
                "No Sheet",
                "Select an Excel sheet first."
            )

            return

        if self.mapping_data is None:

            messagebox.showwarning(
                "No Mapping",
                "Select the broker mapping file."
            )

            return

        transactions = (
            self.get_transaction_indices()
        )

        if not transactions:

            messagebox.showwarning(
                "No Transactions",
                "No complete transactions were found."
            )

            return

        broker_col = (
            self.broker_menu.get()
        )

        matched = []
        unmatched = []
        counts = {}

        for index in transactions:

            broker = self.clean_text(
                self.data.loc[
                    index,
                    broker_col
                ]
            )

            key = broker.casefold()

            if key not in counts:

                counts[key] = {
                    "name": broker,
                    "count": 0
                }

            counts[key]["count"] += 1

        for key, info in counts.items():

            broker = info["name"]
            count = info["count"]

            if key in self.mapping_dict:

                matched.append(
                    (
                        broker,
                        self.mapping_dict[key],
                        count
                    )
                )

            else:

                unmatched.append(
                    (
                        broker,
                        count
                    )
                )

        matched.sort(
            key=lambda x:
            x[0].casefold()
        )

        unmatched.sort(
            key=lambda x:
            x[0].casefold()
        )

        self.mapping_box.delete(
            "1.0",
            "end"
        )

        self.mapping_box.insert(
            "end",
            "BROKER MAPPING RESULT\n"
        )

        self.mapping_box.insert(
            "end",
            "=" * 100
            + "\n\n"
        )

        self.mapping_box.insert(
            "end",
            f"Transactions checked: "
            f"{len(transactions)}\n"
        )

        self.mapping_box.insert(
            "end",
            f"Unique brokers:       "
            f"{len(counts)}\n"
        )

        self.mapping_box.insert(
            "end",
            f"Matched:               "
            f"{len(matched)}\n"
        )

        self.mapping_box.insert(
            "end",
            f"Unmatched:             "
            f"{len(unmatched)}\n\n"
        )

        if matched:

            self.mapping_box.insert(
                "end",
                "MATCHED\n"
            )

            self.mapping_box.insert(
                "end",
                "-" * 100
                + "\n"
            )

            for (
                broker,
                ledger,
                count
            ) in matched:

                self.mapping_box.insert(
                    "end",
                    f"✓ {broker} → {ledger} "
                    f"({count} row(s))\n"
                )

        if unmatched:

            self.mapping_box.insert(
                "end",
                "\nUNMATCHED\n"
            )

            self.mapping_box.insert(
                "end",
                "-" * 100
                + "\n"
            )

            for (
                broker,
                count
            ) in unmatched:

                self.mapping_box.insert(
                    "end",
                    f"⚠ {broker} "
                    f"({count} row(s))\n"
                )

            self.mapping_valid = False

            self.status.configure(
                text=(
                    f"⚠ {len(unmatched)} "
                    f"unmapped broker(s)"
                )
            )

        else:

            self.mapping_box.insert(
                "end",
                "\n✓ ALL BROKERS ARE MAPPED"
            )

            self.mapping_valid = True

            self.status.configure(
                text="✓ Broker mapping successful"
            )

        self.update_generate_state()

    # ============================================================
    # GENERATE BUTTON STATE
    # ============================================================

    def update_generate_state(
        self
    ):

        if (
            self.data_valid
            and self.mapping_valid
        ):

            self.generate_button.configure(
                state="normal"
            )

        else:

            self.generate_button.configure(
                state="disabled"
            )

    # ============================================================
    # BUILD XML
    # ============================================================

    def build_xml(
        self,
        transaction_indices
    ):

        company = (
            self.company_entry
            .get()
            .strip()
        )

        debit_ledger = (
            self.ledger_entry
            .get()
            .strip()
        )

        date_col = (
            self.date_menu.get()
        )

        lr_col = (
            self.lr_menu.get()
        )

        broker_col = (
            self.broker_menu.get()
        )

        amount_col = (
            self.amount_menu.get()
        )

        # ----------------------------------------------------
        # ROOT
        # ----------------------------------------------------

        envelope = ET.Element(
            "ENVELOPE"
        )

        header = ET.SubElement(
            envelope,
            "HEADER"
        )

        ET.SubElement(
            header,
            "TALLYREQUEST"
        ).text = "Import Data"

        body = ET.SubElement(
            envelope,
            "BODY"
        )

        import_data = ET.SubElement(
            body,
            "IMPORTDATA"
        )

        request_desc = ET.SubElement(
            import_data,
            "REQUESTDESC"
        )

        ET.SubElement(
            request_desc,
            "REPORTNAME"
        ).text = "Vouchers"

        static_variables = ET.SubElement(
            request_desc,
            "STATICVARIABLES"
        )

        ET.SubElement(
            static_variables,
            "SVCURRENTCOMPANY"
        ).text = company

        request_data = ET.SubElement(
            import_data,
            "REQUESTDATA"
        )

        # ----------------------------------------------------
        # VOUCHERS
        # ----------------------------------------------------

        created = 0
        skipped = []

        for index in transaction_indices:

            date_value = self.parse_date(
                self.data.loc[
                    index,
                    date_col
                ]
            )

            # =================================================
            # IMPORTANT FIX:
            #
            # Previously:
            #     clean_text(...)
            #
            # That produced:
            #     15044.0
            #
            # Now:
            #     format_lr(...)
            #
            # produces:
            #     15044
            # =================================================

            lr_value = self.format_lr(
                self.data.loc[
                    index,
                    lr_col
                ]
            )

            broker = self.clean_text(
                self.data.loc[
                    index,
                    broker_col
                ]
            )

            amount = self.parse_amount(
                self.data.loc[
                    index,
                    amount_col
                ]
            )

            if (
                date_value is None
                or not lr_value
                or not broker
                or amount is None
                or amount == 0
            ):

                skipped.append(
                    (
                        index,
                        "Invalid transaction data"
                    )
                )

                continue

            key = broker.casefold()

            if key not in self.mapping_dict:

                skipped.append(
                    (
                        index,
                        f"No mapping for '{broker}'"
                    )
                )

                continue

            tally_ledger = (
                self.mapping_dict[key]
            )

            guid = str(
                uuid.uuid4()
            )

            date_string = (
                date_value.strftime(
                    "%Y%m%d"
                )
            )

            amount_string = (
                f"{abs(amount):.2f}"
            )

            message = ET.SubElement(
                request_data,
                "TALLYMESSAGE"
            )

            message.set(
                "xmlns:UDF",
                "TallyUDF"
            )

            voucher = ET.SubElement(
                message,
                "VOUCHER"
            )

            voucher.set(
                "REMOTEID",
                guid
            )

            voucher.set(
                "VCHTYPE",
                "Journal"
            )

            voucher.set(
                "ACTION",
                "Create"
            )

            ET.SubElement(
                voucher,
                "DATE"
            ).text = date_string

            ET.SubElement(
                voucher,
                "GUID"
            ).text = guid

            ET.SubElement(
                voucher,
                "NARRATION"
            ).text = (
                f"Lr no. {lr_value}"
            )

            ET.SubElement(
                voucher,
                "VOUCHERTYPENAME"
            ).text = "Journal"

            ET.SubElement(
                voucher,
                "PARTYLEDGERNAME"
            ).text = tally_ledger

            # ------------------------------------------------
            # STANDARD TALLY FIELDS
            # ------------------------------------------------

            standard_fields = {
                "CSTFORMISSUETYPE": "",
                "CSTFORMRECVTYPE": "",
                "FBTPAYMENTTYPE": "Default",
                "VCHGSTCLASS": "",
                "PERSISTEDVIEW": "Accounting Voucher View",
                "DIFFACTUALQTY": "No",
                "AUDITED": "No",
                "FORJOBCOSTING": "No",
                "ISOPTIONAL": "No",
                "EFFECTIVEDATE": date_string,
                "USEFORINTEREST": "No",
                "USEFORGAINLOSS": "No",
                "USEFORGODOWNTRANSFER": "No",
                "USEFORCOMPOUND": "No",
                "EXCISEOPENING": "No",
                "ISCANCELLED": "No",
                "HASCASHFLOW": "No",
                "ISPOSTDATED": "No",
                "USETRACKINGNUMBER": "No",
                "ISINVOICE": "No",
                "MFGJOURNAL": "No",
                "HASDISCOUNTS": "No",
                "ASPAYSLIP": "No",
                "ISDELETED": "No",
                "ASORIGINAL": "No"
            }

            for tag, value in (
                standard_fields.items()
            ):

                ET.SubElement(
                    voucher,
                    tag
                ).text = value

            # ------------------------------------------------
            # DEBIT — TRANSPORT EXPENSE
            # ------------------------------------------------

            debit_entry = ET.SubElement(
                voucher,
                "ALLLEDGERENTRIES.LIST"
            )

            ET.SubElement(
                debit_entry,
                "LEDGERNAME"
            ).text = debit_ledger

            ET.SubElement(
                debit_entry,
                "GSTCLASS"
            )

            ET.SubElement(
                debit_entry,
                "ISDEEMEDPOSITIVE"
            ).text = "Yes"

            ET.SubElement(
                debit_entry,
                "LEDGERFROMITEM"
            ).text = "No"

            ET.SubElement(
                debit_entry,
                "REMOVEZEROENTRIES"
            ).text = "No"

            ET.SubElement(
                debit_entry,
                "ISPARTYLEDGER"
            ).text = "No"

            ET.SubElement(
                debit_entry,
                "AMOUNT"
            ).text = (
                f"-{amount_string}"
            )

            # ------------------------------------------------
            # CREDIT — BROKER
            # ------------------------------------------------

            credit_entry = ET.SubElement(
                voucher,
                "ALLLEDGERENTRIES.LIST"
            )

            ET.SubElement(
                credit_entry,
                "LEDGERNAME"
            ).text = tally_ledger

            ET.SubElement(
                credit_entry,
                "GSTCLASS"
            )

            ET.SubElement(
                credit_entry,
                "ISDEEMEDPOSITIVE"
            ).text = "No"

            ET.SubElement(
                credit_entry,
                "LEDGERFROMITEM"
            ).text = "No"

            ET.SubElement(
                credit_entry,
                "REMOVEZEROENTRIES"
            ).text = "No"

            ET.SubElement(
                credit_entry,
                "ISPARTYLEDGER"
            ).text = "Yes"

            ET.SubElement(
                credit_entry,
                "AMOUNT"
            ).text = amount_string

            created += 1

        return (
            envelope,
            created,
            skipped
        )

    # ============================================================
    # GENERATE XML
    # ============================================================

    def generate_xml(
        self
    ):

        if not self.data_valid:

            messagebox.showwarning(
                "Data Not Valid",
                "Validate the Excel data first."
            )

            return

        if not self.mapping_valid:

            messagebox.showwarning(
                "Mapping Not Valid",
                "Check broker mapping first."
            )

            return

        transactions = (
            self.get_transaction_indices()
        )

        if not transactions:

            messagebox.showwarning(
                "No Transactions",
                "No transactions available."
            )

            return

        # Test mode = first voucher only.

        if self.test_first_var.get():

            selected_indices = (
                transactions[:1]
            )

        else:

            selected_indices = transactions

        # Build XML.

        try:

            (
                envelope,
                created,
                skipped
            ) = self.build_xml(
                selected_indices
            )

            if created == 0:

                messagebox.showerror(
                    "Generation Failed",
                    "No valid vouchers could be created."
                )

                return

            # Pretty XML.

            rough_xml = ET.tostring(
                envelope,
                encoding="utf-8"
            )

            parsed = minidom.parseString(
                rough_xml
            )

            pretty_xml = (
                parsed.toprettyxml(
                    indent="  ",
                    encoding="UTF-8"
                )
                .decode("utf-8")
            )

            # Save.

            default_name = (
                "Tally_Test.xml"
                if self.test_first_var.get()
                else "Tally_Import.xml"
            )

            path = filedialog.asksaveasfilename(
                title="Save Tally XML",
                defaultextension=".xml",
                initialfile=default_name,
                filetypes=[
                    ("XML Files", "*.xml"),
                    ("All Files", "*.*")
                ]
            )

            if not path:
                return

            with open(
                path,
                "w",
                encoding="utf-8"
            ) as file:

                file.write(
                    pretty_xml
                )

            # Result.

            message = (
                f"XML created successfully.\n\n"
                f"Vouchers created: {created}\n"
                f"Requested: {len(selected_indices)}\n"
                f"Skipped: {len(skipped)}\n\n"
                f"File:\n{path}"
            )

            if skipped:

                message += (
                    "\n\n"
                    "Skipped rows:\n"
                )

                for index, reason in skipped[:10]:

                    message += (
                        f"\nRow {index + 1}: "
                        f"{reason}"
                    )

                if len(skipped) > 10:

                    message += (
                        f"\n\n...and "
                        f"{len(skipped) - 10} more."
                    )

            messagebox.showinfo(
                "XML Generated",
                message
            )

            # Show generated voucher info.

            self.preview_box.delete(
                "1.0",
                "end"
            )

            self.preview_box.insert(
                "end",
                "XML GENERATION RESULT\n"
            )

            self.preview_box.insert(
                "end",
                "=" * 100
                + "\n\n"
            )

            self.preview_box.insert(
                "end",
                f"Vouchers created: "
                f"{created}\n"
            )

            self.preview_box.insert(
                "end",
                f"Skipped: "
                f"{len(skipped)}\n"
            )

            self.preview_box.insert(
                "end",
                f"Saved to:\n"
                f"{path}\n\n"
            )

            self.preview_box.insert(
                "end",
                "The XML is ready to import into Tally."
            )

            self.status.configure(
                text=(
                    f"✓ XML created — "
                    f"{created} voucher(s)"
                )
            )

        except Exception as e:

            messagebox.showerror(
                "XML Error",
                f"Could not generate XML:\n\n{e}"
            )

    # ============================================================
    # CLEAR MENUS
    # ============================================================

    def clear_column_menus(
        self
    ):

        for menu in [
            self.date_menu,
            self.lr_menu,
            self.broker_menu,
            self.amount_menu
        ]:

            menu.configure(
                values=["No columns detected"]
            )

            menu.set(
                "No columns detected"
            )


# ============================================================
# START
# ============================================================

if __name__ == "__main__":

    app = TallyConverterApp()

    app.mainloop()
