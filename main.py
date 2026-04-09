"""
Email Automation — main GUI
Run with: python main.py
Requires Python 3.10+ (standard library only).
"""

import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, simpledialog, ttk

import email_sender as es


# ---------------------------------------------------------------------------
# App
# ---------------------------------------------------------------------------

class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title('Email Automation')
        self.root.geometry('980x720')
        self.root.minsize(800, 580)

        # Shared recipient state (populated by Import CSV or Builder)
        self.config: dict = es.load_config()
        self.recipients: list[dict] = []
        self.headers: list[str] = []
        self._stop_flag = False
        self._sending = False

        # Link href storage for the rich-text editor  {tag_name -> url}
        self._link_hrefs: dict[str, str] = {}
        self._link_counter = 0

        self._apply_style()
        self._build_ui()
        self._populate_smtp_fields()

    # ------------------------------------------------------------------
    # Style
    # ------------------------------------------------------------------

    def _apply_style(self):
        style = ttk.Style(self.root)
        available = style.theme_names()
        for preferred in ('vista', 'winnative', 'clam', 'alt', 'default'):
            if preferred in available:
                style.theme_use(preferred)
                break

        style.configure('TNotebook.Tab', padding=(14, 6), font=('Segoe UI', 10))
        style.configure('TLabel', font=('Segoe UI', 10))
        style.configure('TButton', font=('Segoe UI', 10), padding=(8, 4))
        style.configure('TEntry', font=('Segoe UI', 10))
        style.configure('Header.TLabel', font=('Segoe UI', 11, 'bold'))
        style.configure('Hint.TLabel', font=('Segoe UI', 9), foreground='#666666')
        style.configure('Success.TLabel', foreground='#2e7d32')
        style.configure('Error.TLabel', foreground='#c62828')

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self):
        self.nb = ttk.Notebook(self.root)
        self.nb.pack(fill='both', expand=True, padx=12, pady=(10, 12))

        self._build_smtp_tab()
        self._build_compose_tab()
        self._build_recipients_tab()
        self._build_builder_tab()
        self._build_send_tab()

    # ---- Tab 1: SMTP Setup -------------------------------------------

    PROVIDERS = {
        'Gmail':   {'server': 'smtp.gmail.com',        'port': '587', 'use_ssl': False},
        'Outlook': {'server': 'smtp-mail.outlook.com', 'port': '587', 'use_ssl': False},
        'Yahoo':   {'server': 'smtp.mail.yahoo.com',   'port': '587', 'use_ssl': False},
        'Custom':  {'server': '',                       'port': '587', 'use_ssl': False},
    }

    PROVIDER_HINTS = {
        'Gmail':   ('Requires an App Password — your regular password will not work.\n'
                    'Go to myaccount.google.com → Security → App passwords to create one.'),
        'Outlook': 'Use your regular Outlook / Hotmail password.',
        'Yahoo':   ('Requires an App Password.\n'
                    'Go to Yahoo Account Security → Generate app password.'),
        'Custom':  "Enter your provider's SMTP server and port manually.",
    }

    def _build_smtp_tab(self):
        outer = ttk.Frame(self.nb)
        self.nb.add(outer, text='  SMTP Setup  ')

        frame = ttk.Frame(outer, padding=24)
        frame.pack(fill='both', expand=True)
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text='Email Account Setup', style='Header.TLabel').grid(
            row=0, column=0, columnspan=3, sticky='w', pady=(0, 18))

        # Provider dropdown
        ttk.Label(frame, text='Email Provider:').grid(row=1, column=0, sticky='w', pady=5)
        self._provider_var = tk.StringVar(value='Gmail')
        provider_combo = ttk.Combobox(
            frame, textvariable=self._provider_var,
            values=list(self.PROVIDERS.keys()),
            state='readonly', width=20,
        )
        provider_combo.grid(row=1, column=1, sticky='w', padx=(12, 8), pady=5)
        provider_combo.bind('<<ComboboxSelected>>', self._on_provider_change)

        self._provider_hint = ttk.Label(
            frame, text=self.PROVIDER_HINTS['Gmail'],
            style='Hint.TLabel', wraplength=480, justify='left',
        )
        self._provider_hint.grid(row=1, column=2, sticky='w', pady=5)

        # Fields
        fields = [
            ('Username', 'username', False, 'Your full email address (e.g. you@gmail.com)'),
            ('Password', 'password', True,  'See the hint next to your provider above'),
            ('From Name', 'from_name', False, 'Display name shown to recipients (optional)'),
            ('Delay between sends (ms)', 'delay_ms', False,
             'Pause between emails to avoid rate-limits (default 300)'),
        ]

        self._smtp_vars: dict[str, tk.StringVar] = {}
        for key in ('server', 'port'):
            self._smtp_vars[key] = tk.StringVar()

        for i, (label, key, secret, hint) in enumerate(fields):
            row = i + 2
            ttk.Label(frame, text=label + ':').grid(row=row, column=0, sticky='w', pady=5)
            var = tk.StringVar()
            self._smtp_vars[key] = var
            ttk.Entry(frame, textvariable=var, show='•' if secret else '', width=42).grid(
                row=row, column=1, sticky='ew', padx=(12, 8), pady=5)
            ttk.Label(frame, text=hint, style='Hint.TLabel').grid(
                row=row, column=2, sticky='w', pady=5)

        # Custom server/port row (hidden unless Custom selected)
        self._custom_frame = ttk.Frame(frame)
        self._custom_frame.grid(row=len(fields) + 2, column=0, columnspan=3, sticky='ew', pady=4)
        ttk.Label(self._custom_frame, text='SMTP Server:').grid(row=0, column=0, sticky='w', padx=(0, 6))
        ttk.Entry(self._custom_frame, textvariable=self._smtp_vars['server'], width=30).grid(
            row=0, column=1, sticky='w', padx=(0, 20))
        ttk.Label(self._custom_frame, text='Port:').grid(row=0, column=2, sticky='w', padx=(0, 6))
        ttk.Entry(self._custom_frame, textvariable=self._smtp_vars['port'], width=6).grid(
            row=0, column=3, sticky='w', padx=(0, 20))
        self._use_ssl = tk.BooleanVar(value=False)
        ttk.Checkbutton(self._custom_frame, text='Use SSL', variable=self._use_ssl).grid(
            row=0, column=4, sticky='w')
        self._custom_frame.grid_remove()

        btn_row = len(fields) + 3
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=btn_row, column=0, columnspan=3, sticky='w', pady=(16, 0))
        ttk.Button(btn_frame, text='Save Settings', command=self._save_smtp).pack(side='left', padx=(0, 10))
        ttk.Button(btn_frame, text='Test Connection', command=self._test_connection).pack(side='left')

        self._smtp_status = ttk.Label(frame, text='')
        self._smtp_status.grid(row=btn_row + 1, column=0, columnspan=3, sticky='w', pady=(10, 0))

    def _on_provider_change(self, *_):
        provider = self._provider_var.get()
        s = self.PROVIDERS[provider]
        self._smtp_vars['server'].set(s['server'])
        self._smtp_vars['port'].set(s['port'])
        self._use_ssl.set(s['use_ssl'])
        self._provider_hint.config(text=self.PROVIDER_HINTS[provider])
        if provider == 'Custom':
            self._custom_frame.grid()
        else:
            self._custom_frame.grid_remove()

    # ---- Tab 2: Compose ----------------------------------------------

    def _build_compose_tab(self):
        outer = ttk.Frame(self.nb)
        self.nb.add(outer, text='  Compose  ')

        frame = ttk.Frame(outer, padding=20)
        frame.pack(fill='both', expand=True)
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(3, weight=1)

        ttk.Label(frame, text='Compose Email', style='Header.TLabel').grid(
            row=0, column=0, columnspan=2, sticky='w', pady=(0, 10))

        # Variable hint box
        hint_box = ttk.LabelFrame(frame, text='Template Variables', padding=(12, 6))
        hint_box.grid(row=1, column=0, columnspan=2, sticky='ew', pady=(0, 12))
        hint_box.columnconfigure(0, weight=1)
        ttk.Label(
            hint_box,
            text='Type {{variable_name}} anywhere in the subject or body. '
                 'The name must match a column header in your CSV (case-sensitive).',
            style='Hint.TLabel', wraplength=820, justify='left',
        ).grid(row=0, column=0, sticky='w')
        self._detected_vars_label = ttk.Label(
            hint_box, text='Detected variables: —',
            foreground='#1565c0', font=('Segoe UI', 9, 'bold'),
        )
        self._detected_vars_label.grid(row=1, column=0, sticky='w', pady=(4, 0))

        # Subject
        ttk.Label(frame, text='Subject:').grid(row=2, column=0, sticky='w', pady=(0, 6))
        self._subject_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self._subject_var, font=('Segoe UI', 10)).grid(
            row=2, column=1, sticky='ew', padx=(12, 0), pady=(0, 6))
        self._subject_var.trace_add('write', self._refresh_vars)

        # Body label
        ttk.Label(frame, text='Body:').grid(row=3, column=0, sticky='nw', pady=(4, 0))

        # Body container (toolbar + text area)
        body_frame = ttk.Frame(frame)
        body_frame.grid(row=3, column=1, sticky='nsew', padx=(12, 0))
        body_frame.columnconfigure(0, weight=1)
        body_frame.rowconfigure(1, weight=1)

        # Formatting toolbar
        toolbar = tk.Frame(body_frame, bg='#f0f0f0', relief='flat', bd=1)
        toolbar.grid(row=0, column=0, sticky='ew')

        def fmt_btn(text, cmd, tooltip=''):
            b = tk.Button(toolbar, text=text, command=cmd,
                          font=('Segoe UI', 10, 'bold'), relief='flat',
                          padx=8, pady=2, bg='#f0f0f0', activebackground='#d0d0d0',
                          cursor='hand2', bd=0)
            b.pack(side='left', padx=1, pady=2)
            return b

        fmt_btn('B',  lambda: self._toggle_format('bold'))
        fmt_btn('I',  lambda: self._toggle_format('italic'))
        fmt_btn('U',  lambda: self._toggle_format('underline'))

        # Separator
        tk.Label(toolbar, text='|', bg='#f0f0f0', fg='#aaa').pack(side='left', padx=4)

        fmt_btn('🔗 Link', self._insert_link)

        # Separator
        tk.Label(toolbar, text='|', bg='#f0f0f0', fg='#aaa').pack(side='left', padx=4)
        fmt_btn('✕ Clear Format', self._clear_format)

        # Body text
        self._body_text = tk.Text(
            body_frame, wrap='word', font=('Segoe UI', 10),
            relief='solid', borderwidth=1, padx=6, pady=6, undo=True,
        )
        vsb = ttk.Scrollbar(body_frame, orient='vertical', command=self._body_text.yview)
        self._body_text.configure(yscrollcommand=vsb.set)
        self._body_text.grid(row=1, column=0, sticky='nsew')
        vsb.grid(row=1, column=1, sticky='ns')

        # Configure formatting tags
        self._body_text.tag_configure('bold',      font=('Segoe UI', 10, 'bold'))
        self._body_text.tag_configure('italic',    font=('Segoe UI', 10, 'italic'))
        self._body_text.tag_configure('underline', underline=True)
        self._body_text.tag_configure('bold_italic',
                                      font=('Segoe UI', 10, 'bold italic'))
        self._body_text.tag_configure('bold_underline',
                                      font=('Segoe UI', 10, 'bold'), underline=True)
        self._body_text.tag_configure('italic_underline',
                                      font=('Segoe UI', 10, 'italic'), underline=True)
        self._body_text.tag_configure('bold_italic_underline',
                                      font=('Segoe UI', 10, 'bold italic'), underline=True)

        self._body_text.bind('<KeyRelease>', self._refresh_vars)

        # Keyboard shortcuts
        self._body_text.bind('<Control-b>', lambda e: (self._toggle_format('bold'), 'break'))
        self._body_text.bind('<Control-i>', lambda e: (self._toggle_format('italic'), 'break'))
        self._body_text.bind('<Control-u>', lambda e: (self._toggle_format('underline'), 'break'))

        # Right-click context menu
        self._body_text.bind('<Button-3>', lambda e: self._show_context_menu(e, self._body_text))

    # ---- Tab 3: Recipients (Import CSV) --------------------------------

    def _build_recipients_tab(self):
        outer = ttk.Frame(self.nb)
        self.nb.add(outer, text='  Import CSV  ')

        frame = ttk.Frame(outer, padding=24)
        frame.pack(fill='both', expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(2, weight=1)

        ttk.Label(frame, text='Import Recipients from CSV', style='Header.TLabel').grid(
            row=0, column=0, sticky='w', pady=(0, 12))

        # Toolbar
        toolbar = ttk.Frame(frame)
        toolbar.grid(row=1, column=0, sticky='ew', pady=(0, 10))

        ttk.Button(toolbar, text='Import CSV…', command=self._import_csv).pack(side='left', padx=(0, 8))
        ttk.Button(toolbar, text='Show CSV Format Example', command=self._show_csv_example).pack(
            side='left', padx=(0, 20))

        ttk.Label(toolbar, text='Email column:').pack(side='left')
        self._email_col = tk.StringVar(value='email')
        self._email_col_combo = ttk.Combobox(
            toolbar, textvariable=self._email_col, width=16, state='readonly')
        self._email_col_combo.pack(side='left', padx=(6, 0))

        self._recipient_count_label = ttk.Label(toolbar, text='No CSV loaded.', style='Hint.TLabel')
        self._recipient_count_label.pack(side='right')

        # Treeview
        tree_frame = ttk.Frame(frame)
        tree_frame.grid(row=2, column=0, sticky='nsew')
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)

        self._tree = ttk.Treeview(tree_frame, show='headings', selectmode='browse')
        vsb = ttk.Scrollbar(tree_frame, orient='vertical', command=self._tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient='horizontal', command=self._tree.xview)
        self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self._tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')

        ttk.Label(
            frame,
            text='Tip: first row = column headers. The header names become your {{variable}} names. '
                 'Click "Show CSV Format Example" to see a sample.',
            style='Hint.TLabel', wraplength=880, justify='left',
        ).grid(row=3, column=0, sticky='w', pady=(10, 0))

    # ---- Tab 4: Build Recipients (in-app database) --------------------

    def _build_builder_tab(self):
        outer = ttk.Frame(self.nb)
        self.nb.add(outer, text='  Build Recipients  ')

        frame = ttk.Frame(outer, padding=20)
        frame.pack(fill='both', expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)

        ttk.Label(frame, text='Build Recipient List', style='Header.TLabel').grid(
            row=0, column=0, sticky='w', pady=(0, 10))

        # State for the builder
        self._builder_cols: list[str] = ['email']
        self._builder_rows: list[list[str]] = []

        # Toolbar
        tb = ttk.Frame(frame)
        tb.grid(row=0, column=0, sticky='e')

        ttk.Button(tb, text='+ Add Column', command=self._builder_add_column).pack(side='left', padx=(0, 6))
        ttk.Button(tb, text='- Remove Column', command=self._builder_remove_column).pack(side='left', padx=(0, 6))
        ttk.Button(tb, text='+ Add Row', command=self._builder_add_row).pack(side='left', padx=(0, 6))
        ttk.Button(tb, text='- Delete Row', command=self._builder_delete_row).pack(side='left', padx=(0, 20))
        ttk.Button(tb, text='Save as CSV…', command=self._builder_save_csv).pack(side='left', padx=(0, 6))
        ttk.Button(tb, text='Use for Sending', command=self._builder_use).pack(side='left')

        # Editable treeview
        tree_frame = ttk.Frame(frame)
        tree_frame.grid(row=1, column=0, sticky='nsew', pady=(10, 0))
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)

        self._btree = ttk.Treeview(tree_frame, show='headings', selectmode='browse')
        bvsb = ttk.Scrollbar(tree_frame, orient='vertical', command=self._btree.yview)
        bhsb = ttk.Scrollbar(tree_frame, orient='horizontal', command=self._btree.xview)
        self._btree.configure(yscrollcommand=bvsb.set, xscrollcommand=bhsb.set)
        self._btree.grid(row=0, column=0, sticky='nsew')
        bvsb.grid(row=0, column=1, sticky='ns')
        bhsb.grid(row=1, column=0, sticky='ew')

        self._btree.bind('<Double-1>', self._builder_edit_cell)

        self._builder_status = ttk.Label(frame, text='', style='Hint.TLabel')
        self._builder_status.grid(row=2, column=0, sticky='w', pady=(6, 0))

        ttk.Label(
            frame,
            text='Double-click any cell to edit it.  '
                 '"Use for Sending" copies this list to the Send tab.  '
                 'The first column should be the email address.',
            style='Hint.TLabel', wraplength=880, justify='left',
        ).grid(row=3, column=0, sticky='w', pady=(4, 0))

        self._builder_rebuild_tree()

    # ---- Tab 5: Preview & Send ----------------------------------------

    def _build_send_tab(self):
        outer = ttk.Frame(self.nb)
        self.nb.add(outer, text='  Preview & Send  ')

        frame = ttk.Frame(outer, padding=24)
        frame.pack(fill='both', expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)

        ttk.Label(frame, text='Preview & Send', style='Header.TLabel').grid(
            row=0, column=0, sticky='w', pady=(0, 12))

        # Preview panel
        preview_panel = ttk.LabelFrame(frame, text='Email Preview', padding=(12, 10))
        preview_panel.grid(row=1, column=0, sticky='nsew', pady=(0, 12))
        preview_panel.columnconfigure(1, weight=1)
        preview_panel.rowconfigure(2, weight=1)

        ttk.Label(preview_panel, text='Preview for:').grid(row=0, column=0, sticky='w')
        self._preview_combo = ttk.Combobox(preview_panel, state='readonly')
        self._preview_combo.grid(row=0, column=1, sticky='ew', padx=(10, 10))
        self._preview_combo.bind('<<ComboboxSelected>>', self._update_preview)

        ttk.Button(preview_panel, text='Refresh Preview', command=self._update_preview).grid(
            row=0, column=2, sticky='e')

        self._preview_subject_label = ttk.Label(
            preview_panel, text='Subject: —', font=('Segoe UI', 10, 'bold'))
        self._preview_subject_label.grid(
            row=1, column=0, columnspan=3, sticky='w', pady=(10, 4))

        self._preview_body = scrolledtext.ScrolledText(
            preview_panel, height=9, wrap='word', state='disabled',
            font=('Segoe UI', 10), relief='solid', borderwidth=1, padx=6, pady=6,
        )
        self._preview_body.grid(row=2, column=0, columnspan=3, sticky='nsew')

        # Send controls
        ctrl_frame = ttk.Frame(frame)
        ctrl_frame.grid(row=2, column=0, sticky='ew', pady=(0, 10))

        self._send_btn = ttk.Button(ctrl_frame, text='Send All Emails', command=self._start_send)
        self._send_btn.pack(side='left', padx=(0, 12))

        self._stop_btn = ttk.Button(ctrl_frame, text='Stop', command=self._request_stop, state='disabled')
        self._stop_btn.pack(side='left', padx=(0, 20))

        self._progress = ttk.Progressbar(ctrl_frame, length=320, mode='determinate')
        self._progress.pack(side='left', padx=(0, 10))

        self._progress_label = ttk.Label(ctrl_frame, text='')
        self._progress_label.pack(side='left')

        # Log
        log_frame = ttk.LabelFrame(frame, text='Send Log', padding=(8, 6))
        log_frame.grid(row=3, column=0, sticky='ew')
        log_frame.columnconfigure(0, weight=1)

        self._log_box = scrolledtext.ScrolledText(
            log_frame, height=8, state='disabled', wrap='word',
            font=('Consolas', 9), relief='solid', borderwidth=1,
        )
        self._log_box.grid(row=0, column=0, sticky='ew')

        ttk.Button(log_frame, text='Clear Log', command=self._clear_log).grid(
            row=1, column=0, sticky='e', pady=(6, 0))

    # ------------------------------------------------------------------
    # SMTP logic
    # ------------------------------------------------------------------

    def _populate_smtp_fields(self):
        for key, var in self._smtp_vars.items():
            var.set(self.config.get(key, ''))
        self._use_ssl.set(self.config.get('use_ssl', False))

        saved_server = self.config.get('server', '')
        matched = 'Custom'
        for name, settings in self.PROVIDERS.items():
            if settings['server'] and settings['server'] == saved_server:
                matched = name
                break
        self._provider_var.set(matched)
        if matched == 'Custom':
            self._custom_frame.grid()
        self._provider_hint.config(text=self.PROVIDER_HINTS[matched])

    def _collect_smtp_config(self) -> dict:
        cfg = {k: v.get().strip() for k, v in self._smtp_vars.items()}
        cfg['use_ssl'] = self._use_ssl.get()
        return cfg

    def _save_smtp(self):
        self.config.update(self._collect_smtp_config())
        es.save_config(self.config)
        self._smtp_status.config(text='Settings saved.', style='Success.TLabel')

    def _test_connection(self):
        cfg = self._collect_smtp_config()
        self._smtp_status.config(text='Connecting…', foreground='#555555')
        self.root.update_idletasks()

        def run():
            ok, msg = es.test_connection(cfg)
            style = 'Success.TLabel' if ok else 'Error.TLabel'
            self.root.after(0, lambda: self._smtp_status.config(text=msg, style=style))

        threading.Thread(target=run, daemon=True).start()

    # ------------------------------------------------------------------
    # Compose — formatting
    # ------------------------------------------------------------------

    def _get_selection_or_line(self) -> tuple[str, str]:
        """Return (start_idx, end_idx) for the current selection or current line."""
        try:
            return self._body_text.index('sel.first'), self._body_text.index('sel.last')
        except tk.TclError:
            return (self._body_text.index('insert linestart'),
                    self._body_text.index('insert lineend'))

    def _toggle_format(self, tag: str):
        start, end = self._get_selection_or_line()
        if start == end:
            return
        existing = self._body_text.tag_names(start)
        if tag in existing:
            self._body_text.tag_remove(tag, start, end)
        else:
            self._body_text.tag_add(tag, start, end)
        self._body_text.focus_set()

    def _insert_link(self):
        """Prompt for URL and wrap selected text (or insert) as a coloured link."""
        url = simpledialog.askstring('Insert Link', 'Enter URL:', parent=self.root)
        if not url:
            return
        if not url.startswith(('http://', 'https://')):
            url = 'https://' + url

        try:
            sel_start = self._body_text.index('sel.first')
            sel_end   = self._body_text.index('sel.last')
            link_text = self._body_text.get(sel_start, sel_end)
        except tk.TclError:
            link_text = url
            sel_start = self._body_text.index('insert')
            self._body_text.insert(sel_start, link_text)
            sel_end = self._body_text.index(f'{sel_start}+{len(link_text)}c')

        tag_name = f'link_{self._link_counter}'
        self._link_counter += 1
        self._link_hrefs[tag_name] = url

        self._body_text.tag_configure(
            tag_name, foreground='#1a73e8', underline=True)
        self._body_text.tag_add(tag_name, sel_start, sel_end)
        self._body_text.focus_set()

    def _clear_format(self):
        """Remove all formatting from the selection."""
        start, end = self._get_selection_or_line()
        for tag in ('bold', 'italic', 'underline'):
            self._body_text.tag_remove(tag, start, end)
        for tag in list(self._link_hrefs.keys()):
            self._body_text.tag_remove(tag, start, end)
        self._body_text.focus_set()

    def _show_context_menu(self, event, widget):
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label='Cut',        command=lambda: widget.event_generate('<<Cut>>'))
        menu.add_command(label='Copy',       command=lambda: widget.event_generate('<<Copy>>'))
        menu.add_command(label='Paste',      command=lambda: widget.event_generate('<<Paste>>'))
        menu.add_separator()
        menu.add_command(label='Select All', command=lambda: widget.tag_add('sel', '1.0', 'end'))
        menu.tk_popup(event.x_root, event.y_root)

    def _has_formatting(self) -> bool:
        """Return True if any formatting tags are present in the body."""
        for tag in ('bold', 'italic', 'underline'):
            if self._body_text.tag_ranges(tag):
                return True
        for tag in self._link_hrefs:
            if self._body_text.tag_ranges(tag):
                return True
        return False

    def _get_body_content(self) -> tuple[str, bool]:
        """
        Return (body, is_html).
        If formatting tags are present, converts to HTML automatically.
        """
        if self._has_formatting():
            return self._body_to_html(), True
        raw = self._body_text.get('1.0', 'end-1c')
        return raw, False

    def _body_to_html(self) -> str:
        """Convert the body Text widget (with tags) to an HTML string."""
        widget = self._body_text
        text = widget.get('1.0', 'end-1c')
        if not text:
            return ''

        FMT_TAGS = ('bold', 'italic', 'underline')

        def tk_to_off(tk_idx) -> int:
            line, col = map(int, str(tk_idx).split('.'))
            lines = text.split('\n')
            return min(sum(len(lines[i]) + 1 for i in range(line - 1)) + col, len(text))

        # Collect all boundaries
        boundaries = {0, len(text)}
        for i, ch in enumerate(text):
            if ch == '\n':
                boundaries.add(i)
                boundaries.add(i + 1)
        for tag in FMT_TAGS:
            for r in widget.tag_ranges(tag):
                boundaries.add(tk_to_off(r))
        for tag in self._link_hrefs:
            for r in widget.tag_ranges(tag):
                boundaries.add(tk_to_off(r))
        boundaries = sorted(boundaries)

        parts = ['<div style="font-family:Arial,sans-serif;font-size:14px;line-height:1.6">']

        for i in range(len(boundaries) - 1):
            s, e = boundaries[i], boundaries[i + 1]
            if s >= len(text):
                break
            seg = text[s:e]
            if not seg:
                continue

            if seg == '\n':
                parts.append('<br>')
                continue

            mid = widget.index(f'1.0 + {s} chars')
            active = set(widget.tag_names(mid))

            escaped = (seg.replace('&', '&amp;')
                          .replace('<', '&lt;')
                          .replace('>', '&gt;'))

            if 'underline' in active:
                escaped = f'<u>{escaped}</u>'
            if 'italic' in active:
                escaped = f'<em>{escaped}</em>'
            if 'bold' in active:
                escaped = f'<strong>{escaped}</strong>'

            for tag, href in self._link_hrefs.items():
                if tag in active:
                    escaped = (f'<a href="{href}" '
                               f'style="color:#1a73e8;text-decoration:underline">'
                               f'{escaped}</a>')
                    break

            parts.append(escaped)

        parts.append('</div>')
        return ''.join(parts)

    def _refresh_vars(self, *_):
        subject = self._subject_var.get()
        body = self._body_text.get('1.0', 'end')
        found = es.extract_variables(subject + ' ' + body)
        if found:
            display = '  '.join(f'{{{{{v}}}}}' for v in found)
            self._detected_vars_label.config(text=f'Detected variables: {display}')
        else:
            self._detected_vars_label.config(text='Detected variables: —')

    # ------------------------------------------------------------------
    # Import CSV logic
    # ------------------------------------------------------------------

    def _load_recipients(self, headers: list[str], rows: list[dict], source: str):
        """Shared helper — populates recipient state and updates UI."""
        self.headers = headers
        self.recipients = rows

        self._email_col_combo['values'] = headers
        matched = next((h for h in headers if h.lower() == 'email'), None)
        self._email_col.set(matched if matched else headers[0])

        self._tree['columns'] = headers
        for col in headers:
            self._tree.heading(col, text=col, anchor='w')
            self._tree.column(col, width=130, minwidth=60, anchor='w')
        self._tree.delete(*self._tree.get_children())
        for row in rows:
            self._tree.insert('', 'end', values=[row.get(h, '') for h in headers])

        count = len(rows)
        self._recipient_count_label.config(
            text=f'{count} recipient{"s" if count != 1 else ""} loaded.')

        email_col = self._email_col.get()
        labels = [r.get(email_col, f'Row {i + 1}') for i, r in enumerate(rows)]
        self._preview_combo['values'] = labels
        if labels:
            self._preview_combo.current(0)
            self._update_preview()

        self._log(f'Loaded {count} recipients from {source}.')

    def _import_csv(self):
        path = filedialog.askopenfilename(
            title='Select CSV file',
            filetypes=[('CSV files', '*.csv'), ('All files', '*.*')],
        )
        if not path:
            return
        try:
            headers, rows = es.load_csv(path)
        except Exception as exc:
            messagebox.showerror('CSV Error', str(exc))
            return
        if not headers:
            messagebox.showwarning('Empty CSV', 'The CSV file has no columns.')
            return
        self._load_recipients(headers, rows, path)

    def _show_csv_example(self):
        """Pop up a window showing a sample CSV and explanation."""
        win = tk.Toplevel(self.root)
        win.title('CSV Format Example')
        win.geometry('640x420')
        win.resizable(True, True)

        ttk.Label(win, text='CSV Format Example', style='Header.TLabel').pack(
            anchor='w', padx=20, pady=(18, 8))

        ttk.Label(
            win,
            text='The first row is the column headers. Each header becomes a {{variable}} you can '
                 'use in your email subject or body. One column must be the email address.',
            wraplength=600, justify='left',
        ).pack(anchor='w', padx=20, pady=(0, 12))

        example = (
            'email,first_name,last_name,company,promo_code\n'
            'alice@example.com,Alice,Smith,Acme Corp,SAVE20\n'
            'bob@example.com,Bob,Jones,Globex Inc,SAVE15\n'
            'carol@example.com,Carol,Lee,Initech,SAVE10'
        )

        txt = tk.Text(win, font=('Consolas', 10), height=6, relief='solid',
                      borderwidth=1, padx=8, pady=6, bg='#f8f8f8')
        txt.insert('1.0', example)
        txt.config(state='disabled')
        txt.pack(fill='x', padx=20, pady=(0, 16))

        ttk.Label(
            win,
            text='In your email you would then write:\n\n'
                 '  Subject:  Hey {{first_name}}, a special offer for {{company}}!\n\n'
                 '  Body:     Hi {{first_name}} {{last_name}},\n'
                 '            Use code {{promo_code}} for 20% off.\n\n'
                 'Each recipient gets their own personalised version.',
            font=('Segoe UI', 10), justify='left', foreground='#333',
        ).pack(anchor='w', padx=20)

        ttk.Button(win, text='Close', command=win.destroy).pack(pady=(12, 16))

    # ------------------------------------------------------------------
    # Builder tab logic
    # ------------------------------------------------------------------

    def _builder_rebuild_tree(self):
        cols = self._builder_cols
        self._btree['columns'] = cols
        for col in cols:
            self._btree.heading(col, text=col, anchor='w')
            self._btree.column(col, width=140, minwidth=60, anchor='w')
        self._btree.delete(*self._btree.get_children())
        for row in self._builder_rows:
            self._btree.insert('', 'end', values=row)
        n = len(self._builder_rows)
        self._builder_status.config(
            text=f'{n} row{"s" if n != 1 else ""}  |  '
                 f'Columns: {", ".join(self._builder_cols)}')

    def _builder_add_column(self):
        name = simpledialog.askstring('Add Column', 'Column name (no spaces — use underscore):',
                                      parent=self.root)
        if not name:
            return
        name = name.strip().replace(' ', '_')
        if name in self._builder_cols:
            messagebox.showwarning('Duplicate', f'Column "{name}" already exists.')
            return
        self._builder_cols.append(name)
        # Extend each row with an empty value
        self._builder_rows = [row + [''] for row in self._builder_rows]
        self._builder_rebuild_tree()

    def _builder_remove_column(self):
        if len(self._builder_cols) <= 1:
            messagebox.showwarning('Cannot Remove', 'At least one column is required.')
            return
        name = simpledialog.askstring(
            'Remove Column',
            f'Column to remove? ({", ".join(self._builder_cols)}):',
            parent=self.root)
        if not name or name not in self._builder_cols:
            return
        idx = self._builder_cols.index(name)
        self._builder_cols.pop(idx)
        self._builder_rows = [row[:idx] + row[idx + 1:] for row in self._builder_rows]
        self._builder_rebuild_tree()

    def _builder_add_row(self):
        self._builder_rows.append([''] * len(self._builder_cols))
        self._builder_rebuild_tree()
        # Select the new row
        children = self._btree.get_children()
        if children:
            self._btree.selection_set(children[-1])
            self._btree.see(children[-1])

    def _builder_delete_row(self):
        sel = self._btree.selection()
        if not sel:
            return
        idx = self._btree.index(sel[0])
        self._builder_rows.pop(idx)
        self._builder_rebuild_tree()

    def _builder_edit_cell(self, event):
        region = self._btree.identify('region', event.x, event.y)
        if region != 'cell':
            return

        row_id = self._btree.identify_row(event.y)
        col_id = self._btree.identify_column(event.x)
        col_idx = int(col_id.replace('#', '')) - 1

        if not row_id:
            return

        bbox = self._btree.bbox(row_id, col_id)
        if not bbox:
            return
        x, y, w, h = bbox

        values = list(self._btree.item(row_id, 'values'))
        current = values[col_idx] if col_idx < len(values) else ''

        entry_var = tk.StringVar(value=current)
        entry = tk.Entry(self._btree, textvariable=entry_var, font=('Segoe UI', 10))
        entry.place(x=x, y=y, width=w, height=h)
        entry.focus_set()
        entry.select_range(0, 'end')

        row_idx = self._btree.index(row_id)

        def commit(_event=None):
            new_val = entry_var.get()
            while len(values) <= col_idx:
                values.append('')
            values[col_idx] = new_val
            self._btree.item(row_id, values=values)
            while len(self._builder_rows[row_idx]) <= col_idx:
                self._builder_rows[row_idx].append('')
            self._builder_rows[row_idx][col_idx] = new_val
            entry.destroy()

        entry.bind('<Return>', commit)
        entry.bind('<Tab>',    commit)
        entry.bind('<FocusOut>', commit)
        entry.bind('<Escape>', lambda _: entry.destroy())

    def _builder_save_csv(self):
        path = filedialog.asksaveasfilename(
            title='Save Recipients as CSV',
            defaultextension='.csv',
            filetypes=[('CSV files', '*.csv')],
        )
        if not path:
            return
        rows_as_dicts = [dict(zip(self._builder_cols, row)) for row in self._builder_rows]
        es.write_csv(path, self._builder_cols, rows_as_dicts)
        self._builder_status.config(text=f'Saved to {path}')

    def _builder_use(self):
        """Copy builder data into self.recipients and update the Send tab."""
        if not self._builder_rows:
            messagebox.showwarning('No Data', 'Add some rows first.')
            return
        rows_as_dicts = [dict(zip(self._builder_cols, row)) for row in self._builder_rows]
        self._load_recipients(self._builder_cols, rows_as_dicts, 'Builder')
        messagebox.showinfo('Done', f'{len(rows_as_dicts)} recipients ready for sending.')

    # ------------------------------------------------------------------
    # Send tab logic
    # ------------------------------------------------------------------

    def _update_preview(self, *_):
        if not self.recipients:
            return
        idx = self._preview_combo.current()
        idx = max(0, min(idx, len(self.recipients) - 1))
        row = self.recipients[idx]
        subject = es.render_template(self._subject_var.get(), row)
        body_raw = self._body_text.get('1.0', 'end-1c')
        body = es.render_template(body_raw, row)
        self._preview_subject_label.config(text=f'Subject: {subject}')
        self._preview_body.config(state='normal')
        self._preview_body.delete('1.0', 'end')
        self._preview_body.insert('1.0', body)
        self._preview_body.config(state='disabled')

    def _start_send(self):
        if self._sending:
            return

        if not self.recipients:
            messagebox.showwarning(
                'No Recipients',
                'Load recipients first:\n'
                '• Import a CSV on the "Import CSV" tab, or\n'
                '• Build a list on the "Build Recipients" tab and click "Use for Sending".')
            return

        subject_tmpl = self._subject_var.get().strip()
        if not subject_tmpl:
            messagebox.showwarning('Missing Subject', 'Enter a subject on the Compose tab.')
            return

        body_tmpl, is_html = self._get_body_content()
        if not body_tmpl.strip():
            messagebox.showwarning('Missing Body', 'Write a message body on the Compose tab.')
            return

        cfg = self._collect_smtp_config()
        if not cfg.get('username') or not cfg.get('password'):
            messagebox.showwarning('SMTP Not Configured', 'Fill in SMTP settings first.')
            return

        email_col = self._email_col.get()
        count = len(self.recipients)

        if not messagebox.askyesno(
            'Confirm Send',
            f'Send {count} personalised email{"s" if count != 1 else ""}?\n\n'
            f'Subject: {subject_tmpl[:80]}\n'
            f'Format: {"HTML" if is_html else "Plain text"}',
        ):
            return

        self._sending = True
        self._stop_flag = False
        self._send_btn.config(state='disabled')
        self._stop_btn.config(state='normal')
        self._progress['maximum'] = count
        self._progress['value'] = 0
        self._progress_label.config(text=f'0 / {count}')

        delay = int(cfg.get('delay_ms', 300)) / 1000

        def on_progress(index, total, success, message):
            self.root.after(0, lambda i=index, m=message: self._handle_progress(i, total, m))

        def stop_flag():
            return self._stop_flag

        def run():
            sent, failed = es.send_email_batch(
                config=cfg,
                recipients=self.recipients,
                email_col=email_col,
                subject_tmpl=subject_tmpl,
                body_tmpl=body_tmpl,
                is_html=is_html,
                on_progress=on_progress,
                stop_flag=stop_flag,
                delay_seconds=delay,
            )
            summary = f'\nFinished — {sent} sent, {failed} failed.'
            self.root.after(0, lambda: self._finish_send(summary))

        threading.Thread(target=run, daemon=True).start()

    def _handle_progress(self, index: int, total: int, message: str):
        self._log(message)
        val = index + 1
        self._progress['value'] = val
        self._progress_label.config(text=f'{val} / {total}')

    def _finish_send(self, summary: str):
        self._log(summary)
        self._sending = False
        self._stop_flag = False
        self._send_btn.config(state='normal')
        self._stop_btn.config(state='disabled')

    def _request_stop(self):
        self._stop_flag = True
        self._log('Stop requested — finishing current email…')
        self._stop_btn.config(state='disabled')

    # ------------------------------------------------------------------
    # Log helpers
    # ------------------------------------------------------------------

    def _log(self, msg: str):
        self._log_box.config(state='normal')
        self._log_box.insert('end', msg + '\n')
        self._log_box.see('end')
        self._log_box.config(state='disabled')

    def _clear_log(self):
        self._log_box.config(state='normal')
        self._log_box.delete('1.0', 'end')
        self._log_box.config(state='disabled')


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == '__main__':
    main()
