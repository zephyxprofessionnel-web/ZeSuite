import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
import json
import csv
import os
import re
import zipfile
import datetime
import xml.etree.ElementTree as ET


class OmniOffice365:
    def __init__(self, root):
        self.root = root
        self.root.title("OmniOffice 365 Pro — Suite Bureautique Complète")
        self.root.geometry("1360x880")
        self.root.minsize(1100, 720)

        # Interception de la fermeture (Croix rouge Windows)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        # Palette de couleurs style Microsoft 365 / LibreOffice
        self.colors = {
            "bg_app": "#F3F2F1",
            "sidebar": "#EAEAEA",
            "word": "#2B579A",
            "excel": "#217346",
            "powerpoint": "#B7472A",
            "home": "#0078D4",
            "card": "#FFFFFF",
            "border": "#D2D0CE",
            "text": "#323130",
            "text_muted": "#605E5C",
        }

        self.current_app = "home"
        self.current_file_path = None
        self.recent_files = []

        self._setup_styles()
        self._build_top_menu()
        self._build_main_layout()
        self._bind_shortcuts()
        self.switch_app("home")

    def _setup_styles(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TFrame", background=self.colors["bg_app"])
        style.configure("TButton", font=("Segoe UI", 9), padding=4)

    def _bind_shortcuts(self):
        self.root.bind("<Control-o>", lambda e: self.open_any_file())
        self.root.bind("<Control-s>", lambda e: self.quick_save_current())
        self.root.bind("<F5>", lambda e: self.ppt_start_slideshow() if self.current_app == "powerpoint" else None)

    def quick_save_current(self):
        if self.current_app == "excel":
            self.save_excel_doc()
        elif self.current_app == "powerpoint":
            self.save_ppt_doc()
        else:
            self.save_word_doc()

    # =========================================================================
    # BARRE DE MENUS SUPÉRIEURE (STYLE LIBREOFFICE / OFFICE)
    # =========================================================================
    def _build_top_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        # 1. Fichier
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Nouveau Document Word", command=self.new_word_doc)
        file_menu.add_command(label="Nouveau Classeur Excel", command=self.new_excel_doc)
        file_menu.add_command(label="Nouvelle Présentation PowerPoint", command=self.new_ppt_doc)
        file_menu.add_separator()
        file_menu.add_command(label="Ouvrir (TXT, ODT, ODS, ODP, CSV, PY...)", command=self.open_any_file, accelerator="Ctrl+O")
        file_menu.add_command(label="Enregistrer sous...", command=self.quick_save_current, accelerator="Ctrl+S")
        file_menu.add_separator()
        file_menu.add_command(label="Quitter", command=self.on_close, accelerator="Alt+F4")
        menubar.add_cascade(label="Fichier", menu=file_menu)

        # 2. Édition
        edit_menu = tk.Menu(menubar, tearoff=0)
        edit_menu.add_command(label="Annuler (Word)", command=lambda: self.word_text.edit_undo())
        edit_menu.add_command(label="Rétablir (Word)", command=lambda: self.word_text.edit_redo())
        edit_menu.add_separator()
        edit_menu.add_command(label="Rechercher et remplacer (Word)", command=self.word_find_dialog)
        edit_menu.add_command(label="Tout sélectionner (Word)", command=lambda: self.word_text.tag_add("sel", "1.0", "end"))
        menubar.add_cascade(label="Édition", menu=edit_menu)

        # 3. Affichage
        view_menu = tk.Menu(menubar, tearoff=0)
        view_menu.add_command(label="Accueil OmniOffice 365", command=lambda: self.switch_app("home"))
        view_menu.add_command(label="OmniWord (Traitement de texte)", command=lambda: self.switch_app("word"))
        view_menu.add_command(label="OmniExcel (Tableur)", command=lambda: self.switch_app("excel"))
        view_menu.add_command(label="OmniPowerPoint (Présentation)", command=lambda: self.switch_app("powerpoint"))
        menubar.add_cascade(label="Affichage", menu=view_menu)

        # 4. Insertion
        insert_menu = tk.Menu(menubar, tearoff=0)
        insert_menu.add_command(label="Insérer Date & Heure (Word)", command=self.word_insert_date)
        insert_menu.add_command(label="Insérer un Tableau Texte (Word)", command=self.word_insert_table)
        insert_menu.add_command(label="Insérer une Ligne de séparation (Word)", command=lambda: self.word_text.insert("insert", "\n" + "─" * 65 + "\n"))
        insert_menu.add_separator()
        insert_menu.add_command(label="Nouvelle Diapositive (PowerPoint)", command=self.ppt_add_slide)
        insert_menu.add_command(label="Dupliquer la Diapositive (PowerPoint)", command=self.ppt_duplicate_slide)
        menubar.add_cascade(label="Insertion", menu=insert_menu)

        # 5. Format
        format_menu = tk.Menu(menubar, tearoff=0)
        format_menu.add_command(label="Gras (Word)", command=lambda: self.word_toggle_tag("bold"))
        format_menu.add_command(label="Italique (Word)", command=lambda: self.word_toggle_tag("italic"))
        format_menu.add_command(label="Souligné (Word)", command=lambda: self.word_toggle_tag("underline"))
        format_menu.add_command(label="Barré (Word)", command=lambda: self.word_toggle_tag("overstrike"))
        format_menu.add_separator()
        format_menu.add_command(label="MAJUSCULES (Sélection Word)", command=lambda: self.word_change_case(True))
        format_menu.add_command(label="minuscules (Sélection Word)", command=lambda: self.word_change_case(False))
        menubar.add_cascade(label="Format", menu=format_menu)

        # 6. Diaporama & Outils
        tools_menu = tk.Menu(menubar, tearoff=0)
        tools_menu.add_command(label="Lancer le Diaporama Plein Écran", command=self.ppt_start_slideshow, accelerator="F5")
        tools_menu.add_command(label="Recalculer toutes les formules Excel", command=self.excel_recalculate_all)
        tools_menu.add_command(label="Statistiques du document Word", command=self.word_show_detailed_stats)
        menubar.add_cascade(label="Outils", menu=tools_menu)

        # 7. Aide
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(
            label="À propos d'OmniOffice 365 Pro",
            command=lambda: messagebox.showinfo(
                "OmniOffice 365 Pro",
                "Suite Bureautique Complète en Python.\nCompatible : .txt, .odt, .ods, .odp, .csv, .py, .json, .md, .omniword, .omnippt"
            )
        )
        menubar.add_cascade(label="Aide", menu=help_menu)

    def on_close(self):
        """Popup Windows officielle à la fermeture"""
        answer = messagebox.askyesnocancel(
            "Fermer OmniOffice 365",
            "Voulez-vous enregistrer votre travail en .txt (ou format associé) avant de quitter ?"
        )
        if answer is True:
            self.quick_save_current()
            self.root.destroy()
        elif answer is False:
            self.root.destroy()

    def _build_main_layout(self):
        # Barre supérieure
        self.top_bar = tk.Frame(self.root, bg=self.colors["home"], height=44)
        self.top_bar.pack(side="top", fill="x")
        self.top_bar.pack_propagate(False)

        self.launcher_btn = tk.Button(
            self.top_bar, text="⠿", font=("Segoe UI", 14, "bold"),
            bg=self.colors["home"], fg="white", bd=0, padx=12, cursor="hand2",
            activebackground="#106EBE", activeforeground="white",
            command=lambda: self.switch_app("home")
        )
        self.launcher_btn.pack(side="left", fill="y")

        self.lbl_app_title = tk.Label(
            self.top_bar, text="OmniOffice 365", font=("Segoe UI", 11, "bold"),
            bg=self.colors["home"], fg="white"
        )
        self.lbl_app_title.pack(side="left", padx=8)

        self.lbl_doc_name = tk.Label(
            self.top_bar, text="— Accueil", font=("Segoe UI", 10),
            bg=self.colors["home"], fg="#E1DFDD"
        )
        self.lbl_doc_name.pack(side="left", padx=4)

        btn_open_global = tk.Button(
            self.top_bar, text="📂 Ouvrir un fichier...", font=("Segoe UI", 9, "bold"),
            bg="#FFFFFF", fg="#323130", bd=0, padx=12, pady=4, cursor="hand2",
            command=self.open_any_file
        )
        btn_open_global.pack(side="right", padx=12, pady=7)

        # Corps principal
        self.body_frame = tk.Frame(self.root, bg=self.colors["bg_app"])
        self.body_frame.pack(fill="both", expand=True)

        # Barre latérale gauche
        self.sidebar = tk.Frame(self.body_frame, bg=self.colors["sidebar"], width=72)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        self.nav_buttons = {}
        apps = [
            ("home", "🏠", "Accueil", self.colors["home"]),
            ("word", "W", "Word", self.colors["word"]),
            ("excel", "X", "Excel", self.colors["excel"]),
            ("powerpoint", "P", "PowerPoint", self.colors["powerpoint"]),
        ]

        for app_id, icon, label, color in apps:
            btn_frame = tk.Frame(self.sidebar, bg=self.colors["sidebar"], pady=6)
            btn_frame.pack(fill="x")

            btn = tk.Button(
                btn_frame, text=f"{icon}\n{label}", font=("Segoe UI", 8, "bold"),
                bg=self.colors["sidebar"], fg=color if app_id != "home" else "#323130",
                bd=0, pady=8, cursor="hand2", activebackground="#D2D0CE",
                command=lambda a=app_id: self.switch_app(a)
            )
            btn.pack(fill="x", padx=4)
            self.nav_buttons[app_id] = btn

        self.workspace = tk.Frame(self.body_frame, bg=self.colors["bg_app"])
        self.workspace.pack(side="left", fill="both", expand=True)

        # Barre d'état
        self.status_bar = tk.Frame(self.root, bg="#F3F2F1", height=24, bd=1, relief="solid")
        self.status_bar.pack(side="bottom", fill="x")
        self.lbl_status = tk.Label(
            self.status_bar, text="Prêt • OmniOffice 365 Pro",
            font=("Segoe UI", 8), bg="#F3F2F1", fg="#605E5C", padx=10
        )
        self.lbl_status.pack(side="left")

        self.views = {}
        self._build_home_view()
        self._build_word_view()
        self._build_excel_view()
        self._build_ppt_view()

    def switch_app(self, app_name):
        self.current_app = app_name
        color = self.colors[app_name]
        titles = {
            "home": "OmniOffice 365",
            "word": "OmniWord",
            "excel": "OmniExcel",
            "powerpoint": "OmniPowerPoint"
        }

        self.top_bar.config(bg=color)
        self.launcher_btn.config(bg=color)
        self.lbl_app_title.config(bg=color, text=titles[app_name])
        self.lbl_doc_name.config(bg=color)

        for aid, btn in self.nav_buttons.items():
            btn.config(bg="#FFFFFF" if aid == app_name else self.colors["sidebar"])

        for view in self.views.values():
            view.pack_forget()
        self.views[app_name].pack(fill="both", expand=True)

    # =========================================================================
    # 1. ACCUEIL
    # =========================================================================
    def _build_home_view(self):
        frame = tk.Frame(self.workspace, bg=self.colors["bg_app"], padx=40, pady=30)
        self.views["home"] = frame

        tk.Label(
            frame, text="Bienvenue dans OmniOffice 365 Pro",
            font=("Segoe UI", 22, "bold"), bg=self.colors["bg_app"], fg="#323130"
        ).pack(anchor="w")

        tk.Label(
            frame, text="Compatible Microsoft Office, LibreOffice (.odt, .ods, .odp), Texte (.txt), CSV et Code",
            font=("Segoe UI", 11), bg=self.colors["bg_app"], fg="#605E5C"
        ).pack(anchor="w", pady=(2, 25))

        cards_row = tk.Frame(frame, bg=self.colors["bg_app"])
        cards_row.pack(fill="x", pady=10)

        creations = [
            ("Document Word", "Mise en forme complète, zoom,\ntableaux, .txt, .odt, .md", self.colors["word"], self.new_word_doc),
            ("Classeur Excel", "Formules (=SOMME, =MOYENNE),\ncouleurs cellules, .csv, .ods", self.colors["excel"], self.new_excel_doc),
            ("Présentation PowerPoint", "Text scale, gras/italique, alignement,\nthèmes & diaporama", self.colors["powerpoint"], self.new_ppt_doc),
            ("Ouvrir un fichier PC", "Ouvrir n'importe quel fichier\ndepuis votre ordinateur", "#323130", self.open_any_file),
        ]

        for title, desc, color, cmd in creations:
            card = tk.Frame(cards_row, bg="white", bd=1, relief="solid", padx=20, pady=20, width=265, height=155)
            card.pack(side="left", padx=(0, 18))
            card.pack_propagate(False)

            bar = tk.Frame(card, bg=color, height=5)
            bar.pack(fill="x", side="top", pady=(0, 12))

            tk.Label(card, text=title, font=("Segoe UI", 12, "bold"), bg="white", fg=color).pack(anchor="w")
            tk.Label(card, text=desc, font=("Segoe UI", 9), bg="white", fg="#605E5C", justify="left").pack(anchor="w", pady=8)

            btn = tk.Button(
                card, text="Ouvrir →", font=("Segoe UI", 9, "bold"),
                bg=color, fg="white", bd=0, padx=12, pady=4, cursor="hand2", command=cmd
            )
            btn.pack(anchor="e", side="bottom")

        tk.Label(
            frame, text="Fichiers ouverts récemment (Double-cliquez pour ouvrir)",
            font=("Segoe UI", 13, "bold"), bg=self.colors["bg_app"], fg="#323130"
        ).pack(anchor="w", pady=(35, 10))

        self.recent_listbox = tk.Listbox(
            frame, font=("Segoe UI", 10), bg="white", fg="#323130",
            bd=1, relief="solid", height=10, selectbackground="#0078D4"
        )
        self.recent_listbox.pack(fill="both", expand=True)
        self.recent_listbox.bind("<Double-Button-1>", self._open_selected_recent)

    def _add_recent(self, filepath):
        if filepath in self.recent_files:
            self.recent_files.remove(filepath)
        self.recent_files.insert(0, filepath)
        self.recent_listbox.delete(0, tk.END)
        for f in self.recent_files:
            self.recent_listbox.insert(tk.END, f"📄  {os.path.basename(f)}   —   {f}")

    def _open_selected_recent(self, event):
        sel = self.recent_listbox.curselection()
        if sel:
            path = self.recent_files[sel[0]]
            if os.path.exists(path):
                self.load_file_by_path(path)

    # =========================================================================
    # 2. OMNIWORD (TRAITEMENT DE TEXTE COMPLET)
    # =========================================================================
    def _build_word_view(self):
        frame = tk.Frame(self.workspace, bg="#E1DFDD")
        self.views["word"] = frame

        # Ruban ligne 1 : Fichier & Police & Styles
        ribbon1 = tk.Frame(frame, bg="#FFFFFF", bd=1, relief="solid", padx=8, pady=5)
        ribbon1.pack(fill="x")

        tk.Button(ribbon1, text="📄 Nouveau", command=self.new_word_doc, bd=0, bg="#F3F2F1", padx=6).pack(side="left", padx=2)
        tk.Button(ribbon1, text="📂 Ouvrir", command=self.open_any_file, bd=0, bg="#F3F2F1", padx=6).pack(side="left", padx=2)
        tk.Button(ribbon1, text="💾 Enregistrer (.txt)", command=self.save_word_doc, bd=0, bg="#2B579A", fg="white", padx=8).pack(side="left", padx=2)

        ttk.Separator(ribbon1, orient="vertical").pack(side="left", fill="y", padx=6)

        self.word_font_family = tk.StringVar(value="Calibri")
        fonts_cb = ttk.Combobox(ribbon1, textvariable=self.word_font_family, values=["Calibri", "Arial", "Times New Roman", "Segoe UI", "Consolas", "Georgia", "Verdana"], width=14, state="readonly")
        fonts_cb.pack(side="left", padx=2)
        fonts_cb.bind("<<ComboboxSelected>>", self.word_update_font)

        self.word_font_size = tk.StringVar(value="12")
        size_cb = ttk.Combobox(ribbon1, textvariable=self.word_font_size, values=["8", "9", "10", "11", "12", "14", "16", "18", "20", "24", "28", "36", "48", "72"], width=4, state="readonly")
        size_cb.pack(side="left", padx=2)
        size_cb.bind("<<ComboboxSelected>>", self.word_update_font)

        ttk.Separator(ribbon1, orient="vertical").pack(side="left", fill="y", padx=6)

        # Gras, Italique, Souligné, Barré
        tk.Button(ribbon1, text="G", font=("Segoe UI", 9, "bold"), width=3, command=lambda: self.word_toggle_tag("bold")).pack(side="left", padx=1)
        tk.Button(ribbon1, text="I", font=("Segoe UI", 9, "italic"), width=3, command=lambda: self.word_toggle_tag("italic")).pack(side="left", padx=1)
        tk.Button(ribbon1, text="S", font=("Segoe UI", 9, "underline"), width=3, command=lambda: self.word_toggle_tag("underline")).pack(side="left", padx=1)
        tk.Button(ribbon1, text="ab", font=("Segoe UI", 9, "overstrike"), width=3, command=lambda: self.word_toggle_tag("overstrike")).pack(side="left", padx=1)

        tk.Button(ribbon1, text="🎨 Couleur", command=self.word_change_color, padx=5).pack(side="left", padx=2)
        tk.Button(ribbon1, text="🖍 Surligner", command=self.word_highlight_text, padx=5).pack(side="left", padx=2)

        ttk.Separator(ribbon1, orient="vertical").pack(side="left", fill="y", padx=6)

        # Alignements
        tk.Button(ribbon1, text="⬅ Gauche", command=lambda: self.word_align("left")).pack(side="left", padx=1)
        tk.Button(ribbon1, text="⬛ Centre", command=lambda: self.word_align("center")).pack(side="left", padx=1)
        tk.Button(ribbon1, text="➡ Droite", command=lambda: self.word_align("right")).pack(side="left", padx=1)

        # Ruban ligne 2 : Insertions & Outils rapides
        ribbon2 = tk.Frame(frame, bg="#F8F8F8", bd=1, relief="solid", padx=8, pady=4)
        ribbon2.pack(fill="x")

        tk.Button(ribbon2, text="• Puce", command=lambda: self.word_text.insert("insert", "• "), bd=0, bg="#EAEAEA", padx=6).pack(side="left", padx=2)
        tk.Button(ribbon2, text="📅 Date", command=self.word_insert_date, bd=0, bg="#EAEAEA", padx=6).pack(side="left", padx=2)
        tk.Button(ribbon2, text="⊞ Tableau", command=self.word_insert_table, bd=0, bg="#EAEAEA", padx=6).pack(side="left", padx=2)
        tk.Button(ribbon2, text="AA", command=lambda: self.word_change_case(True), bd=0, bg="#EAEAEA", padx=6).pack(side="left", padx=2)
        tk.Button(ribbon2, text="aa", command=lambda: self.word_change_case(False), bd=0, bg="#EAEAEA", padx=6).pack(side="left", padx=2)
        tk.Button(ribbon2, text="🔍 Rechercher / Remplacer", command=self.word_find_dialog, bd=0, bg="#EAEAEA", padx=8).pack(side="left", padx=6)

        # Contrôle Zoom à droite
        tk.Label(ribbon2, text="Zoom :", bg="#F8F8F8", font=("Segoe UI", 8)).pack(side="right", padx=(6, 2))
        tk.Button(ribbon2, text="+", width=2, command=lambda: self.word_zoom(2)).pack(side="right", padx=1)
        tk.Button(ribbon2, text="-", width=2, command=lambda: self.word_zoom(-2)).pack(side="right", padx=1)

        # Feuille A4
        page_container = tk.Frame(frame, bg="#E1DFDD", pady=15)
        page_container.pack(fill="both", expand=True)

        scrollbar = tk.Scrollbar(page_container)
        scrollbar.pack(side="right", fill="y")

        self.word_text = tk.Text(
            page_container, width=88, font=("Calibri", 12),
            bg="white", fg="#111111", padx=55, pady=45,
            undo=True, wrap="word", relief="solid", bd=1,
            yscrollcommand=scrollbar.set
        )
        self.word_text.pack(fill="y", expand=True)
        scrollbar.config(command=self.word_text.yview)

        self.word_update_font()
        self.word_text.tag_configure("left", justify="left")
        self.word_text.tag_configure("center", justify="center")
        self.word_text.tag_configure("right", justify="right")
        self.word_text.bind("<KeyRelease>", self.word_update_stats)

    def new_word_doc(self):
        self.switch_app("word")
        self.word_text.delete("1.0", tk.END)
        self.current_file_path = None
        self.lbl_doc_name.config(text="— Document1.txt")
        self.word_update_stats()

    def word_update_font(self, event=None):
        fam = self.word_font_family.get()
        sz = int(self.word_font_size.get())
        self.word_text.config(font=(fam, sz))
        self.word_text.tag_configure("bold", font=(fam, sz, "bold"))
        self.word_text.tag_configure("italic", font=(fam, sz, "italic"))
        self.word_text.tag_configure("underline", font=(fam, sz, "underline"))
        self.word_text.tag_configure("overstrike", font=(fam, sz, "overstrike"))

    def word_zoom(self, delta):
        cur = int(self.word_font_size.get())
        nxt = max(8, min(72, cur + delta))
        self.word_font_size.set(str(nxt))
        self.word_update_font()

    def word_toggle_tag(self, tag_name):
        try:
            current_tags = self.word_text.tag_names("sel.first")
            if tag_name in current_tags:
                self.word_text.tag_remove(tag_name, "sel.first", "sel.last")
            else:
                self.word_text.tag_add(tag_name, "sel.first", "sel.last")
        except tk.TclError:
            pass

    def word_change_color(self):
        color = colorchooser.askcolor(title="Couleur du texte")[1]
        if color:
            try:
                tag_name = f"fg_{color}"
                self.word_text.tag_configure(tag_name, foreground=color)
                self.word_text.tag_add(tag_name, "sel.first", "sel.last")
            except tk.TclError:
                pass

    def word_highlight_text(self):
        color = colorchooser.askcolor(title="Couleur de surlignage")[1]
        if color:
            try:
                tag_name = f"bg_{color}"
                self.word_text.tag_configure(tag_name, background=color)
                self.word_text.tag_add(tag_name, "sel.first", "sel.last")
            except tk.TclError:
                pass

    def word_align(self, align):
        for a in ("left", "center", "right"):
            self.word_text.tag_remove(a, "insert linestart", "insert lineend")
        self.word_text.tag_add(align, "insert linestart", "insert lineend")

    def word_insert_date(self):
        now = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
        self.word_text.insert("insert", now)

    def word_insert_table(self):
        table_str = (
            "\n+--------------------+--------------------+--------------------+\n"
            "| Colonne 1          | Colonne 2          | Colonne 3          |\n"
            "+--------------------+--------------------+--------------------+\n"
            "| Donnée A           | Donnée B           | Donnée C           |\n"
            "+--------------------+--------------------+--------------------+\n"
        )
        self.word_text.insert("insert", table_str)

    def word_change_case(self, upper=True):
        try:
            txt = self.word_text.get("sel.first", "sel.last")
            self.word_text.delete("sel.first", "sel.last")
            self.word_text.insert("insert", txt.upper() if upper else txt.lower())
        except tk.TclError:
            pass

    def word_find_dialog(self):
        win = tk.Toplevel(self.root)
        win.title("Rechercher et remplacer")
        win.geometry("360x140")
        win.resizable(False, False)

        tk.Label(win, text="Rechercher :").grid(row=0, column=0, padx=10, pady=8, sticky="w")
        e_find = tk.Entry(win, width=25)
        e_find.grid(row=0, column=1, padx=5)

        tk.Label(win, text="Remplacer par :").grid(row=1, column=0, padx=10, pady=8, sticky="w")
        e_rep = tk.Entry(win, width=25)
        e_rep.grid(row=1, column=1, padx=5)

        def do_replace():
            target = e_find.get()
            rep = e_rep.get()
            if target:
                content = self.word_text.get("1.0", tk.END)
                new_content = content.replace(target, rep)
                self.word_text.delete("1.0", tk.END)
                self.word_text.insert("1.0", new_content[:-1])

        tk.Button(win, text="Tout remplacer", bg="#2B579A", fg="white", command=do_replace).grid(row=2, column=1, pady=10, sticky="e")

    def word_update_stats(self, event=None):
        content = self.word_text.get("1.0", "end-1c")
        words = len(content.split())
        chars = len(content)
        lines = int(self.word_text.index("end-1c").split(".")[0])
        self.lbl_status.config(text=f"OmniWord  |  Lignes : {lines}  |  Mots : {words}  |  Caractères : {chars}")

    def word_show_detailed_stats(self):
        content = self.word_text.get("1.0", "end-1c")
        words = len(content.split())
        chars = len(content)
        chars_no_space = len(content.replace(" ", "").replace("\n", ""))
        messagebox.showinfo("Statistiques du document", f"Mots : {words}\nCaractères (avec espaces) : {chars}\nCaractères (sans espaces) : {chars_no_space}")

    def save_word_doc(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[
                ("Fichier Texte (.txt)", "*.txt"),
                ("Markdown (.md)", "*.md"),
                ("Code Python (.py)", "*.py"),
                ("Tous les fichiers", "*.*")
            ]
        )
        if not path:
            return
        content = self.word_text.get("1.0", "end-1c")
        with open(path, "w", encoding="utf-8") as f:
            f.write(content)
        self.current_file_path = path
        self.lbl_doc_name.config(text=f"— {os.path.basename(path)}")
        self._add_recent(path)
        messagebox.showinfo("Enregistré", f"Fichier enregistré avec succès :\n{path}")

    # =========================================================================
    # 3. OMNIEXCEL (TABLEUR ENRICHI)
    # =========================================================================
    def _build_excel_view(self):
        frame = tk.Frame(self.workspace, bg="#F3F2F1")
        self.views["excel"] = frame

        ribbon = tk.Frame(frame, bg="#FFFFFF", bd=1, relief="solid", padx=8, pady=5)
        ribbon.pack(fill="x")

        tk.Button(ribbon, text="📊 Nouveau", command=self.new_excel_doc, bd=0, bg="#F3F2F1", padx=6).pack(side="left", padx=2)
        tk.Button(ribbon, text="📂 Ouvrir", command=self.open_any_file, bd=0, bg="#F3F2F1", padx=6).pack(side="left", padx=2)
        tk.Button(ribbon, text="💾 Exporter CSV", command=self.save_excel_doc, bd=0, bg="#217346", fg="white", padx=8).pack(side="left", padx=2)

        ttk.Separator(ribbon, orient="vertical").pack(side="left", fill="y", padx=6)

        # Mise en forme cellule
        tk.Button(ribbon, text="G", font=("Segoe UI", 9, "bold"), width=3, command=self.excel_toggle_bold).pack(side="left", padx=1)
        tk.Button(ribbon, text="🎨 Fond Cellule", command=self.excel_cell_bg, padx=5).pack(side="left", padx=2)
        tk.Button(ribbon, text="🖍 Texte Cellule", command=self.excel_cell_fg, padx=5).pack(side="left", padx=2)

        ttk.Separator(ribbon, orient="vertical").pack(side="left", fill="y", padx=6)

        # Alignement cellule
        tk.Button(ribbon, text="⬅", width=3, command=lambda: self.excel_cell_align("left")).pack(side="left", padx=1)
        tk.Button(ribbon, text="⬛", width=3, command=lambda: self.excel_cell_align("center")).pack(side="left", padx=1)
        tk.Button(ribbon, text="➡", width=3, command=lambda: self.excel_cell_align("right")).pack(side="left", padx=1)

        ttk.Separator(ribbon, orient="vertical").pack(side="left", fill="y", padx=6)

        # Formules rapides
        tk.Button(ribbon, text="∑ SOMME", command=lambda: self.excel_insert_quick_formula("SOMME"), bg="#E1DFDD", bd=0, padx=6).pack(side="left", padx=2)
        tk.Button(ribbon, text="Ø MOYENNE", command=lambda: self.excel_insert_quick_formula("MOYENNE"), bg="#E1DFDD", bd=0, padx=6).pack(side="left", padx=2)
        tk.Button(ribbon, text="🔄 Recalculer", command=self.excel_recalculate_all, bg="#E1DFDD", bd=0, padx=6).pack(side="left", padx=2)

        # Barre de formule
        formula_bar = tk.Frame(frame, bg="#F3F2F1", pady=4, padx=10)
        formula_bar.pack(fill="x")

        self.lbl_active_cell = tk.Label(formula_bar, text="A1", width=6, bg="white", bd=1, relief="solid", font=("Segoe UI", 9, "bold"))
        self.lbl_active_cell.pack(side="left", padx=(0, 6))

        tk.Label(formula_bar, text="fx", font=("Segoe UI", 10, "italic", "bold"), bg="#F3F2F1", fg="#217346").pack(side="left", padx=4)

        self.formula_entry = tk.Entry(formula_bar, font=("Consolas", 10), bg="white", bd=1, relief="solid")
        self.formula_entry.pack(side="left", fill="x", expand=True)
        self.formula_entry.bind("<Return>", self.excel_apply_formula_bar)

        grid_container = tk.Frame(frame, bg="white")
        grid_container.pack(fill="both", expand=True, padx=5, pady=5)

        self.excel_rows = 26
        self.excel_cols = 13
        self.excel_cells = {}
        self.excel_raw_data = {}
        self.excel_bold_cells = set()
        self.selected_cell = (0, 0)

        tk.Label(grid_container, text="", width=4, bg="#EAEAEA", bd=1, relief="solid").grid(row=0, column=0, sticky="nsew")
        for c in range(self.excel_cols):
            col_letter = chr(65 + c)
            tk.Label(grid_container, text=col_letter, width=11, bg="#EAEAEA", font=("Segoe UI", 9, "bold"), bd=1, relief="solid").grid(row=0, column=c + 1, sticky="nsew")

        for r in range(self.excel_rows):
            tk.Label(grid_container, text=str(r + 1), width=4, bg="#EAEAEA", font=("Segoe UI", 9, "bold"), bd=1, relief="solid").grid(row=r + 1, column=0, sticky="nsew")
            for c in range(self.excel_cols):
                entry = tk.Entry(grid_container, width=11, font=("Segoe UI", 9), bd=1, relief="solid")
                entry.grid(row=r + 1, column=c + 1, sticky="nsew")
                entry.bind("<FocusIn>", lambda e, row=r, col=c: self.excel_on_focus(row, col))
                entry.bind("<FocusOut>", lambda e, row=r, col=c: self.excel_on_blur(row, col))
                entry.bind("<Return>", lambda e, row=r, col=c: self.excel_on_enter(row, col))
                self.excel_cells[(r, c)] = entry
                self.excel_raw_data[(r, c)] = ""

    def new_excel_doc(self):
        self.switch_app("excel")
        for (r, c), entry in self.excel_cells.items():
            entry.delete(0, tk.END)
            entry.config(bg="white", fg="black", font=("Segoe UI", 9), justify="left")
            self.excel_raw_data[(r, c)] = ""
        self.excel_bold_cells.clear()
        self.lbl_doc_name.config(text="— Classeur1.csv")

    def excel_toggle_bold(self):
        rc = self.selected_cell
        entry = self.excel_cells[rc]
        if rc in self.excel_bold_cells:
            self.excel_bold_cells.remove(rc)
            entry.config(font=("Segoe UI", 9))
        else:
            self.excel_bold_cells.add(rc)
            entry.config(font=("Segoe UI", 9, "bold"))

    def excel_cell_bg(self):
        color = colorchooser.askcolor(title="Couleur de fond de la cellule")[1]
        if color:
            self.excel_cells[self.selected_cell].config(bg=color)

    def excel_cell_fg(self):
        color = colorchooser.askcolor(title="Couleur du texte de la cellule")[1]
        if color:
            self.excel_cells[self.selected_cell].config(fg=color)

    def excel_cell_align(self, align):
        self.excel_cells[self.selected_cell].config(justify=align)

    def excel_insert_quick_formula(self, func):
        r, c = self.selected_cell
        col_let = chr(65 + c)
        end_r = max(1, r)
        formula = f"={func}({col_let}1:{col_let}{end_r})"
        self.excel_raw_data[(r, c)] = formula
        self.formula_entry.delete(0, tk.END)
        self.formula_entry.insert(0, formula)
        self.excel_recalculate_all()

    def _cell_name_to_coords(self, name):
        match = re.match(r"^([A-Z]+)(\d+)$", name.upper().strip())
        if not match:
            return None
        col_str, row_str = match.groups()
        col = ord(col_str[0]) - 65
        row = int(row_str) - 1
        return (row, col)

    def excel_on_focus(self, r, c):
        self.selected_cell = (r, c)
        cell_name = f"{chr(65+c)}{r+1}"
        self.lbl_active_cell.config(text=cell_name)
        raw = self.excel_raw_data.get((r, c), "")
        self.formula_entry.delete(0, tk.END)
        self.formula_entry.insert(0, raw)
        self.excel_cells[(r, c)].delete(0, tk.END)
        self.excel_cells[(r, c)].insert(0, raw)

    def excel_on_blur(self, r, c):
        val = self.excel_cells[(r, c)].get()
        self.excel_raw_data[(r, c)] = val
        self.excel_recalculate_all()

    def excel_on_enter(self, r, c):
        self.excel_on_blur(r, c)
        if r + 1 < self.excel_rows:
            self.excel_cells[(r + 1, c)].focus_set()

    def excel_apply_formula_bar(self, event=None):
        r, c = self.selected_cell
        val = self.formula_entry.get()
        self.excel_raw_data[(r, c)] = val
        self.excel_recalculate_all()

    def _eval_cell(self, r, c, visited=None):
        if visited is None:
            visited = set()
        if (r, c) in visited:
            return "#REF!"
        visited.add((r, c))

        raw = self.excel_raw_data.get((r, c), "").strip()
        if not raw.startswith("="):
            try:
                return float(raw) if "." in raw else int(raw)
            except ValueError:
                return raw

        expr = raw[1:].upper()

        def eval_range(func_name, match):
            start_c = self._cell_name_to_coords(match.group(1))
            end_c = self._cell_name_to_coords(match.group(2))
            if not start_c or not end_c:
                return "0"
            vals = []
            for rr in range(min(start_c[0], end_c[0]), max(start_c[0], end_c[0]) + 1):
                for cc in range(min(start_c[1], end_c[1]), max(start_c[1], end_c[1]) + 1):
                    v = self._eval_cell(rr, cc, visited.copy())
                    if isinstance(v, (int, float)):
                        vals.append(v)
            if not vals:
                return "0"
            if func_name == "SOMME":
                return str(sum(vals))
            elif func_name == "MOYENNE":
                return str(round(sum(vals) / len(vals), 2))
            elif func_name == "MIN":
                return str(min(vals))
            elif func_name == "MAX":
                return str(max(vals))
            return "0"

        for fn in ("SOMME", "MOYENNE", "MIN", "MAX"):
            expr = re.sub(rf"{fn}\(([A-Z]+\d+):([A-Z]+\d+)\)", lambda m, f=fn: eval_range(f, m), expr)

        def replace_ref(match):
            coords = self._cell_name_to_coords(match.group(0))
            if coords:
                v = self._eval_cell(coords[0], coords[1], visited.copy())
                return str(v) if isinstance(v, (int, float)) else "0"
            return "0"

        expr = re.sub(r"\b[A-Z]+\d+\b", replace_ref, expr)

        try:
            res = eval(expr, {"__builtins__": None}, {})
            return round(res, 4) if isinstance(res, float) else res
        except Exception:
            return "#ERREUR!"

    def excel_recalculate_all(self):
        for (r, c), entry in self.excel_cells.items():
            if self.root.focus_get() == entry:
                continue
            val = self._eval_cell(r, c)
            entry.delete(0, tk.END)
            entry.insert(0, str(val))

    def save_excel_doc(self):
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("Fichier CSV", "*.csv"), ("Fichier Texte", "*.txt")])
        if not path:
            return
        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f, delimiter=";")
            for r in range(self.excel_rows):
                row_vals = [self.excel_cells[(r, c)].get() for c in range(self.excel_cols)]
                if any(row_vals):
                    writer.writerow(row_vals)
        self._add_recent(path)
        self.lbl_doc_name.config(text=f"— {os.path.basename(path)}")
        messagebox.showinfo("Exporté", f"Tableur enregistré sous :\n{path}")

    # =========================================================================
    # 4. OMNIPOWERPOINT (ULTRA-COMPLET : TEXT SCALE, ALIGNEMENT, GRAS, ITALIQUE, THÈMES)
    # =========================================================================
    def _build_ppt_view(self):
        frame = tk.Frame(self.workspace, bg="#E1DFDD")
        self.views["powerpoint"] = frame

        self.slides = [self._default_slide("Titre de la présentation", "Cliquez pour ajouter un sous-titre ou du contenu.")]
        self.current_slide_idx = 0

        # Ruban PowerPoint Ligne 1 : Gestion diapos + Thèmes + Diaporama
        ribbon1 = tk.Frame(frame, bg="#FFFFFF", bd=1, relief="solid", padx=8, pady=5)
        ribbon1.pack(fill="x")

        tk.Button(ribbon1, text="➕ Nouvelle diapo", command=self.ppt_add_slide, bg="#B7472A", fg="white", bd=0, padx=8).pack(side="left", padx=2)
        tk.Button(ribbon1, text="📋 Dupliquer", command=self.ppt_duplicate_slide, bg="#F3F2F1", bd=0, padx=6).pack(side="left", padx=2)
        tk.Button(ribbon1, text="⬆", command=lambda: self.ppt_move_slide(-1), bg="#F3F2F1", bd=0, padx=5).pack(side="left", padx=1)
        tk.Button(ribbon1, text="⬇", command=lambda: self.ppt_move_slide(1), bg="#F3F2F1", bd=0, padx=5).pack(side="left", padx=1)
        tk.Button(ribbon1, text="🗑 Supprimer", command=self.ppt_del_slide, bg="#F3F2F1", bd=0, padx=6).pack(side="left", padx=2)
        tk.Button(ribbon1, text="💾 Enregistrer", command=self.save_ppt_doc, bg="#F3F2F1", bd=0, padx=6).pack(side="left", padx=2)

        ttk.Separator(ribbon1, orient="vertical").pack(side="left", fill="y", padx=6)

        # Thèmes de diapositives
        tk.Label(ribbon1, text="Thème :", bg="white", font=("Segoe UI", 8, "bold")).pack(side="left", padx=2)
        tk.Button(ribbon1, text="Clair", command=lambda: self.ppt_apply_theme("#FFFFFF", "#B7472A", "#323130"), bg="#FFFFFF", bd=1, padx=5).pack(side="left", padx=1)
        tk.Button(ribbon1, text="Sombre", command=lambda: self.ppt_apply_theme("#1E1E1E", "#FF7A59", "#F3F2F1"), bg="#1E1E1E", fg="white", bd=1, padx=5).pack(side="left", padx=1)
        tk.Button(ribbon1, text="Bleu Pro", command=lambda: self.ppt_apply_theme("#0F2C59", "#F8F0E5", "#EADBC8"), bg="#0F2C59", fg="white", bd=1, padx=5).pack(side="left", padx=1)
        tk.Button(ribbon1, text="🎨 Fond", command=self.ppt_custom_bg, bg="#F3F2F1", bd=1, padx=5).pack(side="left", padx=2)

        tk.Button(
            ribbon1, text="▶ Diaporama (F5)", command=self.ppt_start_slideshow,
            bg="#107C10", fg="white", font=("Segoe UI", 9, "bold"), bd=0, padx=12
        ).pack(side="right", padx=6)

        # Ruban PowerPoint Ligne 2 : OPTIONS DE TEXTE (Text Scale, Gras, Italique, Souligné, Alignement, Couleur)
        ribbon2 = tk.Frame(frame, bg="#F8F8F8", bd=1, relief="solid", padx=8, pady=5)
        ribbon2.pack(fill="x")

        tk.Label(ribbon2, text="Texte :", bg="#F8F8F8", font=("Segoe UI", 8, "bold")).pack(side="left", padx=(0, 4))

        # Police
        self.ppt_font_var = tk.StringVar(value="Segoe UI")
        ppt_font_cb = ttk.Combobox(ribbon2, textvariable=self.ppt_font_var, values=["Segoe UI", "Arial", "Calibri", "Times New Roman", "Consolas", "Georgia"], width=13, state="readonly")
        ppt_font_cb.pack(side="left", padx=2)
        ppt_font_cb.bind("<<ComboboxSelected>>", lambda e: self.ppt_update_text_style())

        # Text Scale (Taille du texte)
        tk.Label(ribbon2, text="Taille (Scale) :", bg="#F8F8F8", font=("Segoe UI", 8)).pack(side="left", padx=(6, 2))
        tk.Button(ribbon2, text="A-", width=3, command=lambda: self.ppt_scale_text(-2)).pack(side="left", padx=1)
        self.ppt_scale_var = tk.StringVar(value="16")
        ppt_size_cb = ttk.Combobox(ribbon2, textvariable=self.ppt_scale_var, values=["10", "12", "14", "16", "18", "20", "24", "28", "32", "36", "44", "54"], width=4, state="readonly")
        ppt_size_cb.pack(side="left", padx=1)
        ppt_size_cb.bind("<<ComboboxSelected>>", lambda e: self.ppt_update_text_style())
        tk.Button(ribbon2, text="A+", width=3, command=lambda: self.ppt_scale_text(2)).pack(side="left", padx=1)

        ttk.Separator(ribbon2, orient="vertical").pack(side="left", fill="y", padx=6)

        # Gras, Italique, Souligné, Barré
        tk.Button(ribbon2, text="G", font=("Segoe UI", 9, "bold"), width=3, command=lambda: self.ppt_toggle_style("bold")).pack(side="left", padx=1)
        tk.Button(ribbon2, text="I", font=("Segoe UI", 9, "italic"), width=3, command=lambda: self.ppt_toggle_style("italic")).pack(side="left", padx=1)
        tk.Button(ribbon2, text="S", font=("Segoe UI", 9, "underline"), width=3, command=lambda: self.ppt_toggle_style("underline")).pack(side="left", padx=1)
        tk.Button(ribbon2, text="ab", font=("Segoe UI", 9, "overstrike"), width=3, command=lambda: self.ppt_toggle_style("overstrike")).pack(side="left", padx=1)

        # Couleur du texte & Surlignage
        tk.Button(ribbon2, text="🎨 Couleur Texte", command=self.ppt_text_color, padx=5).pack(side="left", padx=3)
        tk.Button(ribbon2, text="🖍 Surligner", command=self.ppt_highlight_text, padx=5).pack(side="left", padx=2)

        ttk.Separator(ribbon2, orient="vertical").pack(side="left", fill="y", padx=6)

        # Alignement : Gauche, Milieu (Centre), Droite
        tk.Button(ribbon2, text="⬅ Gauche", command=lambda: self.ppt_set_align("left")).pack(side="left", padx=1)
        tk.Button(ribbon2, text="⬛ Milieu", command=lambda: self.ppt_set_align("center")).pack(side="left", padx=1)
        tk.Button(ribbon2, text="➡ Droite", command=lambda: self.ppt_set_align("right")).pack(side="left", padx=1)

        ttk.Separator(ribbon2, orient="vertical").pack(side="left", fill="y", padx=6)

        # Insertion rapide dans diapo
        tk.Button(ribbon2, text="• Puce", command=lambda: self.ppt_body_text.insert("insert", "• "), bg="#EAEAEA", bd=0, padx=6).pack(side="left", padx=2)

        # Corps : Miniatures à gauche + Éditeur de diapositive à droite
        body = tk.Frame(frame, bg="#E1DFDD")
        body.pack(fill="both", expand=True, padx=15, pady=15)

        self.slide_listbox = tk.Listbox(body, width=25, font=("Segoe UI", 10), bg="#F3F2F1", bd=1, relief="solid")
        self.slide_listbox.pack(side="left", fill="y", padx=(0, 15))
        self.slide_listbox.bind("<<ListboxSelect>>", self.ppt_select_slide)

        self.slide_editor = tk.Frame(body, bg="white", bd=2, relief="solid", padx=45, pady=35)
        self.slide_editor.pack(side="left", fill="both", expand=True)

        self.ppt_title_entry = tk.Entry(self.slide_editor, font=("Segoe UI", 28, "bold"), fg="#B7472A", bg="white", bd=0, justify="center")
        self.ppt_title_entry.pack(fill="x", pady=(15, 15))
        self.ppt_title_entry.bind("<KeyRelease>", self.ppt_save_current_fields)

        ttk.Separator(self.slide_editor, orient="horizontal").pack(fill="x", pady=10)

        self.ppt_body_text = tk.Text(self.slide_editor, font=("Segoe UI", 16), fg="#323130", bg="white", bd=0, wrap="word", height=12)
        self.ppt_body_text.pack(fill="both", expand=True, pady=10)
        self.ppt_body_text.bind("<KeyRelease>", self.ppt_save_current_fields)

        self.ppt_refresh_list()
        self.ppt_load_slide(0)

    def _default_slide(self, title, body):
        return {
            "title": title,
            "body": body,
            "bg": "#FFFFFF",
            "title_fg": "#B7472A",
            "body_fg": "#323130",
            "font": "Segoe UI",
            "size": 16,
            "bold": False,
            "italic": False,
            "underline": False,
            "overstrike": False,
            "align": "left"
        }

    def new_ppt_doc(self):
        self.switch_app("powerpoint")
        self.slides = [self._default_slide("Nouvelle Présentation", "Ajoutez vos idées ici...")]
        self.ppt_refresh_list()
        self.ppt_load_slide(0)

    def ppt_refresh_list(self):
        self.slide_listbox.delete(0, tk.END)
        for i, s in enumerate(self.slides):
            self.slide_listbox.insert(tk.END, f"Diapo {i+1} : {s['title'][:18]}")
        self.slide_listbox.selection_set(self.current_slide_idx)

    def ppt_load_slide(self, idx):
        self.current_slide_idx = idx
        s = self.slides[idx]

        # Appliquer couleurs et thème de la diapo
        self.slide_editor.config(bg=s["bg"])
        self.ppt_title_entry.config(bg=s["bg"], fg=s["title_fg"], justify=s["align"])
        self.ppt_body_text.config(bg=s["bg"], fg=s["body_fg"])

        self.ppt_title_entry.delete(0, tk.END)
        self.ppt_title_entry.insert(0, s["title"])

        self.ppt_body_text.delete("1.0", tk.END)
        self.ppt_body_text.insert("1.0", s["body"])

        self.ppt_font_var.set(s["font"])
        self.ppt_scale_var.set(str(s["size"]))
        self.ppt_apply_visual_styles()

    def ppt_apply_visual_styles(self):
        s = self.slides[self.current_slide_idx]
        styles = []
        if s["bold"]:
            styles.append("bold")
        if s["italic"]:
            styles.append("italic")
        if s["underline"]:
            styles.append("underline")
        if s["overstrike"]:
            styles.append("overstrike")

        font_tuple = (s["font"], s["size"], " ".join(styles)) if styles else (s["font"], s["size"])
        self.ppt_body_text.config(font=font_tuple, fg=s["body_fg"], bg=s["bg"])

        self.ppt_body_text.tag_configure("align_tag", justify=s["align"])
        self.ppt_body_text.tag_add("align_tag", "1.0", "end")
        self.ppt_title_entry.config(justify=s["align"], fg=s["title_fg"], bg=s["bg"])

    def ppt_scale_text(self, delta):
        s = self.slides[self.current_slide_idx]
        s["size"] = max(10, min(60, s["size"] + delta))
        self.ppt_scale_var.set(str(s["size"]))
        self.ppt_apply_visual_styles()

    def ppt_update_text_style(self):
        s = self.slides[self.current_slide_idx]
        s["font"] = self.ppt_font_var.get()
        s["size"] = int(self.ppt_scale_var.get())
        self.ppt_apply_visual_styles()

    def ppt_toggle_style(self, prop):
        s = self.slides[self.current_slide_idx]
        s[prop] = not s[prop]
        self.ppt_apply_visual_styles()

    def ppt_set_align(self, align):
        s = self.slides[self.current_slide_idx]
        s["align"] = align
        self.ppt_apply_visual_styles()

    def ppt_text_color(self):
        color = colorchooser.askcolor(title="Couleur du texte de la diapositive")[1]
        if color:
            try:
                tag = f"ppt_fg_{color}"
                self.ppt_body_text.tag_configure(tag, foreground=color)
                self.ppt_body_text.tag_add(tag, "sel.first", "sel.last")
            except tk.TclError:
                self.slides[self.current_slide_idx]["body_fg"] = color
                self.ppt_apply_visual_styles()

    def ppt_highlight_text(self):
        color = colorchooser.askcolor(title="Couleur de surlignage")[1]
        if color:
            try:
                tag = f"ppt_bg_{color}"
                self.ppt_body_text.tag_configure(tag, background=color)
                self.ppt_body_text.tag_add(tag, "sel.first", "sel.last")
            except tk.TclError:
                pass

    def ppt_apply_theme(self, bg, title_fg, body_fg):
        s = self.slides[self.current_slide_idx]
        s["bg"] = bg
        s["title_fg"] = title_fg
        s["body_fg"] = body_fg
        self.ppt_apply_visual_styles()

    def ppt_custom_bg(self):
        color = colorchooser.askcolor(title="Couleur de fond de la diapositive")[1]
        if color:
            self.slides[self.current_slide_idx]["bg"] = color
            self.ppt_apply_visual_styles()

    def ppt_save_current_fields(self, event=None):
        self.slides[self.current_slide_idx]["title"] = self.ppt_title_entry.get()
        self.slides[self.current_slide_idx]["body"] = self.ppt_body_text.get("1.0", "end-1c")
        self.ppt_refresh_list()

    def ppt_select_slide(self, event):
        sel = self.slide_listbox.curselection()
        if sel:
            self.ppt_load_slide(sel[0])

    def ppt_add_slide(self):
        self.slides.append(self._default_slide(f"Diapositive {len(self.slides)+1}", "• Point clé 1\n• Point clé 2"))
        self.ppt_refresh_list()
        self.ppt_load_slide(len(self.slides) - 1)

    def ppt_duplicate_slide(self):
        copy_slide = dict(self.slides[self.current_slide_idx])
        copy_slide["title"] += " (Copie)"
        self.slides.insert(self.current_slide_idx + 1, copy_slide)
        self.ppt_refresh_list()
        self.ppt_load_slide(self.current_slide_idx + 1)

    def ppt_move_slide(self, direction):
        idx = self.current_slide_idx
        new_idx = idx + direction
        if 0 <= new_idx < len(self.slides):
            self.slides[idx], self.slides[new_idx] = self.slides[new_idx], self.slides[idx]
            self.current_slide_idx = new_idx
            self.ppt_refresh_list()
            self.ppt_load_slide(new_idx)

    def ppt_del_slide(self):
        if len(self.slides) > 1:
            self.slides.pop(self.current_slide_idx)
            self.current_slide_idx = max(0, self.current_slide_idx - 1)
            self.ppt_refresh_list()
            self.ppt_load_slide(self.current_slide_idx)

    def ppt_start_slideshow(self):
        show = tk.Toplevel(self.root)
        show.attributes("-fullscreen", True)
        idx = [self.current_slide_idx]

        lbl_t = tk.Label(show, text="", font=("Segoe UI", 44, "bold"))
        lbl_t.pack(pady=(110, 30), fill="x", padx=80)

        lbl_b = tk.Label(show, text="", wraplength=1100)
        lbl_b.pack(pady=20, padx=100, fill="x")

        lbl_hint = tk.Label(show, text="Flèches Gauche/Droite ou Espace pour avancer • Échap pour quitter", font=("Segoe UI", 10))
        lbl_hint.pack(side="bottom", pady=20)

        def render():
            s = self.slides[idx[0]]
            show.configure(bg=s["bg"])
            styles = []
            if s["bold"]:
                styles.append("bold")
            if s["italic"]:
                styles.append("italic")
            if s["underline"]:
                styles.append("underline")
            font_tuple = (s["font"], s["size"] + 8, " ".join(styles)) if styles else (s["font"], s["size"] + 8)

            anchor_map = {"left": "w", "center": "center", "right": "e"}
            lbl_t.config(text=s["title"], bg=s["bg"], fg=s["title_fg"], anchor=anchor_map[s["align"]])
            lbl_b.config(text=s["body"], bg=s["bg"], fg=s["body_fg"], font=font_tuple, justify=s["align"], anchor=anchor_map[s["align"]])
            lbl_hint.config(bg=s["bg"], fg="#888888")

        def next_s(e=None):
            if idx[0] < len(self.slides) - 1:
                idx[0] += 1
                render()

        def prev_s(e=None):
            if idx[0] > 0:
                idx[0] -= 1
                render()

        show.bind("<Right>", next_s)
        show.bind("<space>", next_s)
        show.bind("<Left>", prev_s)
        show.bind("<Escape>", lambda e: show.destroy())
        render()

    def save_ppt_doc(self):
        path = filedialog.asksaveasfilename(defaultextension=".omnippt", filetypes=[("Présentation OmniPowerPoint", "*.omnippt"), ("Fichier Texte", "*.txt")])
        if not path:
            return
        with open(path, "w", encoding="utf-8") as f:
            json.dump({"type": "omnippt", "slides": self.slides}, f, ensure_ascii=False, indent=2)
        self._add_recent(path)
        messagebox.showinfo("Enregistré", f"Présentation sauvegardée :\n{path}")

    # =========================================================================
    # 5. OUVERTURE UNIVERSELLE (TXT, CSV, ODT, ODS, ODP, PY, JSON...)
    # =========================================================================
    def open_any_file(self):
        path = filedialog.askopenfilename(
            title="Ouvrir un fichier dans OmniOffice 365 Pro",
            filetypes=[
                ("Tous les fichiers supportés", "*.txt *.csv *.odt *.ods *.odp *.omniword *.omnippt *.md *.py *.json *.log"),
                ("Documents LibreOffice / OpenDocument", "*.odt *.ods *.odp"),
                ("Fichiers Texte", "*.txt"),
                ("Tableurs CSV", "*.csv"),
                ("Tous les fichiers", "*.*")
            ]
        )
        if path:
            self.load_file_by_path(path)

    def load_file_by_path(self, path):
        ext = os.path.splitext(path)[1].lower()
        self.current_file_path = path
        self.lbl_doc_name.config(text=f"— {os.path.basename(path)}")
        self._add_recent(path)

        try:
            # 1. LibreOffice Writer (.odt)
            if ext == ".odt":
                text_content = self._read_odf_text(path)
                self.switch_app("word")
                self.word_text.delete("1.0", tk.END)
                self.word_text.insert("1.0", text_content)
                self.word_update_stats()

            # 2. LibreOffice Calc (.ods)
            elif ext == ".ods":
                rows_data = self._read_ods_file(path)
                self.new_excel_doc()
                for r, row in enumerate(rows_data):
                    if r >= self.excel_rows:
                        break
                    for c, val in enumerate(row):
                        if c >= self.excel_cols:
                            break
                        self.excel_raw_data[(r, c)] = val
                        self.excel_cells[(r, c)].insert(0, val)
                self.excel_recalculate_all()

            # 3. LibreOffice Impress (.odp)
            elif ext == ".odp":
                text_content = self._read_odf_text(path)
                self.switch_app("powerpoint")
                self.slides = [self._default_slide(os.path.basename(path), text_content)]
                self.ppt_refresh_list()
                self.ppt_load_slide(0)

            # 4. Tableur CSV
            elif ext == ".csv":
                self.new_excel_doc()
                with open(path, "r", encoding="utf-8", errors="replace") as f:
                    reader = csv.reader(f, delimiter=";")
                    for r, row in enumerate(reader):
                        if r >= self.excel_rows:
                            break
                        for c, val in enumerate(row):
                            if c >= self.excel_cols:
                                break
                            self.excel_raw_data[(r, c)] = val
                            self.excel_cells[(r, c)].insert(0, val)
                self.excel_recalculate_all()

            # 5. OmniPowerPoint
            elif ext == ".omnippt":
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                self.switch_app("powerpoint")
                loaded = data.get("slides", [])
                self.slides = [self._default_slide(s.get("title", ""), s.get("body", "")) | s for s in loaded]
                self.ppt_refresh_list()
                self.ppt_load_slide(0)

            # 6. Tous les autres fichiers (.txt, .py, .md, .json, .log, etc.)
            else:
                with open(path, "r", encoding="utf-8", errors="replace") as f:
                    content = f.read()
                self.switch_app("word")
                self.word_text.delete("1.0", tk.END)
                self.word_text.insert("1.0", content)
                self.word_update_stats()

        except Exception as e:
            messagebox.showerror("Erreur d'ouverture", f"Impossible de lire ce fichier :\n{e}")

    def _read_odf_text(self, path):
        """Extrait le texte d'un fichier LibreOffice (.odt ou .odp)"""
        texts = []
        try:
            with zipfile.ZipFile(path, 'r') as z:
                if 'content.xml' in z.namelist():
                    root_xml = ET.fromstring(z.read('content.xml'))
                    for elem in root_xml.iter():
                        if elem.tag.endswith('p'):
                            full_p = "".join(elem.itertext())
                            if full_p.strip():
                                texts.append(full_p)
        except Exception:
            pass
        return "\n".join(texts) if texts else "[Document OpenDocument vide]"

    def _read_ods_file(self, path):
        """Extrait les cellules d'un classeur LibreOffice Calc (.ods)"""
        cells_data = []
        try:
            with zipfile.ZipFile(path, 'r') as z:
                if 'content.xml' in z.namelist():
                    root_xml = ET.fromstring(z.read('content.xml'))
                    for table in root_xml.iter('{urn:oasis:names:tc:opendocument:xmlns:table:1.0}table'):
                        for row in table.iter('{urn:oasis:names:tc:opendocument:xmlns:table:1.0}table-row'):
                            row_vals = []
                            for cell in row.iter('{urn:oasis:names:tc:opendocument:xmlns:table:1.0}table-cell'):
                                cell_text = "".join(cell.itertext()).strip()
                                row_vals.append(cell_text)
                            if any(row_vals):
                                cells_data.append(row_vals)
        except Exception:
            pass
        return cells_data


if __name__ == "__main__":
    root = tk.Tk()
    app = OmniOffice365(root)
    root.mainloop()