"""
GUI fuer den MYD Migration Checker.
XML-Dateien + bis zu 10 TXT-Protokolle laden, vergleichen, Excel erstellen.
"""
import tkinter as tk
from tkinter import filedialog, messagebox
import os

from config.theme import (
    ORANGE, ORANGE_LIGHT, DARK, BG, WHITE, GREY_LIGHT, GREY,
    RED, GREEN, YELLOW, MAX_FILES, APP_TITLE, WINDOW_SIZE
)
from core.xml_parser import extract_key_data
from core.txt_parser import parse_migration_protocol, analyze_protocol
from core.excel_writer import create_excel

MAX_PROTOCOLS = 10

class XMLtoExcelApp:

    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.configure(bg=BG)
        self.root.resizable(True, True)

        try:
            icon_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'key_icon.ico')
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except Exception:
            pass

        self.files = []
        self.protocol_files = []
        self._build_ui()

    def _build_ui(self):
        tk.Frame(self.root, bg=ORANGE, height=6).pack(fill='x')

        hf = tk.Frame(self.root, bg=DARK, pady=10)
        hf.pack(fill='x')
        tk.Label(hf, text="SAP Migration Checker", font=('Arial', 18, 'bold'),
                 fg=WHITE, bg=DARK).pack()
        tk.Label(hf, text="XML-Dateien + Migrations-Protokolle vergleichen",
                 font=('Arial', 9), fg=GREY_LIGHT, bg=DARK).pack(pady=(2, 0))

        # 1. XML Dateien
        sf = tk.Frame(self.root, bg=BG)
        sf.pack(padx=20, pady=(12, 3), fill='x')
        tk.Label(sf, text="1  XML-Dateien", font=('Arial', 11, 'bold'),
                 fg=DARK, bg=BG, anchor='w').pack(fill='x')

        self.drop_label = tk.Label(sf, text="Klicken um XML-Dateien auszuwaehlen",
                                   font=('Arial', 11), fg=ORANGE, bg=WHITE,
                                   cursor='hand2', pady=14, relief='solid', bd=1)
        self.drop_label.pack(fill='x', pady=(3, 0))
        self.drop_label.bind('<Button-1>', self._select_files)
        self.drop_label.bind('<Enter>', lambda e: self.drop_label.config(bg=GREY_LIGHT))
        self.drop_label.bind('<Leave>', lambda e: self.drop_label.config(bg=WHITE))

        lf = tk.Frame(self.root, bg=BG)
        lf.pack(padx=20, pady=3, fill='both', expand=True)
        self.file_count_label = tk.Label(lf, text=f"Dateien: 0 / {MAX_FILES}",
                                         font=('Arial', 9, 'bold'), fg=DARK, bg=BG, anchor='w')
        self.file_count_label.pack(fill='x')
        scf = tk.Frame(lf, bg=WHITE, bd=1, relief='solid')
        scf.pack(fill='both', expand=True, pady=3)
        self.scrollbar = tk.Scrollbar(scf)
        self.scrollbar.pack(side='right', fill='y')
        self.file_listbox = tk.Listbox(scf, font=('Consolas', 9), bg=WHITE, fg=DARK,
                                        selectbackground=ORANGE, selectforeground=WHITE,
                                        yscrollcommand=self.scrollbar.set, height=6,
                                        bd=0, highlightthickness=0)
        self.file_listbox.pack(fill='both', expand=True)
        self.scrollbar.config(command=self.file_listbox.yview)

        bf = tk.Frame(self.root, bg=BG)
        bf.pack(padx=20, pady=3, fill='x')
        tk.Button(bf, text="Entfernen", font=('Arial', 9),
                  bg=RED, fg=WHITE, command=self._remove_selected,
                  cursor='hand2', bd=0, padx=8, pady=3).pack(side='left', padx=(0, 5))
        tk.Button(bf, text="Alle entfernen", font=('Arial', 9),
                  bg=GREY, fg=WHITE, command=self._clear_all,
                  cursor='hand2', bd=0, padx=8, pady=3).pack(side='left')

        # 2. Protokoll-Dateien (bis zu 10)
        pf = tk.Frame(self.root, bg=BG)
        pf.pack(padx=20, pady=(10, 3), fill='x')
        tk.Label(pf, text=f"2  Migrations-Protokolle (bis zu {MAX_PROTOCOLS}, optional)",
                 font=('Arial', 11, 'bold'), fg=DARK, bg=BG, anchor='w').pack(fill='x')

        self.proto_drop = tk.Label(pf, text="Klicken um TXT-Protokolle zu laden",
                                   font=('Arial', 11), fg=ORANGE, bg=WHITE,
                                   cursor='hand2', pady=10, relief='solid', bd=1)
        self.proto_drop.pack(fill='x', pady=(3, 0))
        self.proto_drop.bind('<Button-1>', self._select_protocols)
        self.proto_drop.bind('<Enter>', lambda e: self.proto_drop.config(bg=GREY_LIGHT))
        self.proto_drop.bind('<Leave>', lambda e: self.proto_drop.config(bg=WHITE))

        plf = tk.Frame(self.root, bg=BG)
        plf.pack(padx=20, pady=3, fill='x')
        self.proto_count_label = tk.Label(plf, text=f"Protokolle: 0 / {MAX_PROTOCOLS}",
                                          font=('Arial', 9, 'bold'), fg=DARK, bg=BG, anchor='w')
        self.proto_count_label.pack(fill='x')
        pscf = tk.Frame(plf, bg=WHITE, bd=1, relief='solid')
        pscf.pack(fill='x', pady=3)
        self.proto_scrollbar = tk.Scrollbar(pscf)
        self.proto_scrollbar.pack(side='right', fill='y')
        self.proto_listbox = tk.Listbox(pscf, font=('Consolas', 9), bg=WHITE, fg=DARK,
                                         selectbackground=ORANGE, selectforeground=WHITE,
                                         yscrollcommand=self.proto_scrollbar.set,
                                         height=4, bd=0, highlightthickness=0)
        self.proto_listbox.pack(fill='both', expand=True)
        self.proto_scrollbar.config(command=self.proto_listbox.yview)

        pbf = tk.Frame(self.root, bg=BG)
        pbf.pack(padx=20, pady=3, fill='x')
        tk.Button(pbf, text="Entfernen", font=('Arial', 9),
                  bg=RED, fg=WHITE, command=self._remove_protocol,
                  cursor='hand2', bd=0, padx=8, pady=3).pack(side='left', padx=(0, 5))
        tk.Button(pbf, text="Alle entfernen", font=('Arial', 9),
                  bg=GREY, fg=WHITE, command=self._clear_protocols,
                  cursor='hand2', bd=0, padx=8, pady=3).pack(side='left')

        self.proto_info = tk.Label(self.root, text="", font=('Arial', 9), fg=GREY, bg=BG, anchor='w')
        self.proto_info.pack(padx=20, fill='x')

        # Export
        tk.Button(self.root, text="Excel erstellen", font=('Arial', 13, 'bold'),
                  bg=ORANGE, fg=WHITE, activebackground=ORANGE_LIGHT,
                  command=self._export, cursor='hand2', bd=0, pady=8
                  ).pack(padx=20, pady=(10, 4), fill='x')

        self.status_label = tk.Label(self.root, text="", font=('Arial', 9),
                                      fg=GREY, bg=BG, wraplength=720)
        self.status_label.pack(pady=(0, 5))
        tk.Frame(self.root, bg=ORANGE, height=4).pack(fill='x', side='bottom')

    def _select_files(self, event=None):
        fps = filedialog.askopenfilenames(
            title="XML-Dateien auswaehlen",
            filetypes=[("XML-Dateien", "*.xml"), ("Alle Dateien", "*.*")]
        )
        for fp in fps:
            if len(self.files) >= MAX_FILES:
                messagebox.showwarning("Maximum", f"Maximal {MAX_FILES} Dateien!")
                break
            if fp not in self.files:
                self.files.append(fp)
        self._update_xml_list()

    def _update_xml_list(self):
        self.file_listbox.delete(0, tk.END)
        for i, fp in enumerate(self.files):
            self.file_listbox.insert(tk.END, f"  {i+1}.  {os.path.basename(fp)}")
        self.file_count_label.config(text=f"Dateien: {len(self.files)} / {MAX_FILES}")

    def _remove_selected(self):
        for idx in reversed(self.file_listbox.curselection()):
            self.files.pop(idx)
        self._update_xml_list()

    def _clear_all(self):
        self.files.clear()
        self._update_xml_list()
        self.protocol_files.clear()
        self._update_proto_list()
        self.status_label.config(text="", fg=GREY)

    def _select_protocols(self, event=None):
        fps = filedialog.askopenfilenames(
            title="Migrations-Protokolle auswaehlen",
            filetypes=[("Text-Dateien", "*.txt"), ("CSV-Dateien", "*.csv"), ("Alle Dateien", "*.*")]
        )
        for fp in fps:
            if len(self.protocol_files) >= MAX_PROTOCOLS:
                messagebox.showwarning("Maximum", f"Maximal {MAX_PROTOCOLS} Protokolle!")
                break
            if fp not in self.protocol_files:
                self.protocol_files.append(fp)
        self._update_proto_list()

    def _update_proto_list(self):
        self.proto_listbox.delete(0, tk.END)
        total = 0
        for i, fp in enumerate(self.protocol_files):
            fname = os.path.basename(fp)
            try:
                headers, data, _ = parse_migration_protocol(fp)
                stats = analyze_protocol(data)
                total += stats['total']
                status = "OK" if stats['all_migrated'] and stats['all_success'] else "PRUEFEN"
                self.proto_listbox.insert(tk.END, f"  {i+1}.  {fname}  ({stats['total']} Saetze, {status})")
            except Exception:
                self.proto_listbox.insert(tk.END, f"  {i+1}.  {fname}  (FEHLER)")
        self.proto_count_label.config(text=f"Protokolle: {len(self.protocol_files)} / {MAX_PROTOCOLS}")
        if self.protocol_files:
            self.proto_info.config(text=f"  Gesamt: {total} Datensaetze in {len(self.protocol_files)} Protokoll(en)", fg=GREEN)
        else:
            self.proto_info.config(text="", fg=GREY)

    def _remove_protocol(self):
        for idx in reversed(self.proto_listbox.curselection()):
            self.protocol_files.pop(idx)
        self._update_proto_list()

    def _clear_protocols(self):
        self.protocol_files.clear()
        self._update_proto_list()

    def _gen_path(self):
        first = self.files[0]
        folder = os.path.dirname(first)
        base = os.path.splitext(os.path.basename(first))[0]
        out = os.path.join(folder, f"{base}.xlsx")
        c = 1
        while os.path.exists(out):
            out = os.path.join(folder, f"{base}_{c}.xlsx")
            c += 1
        return out

    def _export(self):
        if not self.files:
            messagebox.showwarning("Keine Dateien", "Bitte zuerst XML-Dateien hinzufuegen!")
            return

        output_path = self._gen_path()
        self.status_label.config(text="Verarbeite...", fg=ORANGE)
        self.root.update()

        try:
            all_data = {}
            skipped = []

            for fp in self.files:
                fname = os.path.basename(fp)
                self.status_label.config(text=f"XML: {fname}")
                self.root.update()
                try:
                    key_sig, key_fields, data_rows = extract_key_data(fp)
                    for row in data_rows:
                        row['_filename'] = fname
                    if key_sig not in all_data:
                        all_data[key_sig] = {'fields': key_fields, 'rows': []}
                    all_data[key_sig]['rows'].extend(data_rows)
                except Exception as e:
                    skipped.append((fname, str(e)))

            if not all_data and not skipped:
                raise ValueError("Keine Daten konnten extrahiert werden!")

            # Alle Key-Felder aus XML sammeln
            all_key_fields = []
            for info in all_data.values():
                for f in info['fields']:
                    if f not in all_key_fields:
                        all_key_fields.append(f)

            # XML-Daten mit _match_key anreichern
            for info in all_data.values():
                for row in info['rows']:
                    key_parts = [row.get(f, '').strip() for f in info['fields']]
                    row['_match_key'] = '|'.join(key_parts)

            all_protocols = []
            for fp in self.protocol_files:
                fname = os.path.basename(fp)
                self.status_label.config(text=f"Protokoll: {fname}")
                self.root.update()
                try:
                    headers, data, proto_type = parse_migration_protocol(fp, xml_key_fields=all_key_fields)
                    stats = analyze_protocol(data)
                    all_protocols.append({
                        'filename': fname,
                        'headers': headers,
                        'data': data,
                        'stats': stats,
                        'proto_type': proto_type,
                    })
                except Exception as e:
                    skipped.append((fname, f"Protokoll-Fehler: {str(e)}"))

            self.status_label.config(text="Excel erstellen...")
            self.root.update()

            create_excel(all_data, skipped, output_path, all_protocols=all_protocols)

            total = sum(len(i['rows']) for i in all_data.values())
            out_name = os.path.basename(output_path)

            sheets_info = f"{len(all_data)} Key-Sheet(s)"
            if all_protocols:
                sheets_info += f" + {len(all_protocols)} Protokoll(e) + Vergleich(e)"

            msg = f"Fertig! {total} Datensaetze -> {out_name}"
            if skipped:
                msg += f" | {len(skipped)} fehlerhaft"
                self.status_label.config(text=msg, fg=YELLOW)
            else:
                self.status_label.config(text=msg, fg=GREEN)

            info = (f"Excel erstellt!\n\n"
                    f"XML-Dateien: {len(self.files) - len([x for x in skipped if not x[1].startswith('Protokoll')])}/{len(self.files)}\n"
                    f"Datensaetze: {total}\n"
                    f"Sheets: {sheets_info}\n")

            if all_protocols:
                info += f"\nProtokolle: {len(all_protocols)}\n"
                for p in all_protocols:
                    st = p['stats']
                    status = "ALLE OK" if st['all_migrated'] and st['all_success'] else "PRUEFEN"
                    info += f"  - {p['filename']}: {st['total']} Saetze ({status})\n"

            info += f"\nSpeicherort: {output_path}"
            messagebox.showinfo("Erfolg!", info)

        except Exception as e:
            self.status_label.config(text=f"Fehler: {str(e)}", fg=RED)
            messagebox.showerror("Fehler", str(e))
