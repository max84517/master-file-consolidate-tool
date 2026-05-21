"""
Main UI — CustomTkinter dark mode.

Layout:
  ┌─ Source Folders ─────────────────────────────────────────────┐
  │  NB KB Path        [entry]  [Browse]                         │
  │  DT KB Path        [entry]  [Browse]                         │
  │  Peripheral Path   [entry]  [Browse]                         │
  │                    [Ingest Files]                            │
  ├─ Output Paths ───────────────────────────────────────────────┤
  │  NB Output         [entry]  [Browse]                         │
  │  DT Output         [entry]  [Browse]                         │
  │  Peripheral Output [entry]  [Browse]                         │
  │  Consolidate All   [entry]  [Browse]                         │
  ├─ Consolidation Options ──────────────────────────────────────┤
  │  FY Year           [dropdown]                                │
  │  Value Columns     [✓ HP Cost] [✓ ODM Cost] [✓ Rebate]       │
  ├─ Actions ────────────────────────────────────────────────────┤
  │  [Consolidate by Segment]   [Consolidate All]                │
  ├─ Status Log ─────────────────────────────────────────────────┤
  │  (scrollable textbox)                                        │
  └──────────────────────────────────────────────────────────────┘
"""

from __future__ import annotations

import sys
import threading
from pathlib import Path

import customtkinter as ctk
from tkinter import filedialog

# Dark mode — must be set before any CTk widget
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class App(ctk.CTk):
    def __init__(self, cfg: dict, save_cfg_fn):
        super().__init__()
        self.cfg = cfg
        self.save_cfg_fn = save_cfg_fn

        self.title("Master File Consolidation Tool")
        self.geometry("800x560")
        self.resizable(True, True)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        self._build_ui()
        self._refresh_fy_dropdown()

    # ── close ────────────────────────────────────────────────────────────────

    def _on_close(self):
        self._save_cfg()
        self.destroy()
        sys.exit(0)

    def _save_cfg(self):
        self.cfg["nb_kb_path"] = self.var_nb.get()
        self.cfg["dt_kb_path"] = self.var_dt.get()
        self.cfg["peripheral_path"] = self.var_periph.get()
        self.cfg["output_nb_path"] = self.var_out_nb.get()
        self.cfg["output_dt_path"] = self.var_out_dt.get()
        self.cfg["output_peripheral_path"] = self.var_out_periph.get()
        self.cfg["output_consolidate_all_path"] = self.var_out_all.get()
        self.save_cfg_fn(self.cfg)

    # ── UI construction ──────────────────────────────────────────────────────

    def _build_ui(self):
        LABEL_W = 120
        ENTRY_W = 390
        SEC_FONT = ctk.CTkFont(size=11, weight="bold")

        # ── Source Folders (title + Ingest on same row) ──
        src_frame = ctk.CTkFrame(self)
        src_frame.pack(fill="x", padx=10, pady=(8, 2))

        src_hdr = ctk.CTkFrame(src_frame, fg_color="transparent")
        src_hdr.pack(fill="x", padx=8, pady=(4, 2))
        ctk.CTkLabel(src_hdr, text="Source Folders", font=SEC_FONT).pack(side="left")
        ctk.CTkButton(src_hdr, text="Ingest Files", height=24, width=90,
                      command=self._run_ingest).pack(side="right")

        self.var_nb = ctk.StringVar(value=self.cfg.get("nb_kb_path", ""))
        self.var_dt = ctk.StringVar(value=self.cfg.get("dt_kb_path", ""))
        self.var_periph = ctk.StringVar(value=self.cfg.get("peripheral_path", ""))

        self._path_row(src_frame, "NB KB:", self.var_nb, LABEL_W, ENTRY_W)
        self._path_row(src_frame, "DT KB:", self.var_dt, LABEL_W, ENTRY_W)
        self._path_row(src_frame, "Peripheral:", self.var_periph, LABEL_W, ENTRY_W)

        # ──  Output Paths (2×2 layout) ──
        out_frame = ctk.CTkFrame(self)
        out_frame.pack(fill="x", padx=10, pady=2)
        ctk.CTkLabel(out_frame, text="Output Paths", font=SEC_FONT).pack(anchor="w", padx=8, pady=(4, 2))

        from master_consolidate.config.paths import RESULT_BY_SEGMENT_DIR, CONSOLIDATE_ALL_DIR

        self.var_out_nb = ctk.StringVar(value=self.cfg.get("output_nb_path") or str(RESULT_BY_SEGMENT_DIR / "NB"))
        self.var_out_dt = ctk.StringVar(value=self.cfg.get("output_dt_path") or str(RESULT_BY_SEGMENT_DIR / "DT"))
        self.var_out_periph = ctk.StringVar(value=self.cfg.get("output_peripheral_path") or str(RESULT_BY_SEGMENT_DIR / "Peripheral"))
        self.var_out_all = ctk.StringVar(value=self.cfg.get("output_consolidate_all_path") or str(CONSOLIDATE_ALL_DIR))

        out_r1 = ctk.CTkFrame(out_frame, fg_color="transparent")
        out_r1.pack(fill="x", padx=8, pady=1)
        self._inline_path(out_r1, "NB:", self.var_out_nb)
        self._inline_path(out_r1, "DT:", self.var_out_dt)

        out_r2 = ctk.CTkFrame(out_frame, fg_color="transparent")
        out_r2.pack(fill="x", padx=8, pady=(1, 4))
        self._inline_path(out_r2, "Peripheral:", self.var_out_periph)
        self._inline_path(out_r2, "Consol. All:", self.var_out_all)

        # ── Options (FY + checkboxes on one row) ──
        opt_frame = ctk.CTkFrame(self)
        opt_frame.pack(fill="x", padx=10, pady=2)

        opt_row = ctk.CTkFrame(opt_frame, fg_color="transparent")
        opt_row.pack(fill="x", padx=8, pady=(5, 2))
        ctk.CTkLabel(opt_row, text="FY:", width=30, anchor="w").pack(side="left")
        self.fy_var = ctk.StringVar(value="")
        self.fy_dropdown = ctk.CTkOptionMenu(
            opt_row, variable=self.fy_var, values=["(no data yet)"],
            command=self._on_fy_changed, width=90,
        )
        self.fy_dropdown.pack(side="left", padx=(2, 16))
        ctk.CTkLabel(opt_row, text="Value Columns:", anchor="w").pack(side="left")
        self.chk_hp = ctk.CTkCheckBox(opt_row, text="HP Cost", width=85)
        self.chk_hp.select()
        self.chk_hp.pack(side="left", padx=(8, 4))
        self.chk_odm = ctk.CTkCheckBox(opt_row, text="ODM Cost", width=90)
        self.chk_odm.select()
        self.chk_odm.pack(side="left", padx=4)
        self.chk_rebate = ctk.CTkCheckBox(opt_row, text="Rebate", width=75)
        self.chk_rebate.select()
        self.chk_rebate.pack(side="left", padx=4)

        # Coverage banner
        cov_row = ctk.CTkFrame(opt_frame, fg_color="transparent")
        cov_row.pack(fill="x", padx=8, pady=(0, 4))
        self.coverage_label = ctk.CTkLabel(
            cov_row, text="", anchor="w", justify="left",
            wraplength=720, font=ctk.CTkFont(size=11),
        )
        self.coverage_label.pack(anchor="w")

        # ── Actions ──
        act_frame = ctk.CTkFrame(self)
        act_frame.pack(fill="x", padx=10, pady=2)
        ctk.CTkButton(act_frame, text="Consolidate by Segment", height=28,
                      command=self._run_by_segment).pack(side="left", padx=8, pady=5)
        ctk.CTkButton(act_frame, text="Consolidate All", height=28,
                      command=self._run_consolidate_all).pack(side="left", padx=4, pady=5)

        # ── Status log ──
        log_frame = ctk.CTkFrame(self)
        log_frame.pack(fill="both", expand=True, padx=10, pady=(2, 8))
        ctk.CTkLabel(log_frame, text="Log", font=SEC_FONT).pack(anchor="w", padx=8, pady=(3, 1))
        self.log_box = ctk.CTkTextbox(log_frame, state="disabled", wrap="word", height=90)
        self.log_box.pack(fill="both", expand=True, padx=8, pady=(0, 5))

    def _path_row(self, parent, label: str, var: ctk.StringVar, label_w: int, entry_w: int):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", padx=8, pady=1)
        ctk.CTkLabel(row, text=label, width=label_w, anchor="w").pack(side="left")
        ctk.CTkEntry(row, textvariable=var, width=entry_w, height=26).pack(side="left", padx=4)
        ctk.CTkButton(row, text="Browse", width=60, height=26,
                      command=lambda v=var: self._browse_folder(v)).pack(side="left", padx=2)

    def _inline_path(self, parent, label: str, var: ctk.StringVar):
        """Compact path input used in the 2-column output section."""
        ctk.CTkLabel(parent, text=label, width=75, anchor="w").pack(side="left")
        ctk.CTkEntry(parent, textvariable=var, width=248, height=26).pack(side="left", padx=2)
        ctk.CTkButton(parent, text="...", width=28, height=26,
                      command=lambda v=var: self._browse_folder(v)).pack(side="left", padx=(0, 14))

    def _browse_folder(self, var: ctk.StringVar):
        path = filedialog.askdirectory(initialdir=var.get() or "/")
        if path:
            var.set(path)

    # ── FY dropdown refresh ──────────────────────────────────────────────────

    def _refresh_fy_dropdown(self):
        from master_consolidate.config.paths import SOURCE_DIR
        from master_consolidate.ingestion.ingest import available_fy_years
        try:
            years = available_fy_years(SOURCE_DIR)
        except Exception:
            years = []
        if years:
            self.fy_dropdown.configure(values=years)
            self.fy_var.set(years[-1])
            self._on_fy_changed(years[-1])
        else:
            self.fy_dropdown.configure(values=["(no data yet)"])
            self.fy_var.set("(no data yet)")
            self.coverage_label.configure(text="", text_color="gray")

    def _on_fy_changed(self, fy: str):
        """Immediately show coverage status when user picks a different FY."""
        if not fy or fy == "(no data yet)":
            self.coverage_label.configure(text="", text_color="gray")
            return
        threading.Thread(target=self._coverage_worker, args=(fy,), daemon=True).start()

    def _coverage_worker(self, fy: str):
        from master_consolidate.config.paths import SOURCE_DIR
        from master_consolidate.ingestion.ingest import check_fy_coverage
        try:
            coverage = check_fy_coverage(SOURCE_DIR, fy)
        except Exception as exc:
            self.after(0, lambda: self.coverage_label.configure(
                text=f"Coverage check error: {exc}", text_color="orange"
            ))
            return

        missing_lines = []
        all_empty = True
        for seg, suppliers in coverage.items():
            if not suppliers:
                continue
            all_empty = False
            for supplier, has_sheet in suppliers.items():
                if not has_sheet:
                    missing_lines.append(f"  [{seg}] {supplier}")

        if all_empty:
            msg, color = "No ingested files found. Run Ingest first.", "gray"
        elif missing_lines:
            lines = "\n".join(missing_lines)
            msg = f"⚠ Missing {fy} sheet in:\n{lines}"
            color = "#E8A020"   # amber
        else:
            msg = f"✓ All suppliers have {fy} sheet."
            color = "#4CAF50"   # green

        self.after(0, lambda m=msg, c=color: self.coverage_label.configure(text=m, text_color=c))

    # ── logging ──────────────────────────────────────────────────────────────

    def _log(self, msg: str):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", msg + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def _log_clear(self):
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

    # ── ingest ───────────────────────────────────────────────────────────────

    def _run_ingest(self):
        self._save_cfg()
        self._log_clear()
        self._log("=== Ingesting files ===")
        threading.Thread(target=self._ingest_worker, daemon=True).start()

    def _ingest_worker(self):
        from master_consolidate.config.paths import SOURCE_DIR
        from master_consolidate.ingestion.ingest import ingest_segment

        segments = {
            "NB": self.var_nb.get(),
            "DT": self.var_dt.get(),
            "Peripheral": self.var_periph.get(),
        }

        for seg, src_path_str in segments.items():
            if not src_path_str:
                self._log(f"[{seg}] ⚠ No path set, skipping.")
                continue
            src_path = Path(src_path_str)
            dest = SOURCE_DIR / seg
            try:
                summary = ingest_segment(seg, src_path, dest)
                if not summary:
                    self._log(f"[{seg}] No supplier folders found.")
                for supplier, sheets in summary.items():
                    if sheets:
                        self._log(f"[{seg}] {supplier}: {', '.join(sheets)}")
                    else:
                        self._log(f"[{seg}] {supplier}: (no matching sheets found)")
            except Exception as exc:
                self._log(f"[{seg}] ERROR: {exc}")

        self._log("=== Ingestion complete ===")
        # Refresh FY dropdown after ingest
        self.after(0, self._refresh_fy_dropdown)

    # ── consolidate by segment ───────────────────────────────────────────────

    def _run_by_segment(self):
        self._save_cfg()
        fy = self.fy_var.get()
        if not fy or fy == "(no data yet)":
            self._log("Please ingest files first, then select a FY year.")
            return
        keywords = self._selected_keywords()
        if not keywords:
            self._log("Please select at least one Value Column.")
            return
        self._log_clear()
        self._log(f"=== Consolidate by Segment | {fy} | Columns: {', '.join(keywords)} ===")
        threading.Thread(
            target=self._by_segment_worker,
            args=(fy, keywords),
            daemon=True,
        ).start()

    def _by_segment_worker(self, fy: str, keywords: list[str]) -> dict[str, Path]:
        from master_consolidate.config.paths import SOURCE_DIR
        from master_consolidate.consolidation.consolidate import consolidate_segment

        out_paths = {
            "NB": Path(self.var_out_nb.get()),
            "DT": Path(self.var_out_dt.get()),
            "Peripheral": Path(self.var_out_periph.get()),
        }
        results: dict[str, Path] = {}
        for seg in ("NB", "DT", "Peripheral"):
            source_dir = SOURCE_DIR / seg
            try:
                out_file = consolidate_segment(
                    seg, source_dir, fy, keywords, out_paths[seg]
                )
                results[seg] = out_file
                self._log(f"[{seg}] Saved → {out_file}")
            except Exception as exc:
                self._log(f"[{seg}] ERROR: {exc}")
        self._log("=== Consolidation by Segment complete ===")
        return results

    # ── consolidate all ──────────────────────────────────────────────────────

    def _run_consolidate_all(self):
        self._save_cfg()
        fy = self.fy_var.get()
        if not fy or fy == "(no data yet)":
            self._log("Please ingest files first, then select a FY year.")
            return
        keywords = self._selected_keywords()
        if not keywords:
            self._log("Please select at least one Value Column.")
            return
        self._log_clear()
        self._log(f"=== Consolidate All | {fy} | Columns: {', '.join(keywords)} ===")
        threading.Thread(
            target=self._consolidate_all_worker,
            args=(fy, keywords),
            daemon=True,
        ).start()

    def _consolidate_all_worker(self, fy: str, keywords: list[str]):
        from master_consolidate.consolidation.consolidate import consolidate_all

        segment_files = self._by_segment_worker(fy, keywords)
        if not segment_files:
            self._log("No segment files produced, aborting Consolidate All.")
            return
        out_dir = Path(self.var_out_all.get())
        try:
            final = consolidate_all(segment_files, out_dir)
            self._log(f"[All] Saved → {final}")
        except Exception as exc:
            self._log(f"[All] ERROR: {exc}")
        self._log("=== Consolidate All complete ===")

    # ── helpers ──────────────────────────────────────────────────────────────

    def _selected_keywords(self) -> list[str]:
        kw = []
        if self.chk_hp.get():
            kw.append("HP Cost")
        if self.chk_odm.get():
            kw.append("ODM Cost")
        if self.chk_rebate.get():
            kw.append("Rebate")
        return kw


def run(cfg: dict, save_cfg_fn):
    app = App(cfg, save_cfg_fn)
    app.mainloop()
