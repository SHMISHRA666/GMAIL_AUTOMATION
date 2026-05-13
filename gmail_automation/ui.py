from __future__ import annotations

import queue
import threading
from pathlib import Path
from typing import Callable, TypeVar

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .models import SendBatch
from .ui_workflow import BatchRunResult, GmailAutomationWorkflow, WorkflowSummary

T = TypeVar("T")


class GmailAutomationUI:
    def __init__(self, root: tk.Tk, config_path: Path | None = None) -> None:
        self.root = root
        self.workflow = GmailAutomationWorkflow(config_path)
        self.work_queue: queue.Queue[tuple[bool, object, Callable[[object], None] | None]] = queue.Queue()
        self.selected_file = tk.StringVar()
        self.batch_size = tk.IntVar(value=self.workflow.config.batch_size)
        self.status_text = tk.StringVar(value="Choose an Excel workbook to begin.")
        self.summary_text = tk.StringVar(value=self._config_summary())
        self.batch_filter = tk.StringVar(value="All batches")
        self.current_batch: SendBatch | None = None

        self.root.title("Gmail Confirmation Automation")
        self.root.geometry("1100x720")
        self.root.minsize(980, 620)
        self._build()
        self._poll_work_queue()

    def _build(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=0)

        notebook = ttk.Notebook(self.root)
        notebook.grid(row=0, column=0, sticky="nsew", padx=12, pady=(12, 6))

        run_tab = ttk.Frame(notebook, padding=12)
        results_tab = ttk.Frame(notebook, padding=12)
        notebook.add(run_tab, text="Run")
        notebook.add(results_tab, text="Results")

        self._build_run_tab(run_tab)
        self._build_results_tab(results_tab)

        status = ttk.Label(self.root, textvariable=self.status_text, anchor="w")
        status.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 10))

    def _build_run_tab(self, tab: ttk.Frame) -> None:
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(3, weight=1)

        file_frame = ttk.LabelFrame(tab, text="Input Workbook", padding=10)
        file_frame.grid(row=0, column=0, sticky="ew")
        file_frame.columnconfigure(1, weight=1)

        ttk.Label(file_frame, text="Excel file").grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Entry(file_frame, textvariable=self.selected_file).grid(row=0, column=1, sticky="ew", padx=(0, 8))
        ttk.Button(file_frame, text="Browse...", command=self._choose_file).grid(row=0, column=2)

        controls = ttk.LabelFrame(tab, text="Batch Controls", padding=10)
        controls.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        for col in range(5):
            controls.columnconfigure(col, weight=0)
        controls.columnconfigure(5, weight=1)

        ttk.Label(controls, text="Batch size").grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Spinbox(controls, from_=1, to=500, textvariable=self.batch_size, width=8).grid(row=0, column=1, sticky="w", padx=(0, 12))
        ttk.Button(controls, text="Load Workbook", command=self._load_workbook).grid(row=0, column=2, padx=(0, 8))
        ttk.Button(controls, text="Generate Documents", command=self._generate_documents).grid(row=0, column=3, padx=(0, 8))
        ttk.Button(controls, text="Run Preview Batch", command=self._run_next_batch).grid(row=0, column=4, padx=(0, 8))
        ttk.Label(controls, textvariable=self.summary_text, anchor="e").grid(row=0, column=5, sticky="ew")

        ttk.Label(tab, text="Current Batch Preview").grid(row=2, column=0, sticky="w", pady=(14, 4))
        self.batch_tree = ttk.Treeview(tab, columns=("party", "email", "docs", "batch", "sent", "status"), show="headings", height=10)
        self._configure_tree(
            self.batch_tree,
            {
                "party": ("Customer", 260),
                "email": ("Email", 240),
                "docs": ("Docs Created", 120),
                "batch": ("Batch Selected", 160),
                "sent": ("Mail Sent", 110),
                "status": ("Status", 140),
            },
        )
        self.batch_tree.grid(row=3, column=0, sticky="nsew")
        ttk.Button(tab, text="Preview Next Batch", command=self._refresh_all).grid(row=4, column=0, sticky="e", pady=(8, 0))

    def _build_results_tab(self, tab: ttk.Frame) -> None:
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(1, weight=1)

        filter_frame = ttk.Frame(tab)
        filter_frame.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        filter_frame.columnconfigure(4, weight=1)

        ttk.Label(filter_frame, text="Batch").grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.batch_filter_box = ttk.Combobox(filter_frame, textvariable=self.batch_filter, values=["All batches"], state="readonly", width=32)
        self.batch_filter_box.grid(row=0, column=1, sticky="w", padx=(0, 8))
        self.batch_filter_box.bind("<<ComboboxSelected>>", lambda _event: self._refresh_results())
        ttk.Button(filter_frame, text="Refresh Results", command=self._refresh_all).grid(row=0, column=2, padx=(0, 8))
        ttk.Button(filter_frame, text="Send/Resend Selected Pending", command=self._send_selected_pending).grid(row=0, column=3, padx=(0, 8))
        ttk.Button(filter_frame, text="Send Pending In Batch", command=self._send_pending_in_batch).grid(row=0, column=4, sticky="w")

        self.results_tree = ttk.Treeview(
            tab,
            columns=("party", "email", "docs", "batch", "sent", "status", "error"),
            show="headings",
            selectmode="extended",
        )
        self._configure_tree(
            self.results_tree,
            {
                "party": ("Customer", 220),
                "email": ("Email", 210),
                "docs": ("Docs Created", 110),
                "batch": ("Batch Selected", 170),
                "sent": ("Mail Sent", 100),
                "status": ("Status", 130),
                "error": ("Error", 280),
            },
        )
        self.results_tree.grid(row=1, column=0, sticky="nsew")

    def _choose_file(self) -> None:
        path = filedialog.askopenfilename(
            title="Select input Excel workbook",
            filetypes=[("Excel workbooks", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        if path:
            self.selected_file.set(path)
            self._load_workbook()

    def _load_workbook(self) -> None:
        path = self.selected_file.get().strip()
        if not path:
            messagebox.showinfo("Choose Excel File", "Please choose the input Excel workbook.")
            return
        try:
            batch_size = self._validated_batch_size()
        except ValueError as exc:
            messagebox.showerror("Invalid Batch Size", str(exc))
            return

        self._run_background(
            lambda: self.workflow.load_input_file(Path(path), batch_size),
            self._on_summary_updated,
            "Loading workbook...",
        )

    def _generate_documents(self) -> None:
        if not self._ensure_loaded():
            return
        self._run_background(
            self.workflow.generate_documents,
            self._on_summary_updated,
            "Generating and verifying documents...",
        )

    def _run_next_batch(self) -> None:
        if not self._ensure_loaded():
            return
        self._refresh_preview()
        if self.current_batch is None or not self.current_batch.rows:
            messagebox.showinfo("No Batch Ready", "There are no docs-ready customers with mail pending.")
            return
        names = ", ".join(row.party_name for row in self.current_batch.rows[:5])
        suffix = "..." if len(self.current_batch.rows) > 5 else ""
        approved = messagebox.askyesno(
            "Approve Batch Send",
            f"Send the currently previewed batch of {len(self.current_batch.rows)} customer(s)?\n\n{names}{suffix}",
        )
        if not approved:
            return

        self._run_background(
            self.workflow.run_next_batch,
            self._on_batch_run,
            "Sending approved batch...",
        )

    def _send_selected_pending(self) -> None:
        if not self._ensure_loaded():
            return
        row_ids = {item for item in self.results_tree.selection()}
        if not row_ids:
            messagebox.showinfo("Select Rows", "Select one or more docs-ready rows with mail not sent.")
            return
        if not messagebox.askyesno("Approve Send", f"Send or resend {len(row_ids)} selected pending email(s)?"):
            return
        self._run_background(
            lambda: self.workflow.send_pending(row_ids=row_ids),
            self._on_batch_run,
            "Sending selected pending emails...",
        )

    def _send_pending_in_batch(self) -> None:
        if not self._ensure_loaded():
            return
        batch_id = self._selected_batch_id()
        if not batch_id:
            messagebox.showinfo("Select Batch", "Choose a batch in the results filter first.")
            return
        if not messagebox.askyesno("Approve Send", f"Send all docs-ready pending emails from batch {batch_id}?"):
            return
        self._run_background(
            lambda: self.workflow.send_pending(batch_id=batch_id),
            self._on_batch_run,
            "Sending pending emails from batch...",
        )

    def _refresh_all(self) -> None:
        if not self._ensure_loaded(show_message=False):
            return
        self._refresh_preview()
        self._refresh_results()
        self.summary_text.set(self._summary_text(self.workflow.get_summary()))

    def _refresh_preview(self) -> None:
        try:
            self.current_batch = self.workflow.preview_next_batch()
            rows = self.current_batch.rows if self.current_batch else []
            statuses = {status.row_id: status for status in self.workflow.get_results()}
            self._clear_tree(self.batch_tree)
            for row in rows:
                status = statuses.get(row.row_id)
                self.batch_tree.insert(
                    "",
                    "end",
                    iid=row.row_id,
                    values=(
                        row.party_name,
                        row.email,
                        status.docs_created if status else row.state.status,
                        status.batch_selected if status else "Not selected",
                        status.mail_sent if status else "Not sent",
                        status.status if status else row.state.status,
                    ),
                )
        except Exception as exc:
            self._show_error(exc)

    def _refresh_results(self) -> None:
        try:
            selected_batch = self._selected_batch_id()
            rows = self.workflow.get_results()
            if selected_batch:
                rows = [row for row in rows if row.batch_id == selected_batch]
            self._clear_tree(self.results_tree)
            for row in rows:
                self.results_tree.insert(
                    "",
                    "end",
                    iid=row.row_id,
                    values=(row.party_name, row.email, row.docs_created, row.batch_selected, row.mail_sent, row.status, row.error),
                )
            batch_ids = self.workflow.get_batch_ids()
            values = ["All batches"] + batch_ids
            self.batch_filter_box.configure(values=values)
            if self.batch_filter.get() not in values:
                self.batch_filter.set("All batches")
        except Exception as exc:
            self._show_error(exc)

    def _run_background(self, operation: Callable[[], T], on_success: Callable[[T], None], busy_text: str) -> None:
        self.status_text.set(busy_text)

        def worker() -> None:
            try:
                self.work_queue.put((True, operation(), on_success))
            except Exception as exc:
                self.work_queue.put((False, exc, None))

        threading.Thread(target=worker, daemon=True).start()

    def _poll_work_queue(self) -> None:
        try:
            while True:
                success, payload, callback = self.work_queue.get_nowait()
                if success:
                    if callback is not None:
                        callback(payload)
                    self.status_text.set("Ready.")
                else:
                    self._show_error(payload)
        except queue.Empty:
            pass
        self.root.after(100, self._poll_work_queue)

    def _on_summary_updated(self, summary: WorkflowSummary) -> None:
        self.summary_text.set(self._summary_text(summary))
        self._refresh_all()

    def _on_batch_run(self, result: BatchRunResult) -> None:
        messagebox.showinfo(
            "Batch Complete",
            f"Batch: {result.batch_id or 'None'}\nAttempted: {result.attempted}\nSent: {result.sent}\nFailed: {result.failed}",
        )
        self._refresh_all()

    def _selected_batch_id(self) -> str:
        value = self.batch_filter.get()
        return "" if value == "All batches" else value

    def _validated_batch_size(self) -> int:
        value = int(self.batch_size.get())
        if value < 1:
            raise ValueError("Batch size must be at least 1.")
        self.workflow.set_batch_size(value)
        return value

    def _ensure_loaded(self, show_message: bool = True) -> bool:
        if self.workflow.master_path is not None:
            return True
        if show_message:
            messagebox.showinfo("Choose Excel File", "Please choose and load the input Excel workbook first.")
        return False

    def _summary_text(self, summary: WorkflowSummary) -> str:
        return (
            f"{self._config_summary()} | Rows: {summary.total_rows} | "
            f"Ready: {summary.ready_to_send} | Sent: {summary.sent_rows} | Issues: {summary.failed_rows}"
        )

    def _config_summary(self) -> str:
        sender = self.workflow.config.sender_email or "not configured"
        return f"Sender: {sender} | send_mode: {self.workflow.config.send_mode}"

    def _show_error(self, error: object) -> None:
        self.status_text.set("Action failed.")
        messagebox.showerror("Gmail Automation", str(error))

    def _configure_tree(self, tree: ttk.Treeview, columns: dict[str, tuple[str, int]]) -> None:
        for key, (heading, width) in columns.items():
            tree.heading(key, text=heading)
            tree.column(key, width=width, minwidth=70, stretch=True)

    def _clear_tree(self, tree: ttk.Treeview) -> None:
        for item in tree.get_children():
            tree.delete(item)


def launch_ui(config_path: Path | None = None) -> int:
    root = tk.Tk()
    GmailAutomationUI(root, config_path)
    root.mainloop()
    return 0
