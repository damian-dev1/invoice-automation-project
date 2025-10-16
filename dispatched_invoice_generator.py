"""
Dispatched Invoice Generator (Interactive Edition)
- Corrects From/To placement and spacing in PDF
- Adds Export to CSV / Excel for the loaded orders
"""
import os
import json
import csv
import threading
from typing import Any, Dict, List, Optional
import tkinter as tk
from tkinter import filedialog, messagebox, Menu
import ttkbootstrap as ttk
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import requests
from fpdf import FPDF
try:
    import pandas as pd  # type: ignore
    HAS_PANDAS = True
except Exception:
    HAS_PANDAS = False
CONFIG_FILE = "invoice_dispatch_app_config.json"
def load_config() -> Dict[str, Any]:
    defaults = {
        "auth_token": "",
        "mark_invoiced_url": "",
        "extra_headers_json": "{}",
        "default_status": "DISPATCH",
        "default_limit": 10,
        "default_offset": 0,
        "default_sort": "desc",
        "theme": "cosmo",
    }
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                config = json.load(f)
                defaults.update(config)
        except (json.JSONDecodeError, IOError) as e:
            print(f"Warning: Could not load config file. Error: {e}")
    return defaults
def save_config(cfg: Dict[str, Any]) -> None:
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=4)
    except IOError as e:
        print(f"Error: Could not save config file. Error: {e}")
class APIClient:
    def __init__(self, auth_token: str, timeout: int = 30, retries: int = 3, backoff: float = 0.5):
        self.base_urls = {
            "orders": "https://api.virtualstock.com/restapi/v4/orders/",
            "suppliers": "https://www.the-edge.io/restapi/v4/suppliers/",
        }
        self.session = self._create_session(auth_token, retries, backoff)
        self.timeout = timeout
    def _create_session(self, auth_token: str, retries: int, backoff: float) -> requests.Session:
        session = requests.Session()
        session.headers.update(
            {
                "Authorization": f"Basic {auth_token}",
                "Accept": "application/json",
                "Content-Type": "application/json",
                "User-Agent": "HN-InvoiceApp/3.0",
            }
        )
        retry_strategy = Retry(
            total=retries,
            read=retries,
            connect=retries,
            backoff_factor=backoff,
            status_forcelist=(429, 500, 502, 503, 504),
            allowed_methods={"GET", "POST"},
            raise_on_status=False,
        )
        adapter = HTTPAdapter(max_retries=retry_strategy)
        session.mount("https://", adapter)
        session.mount("http://", adapter)
        return session
    def _request(self, method: str, url: str, **kwargs) -> requests.Response:
        try:
            resp = self.session.request(method, url, timeout=self.timeout, **kwargs)
            resp.raise_for_status()
            return resp
        except requests.exceptions.RequestException as e:
            raise ConnectionError(f"API request failed: {e}") from e
    def get(self, url: str, **kwargs) -> requests.Response:
        return self._request("get", url, **kwargs)
    def post(self, url: str, json_body: dict, headers: Optional[dict] = None) -> requests.Response:
        hdrs = self.session.headers.copy()
        if headers:
            hdrs.update(headers)
        return self._request("post", url, json=json_body, headers=hdrs)
    def fetch_orders(self, status: str, limit: int, offset: int, sort: str) -> Dict[str, Any]:
        params = {"status": status, "limit": limit, "offset": offset, "sort": sort}
        return self.get(self.base_urls["orders"], params=params).json()
    def fetch_supplier_details(self, supplier_id: str) -> Dict[str, Any]:
        return self.get(f"{self.base_urls['suppliers']}{supplier_id}/").json()
class ProInvoicePDF(FPDF):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.set_margins(12, 15, 12)
        self.set_auto_page_break(auto=True, margin=15)
    def header(self):
        self.set_font("Helvetica", "B", 18)
        self.cell(0, 8, "TAX INVOICE", ln=True, align="C")
        self.ln(2)
        self.set_draw_color(220, 220, 220)
        self.line(self.l_margin, self.get_y(), self.w - self.r_margin, self.get_y())
        self.ln(5)
    def footer(self):
        self.set_y(-12)
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(128, 128, 128)
        self.cell(0, 8, f"Page {self.page_no()}/{{nb}}", 0, 0, "C")
        self.set_text_color(0, 0, 0)
def _money(val: Any) -> str:
    try:
        return f"{float(val):.2f}"
    except (ValueError, TypeError):
        return "0.00"
def export_invoice_pdf_pro(inv_data: Dict[str, Any], save_path: str) -> None:
    """
    Draws two non-overlapping party cards ("From" on left, "To" on right) and items/summary.
    Also auto-corrects if bill_from and bill_to have the same company name.
    """
    bf = inv_data.get("bill_from", {}) or {}
    bt = inv_data.get("bill_to", {}) or {}
    if bf.get("company_name") == bt.get("company_name"):
        bf, bt = ({"company_name": "Supplier"}, bt)
        inv_data["bill_from"] = bf
        inv_data["bill_to"] = bt
    pdf = ProInvoicePDF(orientation="P", unit="mm", format="A4")
    pdf.alias_nb_pages()
    pdf.add_page()
    page_w = pdf.w - pdf.l_margin - pdf.r_margin
    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(0, 6, f"Invoice #: {inv_data.get('invoice_number', 'N/A')}", ln=True)
    pdf.set_font("Helvetica", "", 10)
    pdf.cell(0, 6, f"Invoice Date: {inv_data.get('invoice_date', 'N/A')}", ln=True)
    pdf.cell(0, 6, f"Currency: {inv_data.get('currency', 'AUD')}", ln=True)
    pdf.ln(4)  # slightly tighter but clean
    def draw_party_card(title: str, party: dict, width: float) -> float:
        start_y = pdf.get_y()
        pdf.set_font("Helvetica", "B", 11)
        pdf.cell(width, 6, title, 0, 1, "L")  # title as a single cell
        pdf.set_font("Helvetica", "", 10)
        lines = [
            party.get("company_name"),
            (party.get("address", {}) or {}).get("line_1"),
            " ".join(
                filter(
                    None,
                    [
                        (party.get("address", {}) or {}).get("city"),
                        (party.get("address", {}) or {}).get("state"),
                        (party.get("address", {}) or {}).get("postal_code"),
                    ],
                )
            ),
            f"Phone: {party.get('phone')}" if party.get("phone") else None,
            f"Email: {party.get('email')}" if party.get("email") else None,
        ]
        for line in filter(None, lines):
            pdf.multi_cell(width, 5, line, 0, "L")
        end_y = pdf.get_y()
        pdf.set_draw_color(230, 230, 230)
        pdf.line(pdf.get_x(), end_y, pdf.get_x() + width, end_y)
        pdf.set_draw_color(0, 0, 0)
        return end_y - start_y
    y_start = pdf.get_y()
    col_w = (page_w / 2) - 5
    pdf.set_xy(pdf.l_margin, y_start)
    x_left = pdf.get_x()
    y_left = pdf.get_y()
    h_left = draw_party_card("From", inv_data.get("bill_from", {}) or {}, col_w)
    pdf.set_xy(pdf.l_margin + col_w + 10, y_start)
    x_right = pdf.get_x()
    y_right = pdf.get_y()
    h_right = draw_party_card("To", inv_data.get("bill_to", {}) or {}, col_w)
    pdf.set_y(y_start + max(h_left, h_right) + 6)
    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 6, "Order Details", ln=True)
    od = inv_data.get("order_details", {}) or {}
    details = {"Order Ref": od.get("order_reference"), "Dispatch Date": od.get("dispatch_date")}
    for label, value in details.items():
        if value:
            pdf.set_font("Helvetica", "B", 10)
            pdf.cell(40, 6, f"{label}:", 0, 0)
            pdf.set_font("Helvetica", "", 10)
            pdf.cell(0, 6, str(value), 0, 1)
    pdf.ln(4)
    widths = [page_w * 0.55, page_w * 0.12, page_w * 0.16, page_w * 0.17]
    def draw_table_header():
        pdf.set_font("Helvetica", "B", 10)
        pdf.set_fill_color(240, 240, 240)
        for i, h in enumerate(["Item Description", "Qty", "Unit Price", "Line Total"]):
            pdf.cell(widths[i], 7, h, border=1, align="C", fill=True)
        pdf.ln()
    draw_table_header()
    pdf.set_font("Helvetica", "", 10)
    for item in inv_data.get("items", []):
        if pdf.get_y() > (pdf.h - pdf.b_margin - 20):
            pdf.add_page()
            draw_table_header()
        y0 = pdf.get_y()
        pdf.multi_cell(widths[0], 7, str(item.get("name", "N/A")), border="LR")
        h = pdf.get_y() - y0
        pdf.set_xy(pdf.l_margin + widths[0], y0)
        pdf.cell(widths[1], h, str(item.get("quantity", 0)), border="R", align="R")
        pdf.cell(widths[2], h, _money(item.get("unit_cost_price", 0)), border="R", align="R")
        pdf.cell(widths[3], h, _money(item.get("total", 0)), border="R", align="R")
        pdf.ln(h)
    pdf.cell(sum(widths), 0, "", "T")
    pdf.ln()
    totals = inv_data.get("totals", {}) or {}
    totals_data = [
        ("Subtotal", _money(totals.get("subtotal"))),
        ("Freight", _money(totals.get("freight"))),
        ("Tax", _money(totals.get("tax"))),
        ("Grand Total", _money(totals.get("grand_total"))),
    ]
    if pdf.get_y() > (pdf.h - pdf.b_margin - 35):
        pdf.add_page()
    for label, value in totals_data:
        is_total = label == "Grand Total"
        pdf.set_x(page_w + pdf.l_margin - 80)
        pdf.set_font("Helvetica", "B" if is_total else "", 12 if is_total else 10)
        pdf.cell(45, 8, label, border="T" if is_total else 0, align="R")
        pdf.cell(35, 8, value, border="T" if is_total else 0, align="R", ln=True)
    pdf.ln(5)
    pdf.set_font("Helvetica", "I", 9)
    pdf.set_text_color(100, 100, 100)
    pdf.multi_cell(0, 5, "Note: Please contact accounts within 7 days for any discrepancies.")
    pdf.output(save_path)
class EditInvoiceWindow(tk.Toplevel):
    def __init__(self, parent, invoice_data: Dict[str, Any]):
        super().__init__(parent)
        self.title("Edit Invoice")
        self.transient(parent)
        self.grab_set()
        self.resizable(True, True)
        self.invoice_data = invoice_data
        self.geometry("500x300")
        self.freight_var = tk.StringVar(value=_money(self.invoice_data["totals"].get("freight", 0.0)))
        self.grand_total_var = tk.StringVar(value=_money(self.invoice_data["totals"].get("grand_total", 0.0)))
        self._build()
        self.freight_var.trace_add("write", self._recalc)
        self.protocol("WM_DELETE_WINDOW", self.destroy)
    def _build(self):
        main = ttk.Frame(self, padding=15)
        main.pack(fill=tk.BOTH, expand=True)
        ttk.Label(
            main, text=f"Order: {self.invoice_data['order_details']['order_reference']}", font="-weight bold"
        ).pack(fill=tk.X)
        ttk.Separator(main, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)
        row1 = ttk.Frame(main); row1.pack(fill=tk.X, pady=5)
        ttk.Label(row1, text="Freight Cost:", width=15).pack(side=tk.LEFT)
        entry = ttk.Entry(row1, textvariable=self.freight_var); entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        entry.focus_set()
        row2 = ttk.Frame(main); row2.pack(fill=tk.X, pady=5)
        ttk.Label(row2, text="Grand Total:", width=15).pack(side=tk.LEFT)
        ttk.Label(row2, textvariable=self.grand_total_var, font="-weight bold").pack(side=tk.LEFT)
        btns = ttk.Frame(main); btns.pack(side=tk.BOTTOM, fill=tk.X, pady=(20, 0))
        ttk.Button(btns, text="Save Changes", command=self._save, bootstyle="success").pack(side=tk.RIGHT)
        ttk.Button(btns, text="Cancel", command=self.destroy, bootstyle="secondary").pack(side=tk.RIGHT, padx=5)
    def _recalc(self, *_):
        try:
            freight = float(self.freight_var.get())
        except ValueError:
            freight = 0.0
        subtotal = float(self.invoice_data["totals"].get("subtotal", 0.0) or 0.0)
        tax = float(self.invoice_data["totals"].get("tax", 0.0) or 0.0)
        self.grand_total_var.set(_money(subtotal + tax + freight))
    def _save(self):
        try:
            self.invoice_data["totals"]["freight"] = float(self.freight_var.get())
            self.invoice_data["totals"]["grand_total"] = float(self.grand_total_var.get())
            self.destroy()
        except ValueError:
            messagebox.showerror("Invalid Input", "Freight must be a valid number.", parent=self)
class InvoiceApp:
    def __init__(self, themename: str = "cosmo"):
        self.cfg = load_config()
        self.root = ttk.Window()
        self.root.title("Dispatched Invoice Generator")
        self.root.geometry("1050x520")
        self.fetched_invoices: List[Dict[str, Any]] = []
        self._setup_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)
    def _setup_ui(self):
        outer = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        outer.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        sidebar = ttk.Frame(outer, width=360, padding=10)
        main_panel = ttk.Frame(outer, padding=10)
        outer.add(sidebar, weight=0)
        outer.add(main_panel, weight=1)
        self._build_sidebar(sidebar)
        self._build_main(main_panel)
    def _build_sidebar(self, parent: ttk.Frame):
        parent.columnconfigure(0, weight=1)
        r = 0
        auth = ttk.LabelFrame(parent, text="Authentication", padding=8)
        auth.grid(row=r, column=0, sticky="ew", pady=(0, 6)); r += 1
        auth.columnconfigure(0, weight=1)
        self.auth_var = tk.StringVar(value=self.cfg.get("auth_token"))
        ttk.Entry(auth, show="*", textvariable=self.auth_var).grid(row=0, column=0, sticky="ew")
        opts = ttk.LabelFrame(parent, text="Fetch Options", padding=8)
        opts.grid(row=r, column=0, sticky="ew", pady=6); r += 1
        opts.columnconfigure(1, weight=1)
        self.status_var = tk.StringVar(value=self.cfg.get("default_status"))
        self.limit_var = tk.StringVar(value=str(self.cfg.get("default_limit")))
        self.offset_var = tk.StringVar(value=str(self.cfg.get("default_offset")))
        self.sort_var = tk.StringVar(value=self.cfg.get("default_sort"))
        self._lbl(opts, "Status", ttk.Combobox(opts, textvariable=self.status_var,
                                               values=["ORDER", "PROCESS", "DISPATCH", "CANCEL"], state="readonly"), 0)
        self._lbl(opts, "Limit", ttk.Entry(opts, textvariable=self.limit_var, width=10), 1)
        self._lbl(opts, "Offset", ttk.Entry(opts, textvariable=self.offset_var, width=10), 2)
        self._lbl(opts, "Sort", ttk.Combobox(opts, textvariable=self.sort_var,
                                             values=["asc", "desc"], state="readonly"), 3)
        btns = ttk.Frame(parent); btns.grid(row=r, column=0, sticky="ew", pady=(8, 0)); r += 1
        btns.columnconfigure(0, weight=1)
        self.fetch_btn = ttk.Button(btns, text="Fetch Orders", command=self.fetch_orders_threaded, bootstyle="primary")
        self.fetch_btn.grid(row=0, column=0, sticky="ew", ipady=5)
        self.bulk_export_btn = ttk.Button(btns, text="Bulk Export PDFs", command=self.bulk_export_threaded,
                                          bootstyle="secondary-outline")
        self.bulk_export_btn.grid(row=1, column=0, sticky="ew", pady=4)
        exp_row = ttk.Frame(parent); exp_row.grid(row=r, column=0, sticky="ew", pady=(8, 0)); r += 1
        exp_row.columnconfigure(0, weight=1)
        ttk.Button(exp_row, text="Export Table → CSV", command=self.export_csv).grid(row=0, column=0, sticky="ew")
        ttk.Button(exp_row, text="Export Table → Excel", command=self.export_excel,
                   bootstyle="info" if HAS_PANDAS else "secondary").grid(row=1, column=0, sticky="ew", pady=4)
        theme = ttk.LabelFrame(parent, text="Theme", padding=8)
        theme.grid(row=r, column=0, sticky="ew", pady=(12, 0)); r += 1
        theme.columnconfigure(0, weight=1)
        self.theme_var = tk.StringVar(value=self.root.style.theme_use())
        theme_combo = ttk.Combobox(
            theme,
            textvariable=self.theme_var,
            values=sorted(self.root.style.theme_names()),
            state="readonly"
        )
        theme_combo.pack(fill="x")
        theme_combo.bind("<<ComboboxSelected>>", self.change_theme)
    def _lbl(self, parent, text, widget, row):
        ttk.Label(parent, text=text).grid(row=row, column=0, sticky="w", padx=(0, 10))
        widget.grid(row=row, column=1, sticky="ew")
    def change_theme(self, _evt=None):
        self.root.style.theme_use(self.theme_var.get())
    def _build_main(self, parent: ttk.Frame):
        parent.rowconfigure(0, weight=1); parent.columnconfigure(0, weight=1)
        columns = ("order_reference", "invoice_number", "invoice_date", "customer", "grand_total", "currency")
        self.order_tree = ttk.Treeview(parent, columns=columns, show="headings", bootstyle="primary")
        self.order_tree.grid(row=0, column=0, sticky="nsew")
        yscroll = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=self.order_tree.yview)
        yscroll.grid(row=0, column=1, sticky="ns"); self.order_tree.configure(yscrollcommand=yscroll.set)
        col_defs = {
            "order_reference": ("Order Reference", 260),
            "invoice_number": ("Invoice #", 140),
            "invoice_date": ("Date", 110),
            "customer": ("Customer", 220),
            "grand_total": ("Grand Total", 110),
            "currency": ("Currency", 90),
        }
        for cid, (label, width) in col_defs.items():
            self.order_tree.heading(cid, text=label)
            self.order_tree.column(cid, width=width, anchor="w")
        self.order_tree.bind("<Double-1>", self.edit_selected_order)
        self.order_tree.bind("<Button-3>", self._on_right_click)
        self.status = ttk.Label(parent, text="Ready", anchor="w")
        self.status.grid(row=1, column=0, sticky="ew", pady=(5, 0))
        self.progress = ttk.Progressbar(parent, mode="determinate", maximum=100, bootstyle="striped")
        self.progress.grid(row=2, column=0, sticky="ew", pady=(2, 0))
    def _table_rows(self) -> List[Dict[str, str]]:
        rows = []
        for iid in self.order_tree.get_children():
            v = self.order_tree.item(iid, "values")
            rows.append(
                {
                    "order_reference": v[0],
                    "invoice_number": v[1],
                    "invoice_date": v[2],
                    "customer": v[3],
                    "grand_total": v[4],
                    "currency": v[5],
                }
            )
        return rows
    def export_csv(self):
        data = self._table_rows()
        if not data:
            messagebox.showinfo("Export CSV", "No rows to export.")
            return
        path = filedialog.asksaveasfilename(
            title="Save CSV",
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")],
            initialfile="invoices_table.csv",
        )
        if not path:
            return
        try:
            with open(path, "w", newline="", encoding="utf-8") as f:
                writer = csv.DictWriter(
                    f, fieldnames=["order_reference", "invoice_number", "invoice_date", "customer", "grand_total", "currency"]
                )
                writer.writeheader()
                writer.writerows(data)
            messagebox.showinfo("Export CSV", f"Exported {len(data)} rows to:\n{path}")
        except Exception as e:
            messagebox.showerror("Export CSV", str(e))
    def export_excel(self):
        if not HAS_PANDAS:
            messagebox.showwarning(
                "Export Excel",
                "pandas is not installed. Run:\n\n  pip install pandas openpyxl\n\nOr use CSV export.",
            )
            return
        data = self._table_rows()
        if not data:
            messagebox.showinfo("Export Excel", "No rows to export.")
            return
        path = filedialog.asksaveasfilename(
            title="Save Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
            initialfile="invoices_table.xlsx",
        )
        if not path:
            return
        try:
            df = pd.DataFrame(data)
            if "grand_total" in df.columns:
                df["grand_total"] = pd.to_numeric(df["grand_total"], errors="coerce")
            df.to_excel(path, index=False)
            messagebox.showinfo("Export Excel", f"Exported {len(data)} rows to:\n{path}")
        except Exception as e:
            messagebox.showerror("Export Excel", str(e))
    def _on_right_click(self, event):
        row_id = self.order_tree.identify_row(event.y)
        if row_id:
            self.order_tree.selection_set(row_id)
            self.order_tree.focus(row_id)
        self._show_context_menu(event)
    def _show_context_menu(self, event):
        if not self.order_tree.selection():
            return
        menu = Menu(self.root, tearoff=0)
        menu.add_command(label="Edit Selected Order", command=self.edit_selected_order)
        menu.add_command(label="Delete Selected Order(s)", command=self.delete_selected_orders)
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()
    def _on_closing(self):
        self.cfg.update(
            {
                "auth_token": self.auth_var.get(),
                "default_status": self.status_var.get(),
                "theme": self.theme_var.get(),
                "default_limit": int(self.limit_var.get() or 10),
                "default_offset": int(self.offset_var.get() or 0),
                "default_sort": self.sort_var.get(),
            }
        )
        save_config(self.cfg)
        self.root.destroy()
    def fetch_orders_threaded(self):
        if not self.auth_var.get().strip():
            messagebox.showerror("Authentication Error", "Please provide the Auth Token.")
            return
        threading.Thread(target=self._fetch_orders_task, daemon=True).start()
    def _fetch_orders_task(self):
        self.root.after(0, lambda: self.bulk_export_btn.config(bootstyle="secondary-outline"))
        self._update_ui_state(is_busy=True)
        try:
            params = {
                "status": self.status_var.get(),
                "limit": int(self.limit_var.get()),
                "offset": int(self.offset_var.get()),
                "sort": self.sort_var.get(),
            }
            client = APIClient(auth_token=self.auth_var.get().strip())
            self.set_status("Fetching orders...", 0)
            payload = client.fetch_orders(**params)
            results = payload.get("results", [])
            if not results:
                self.set_status("No results found.", 100)
                return
            self.fetched_invoices.clear()
            total = len(results)
            for i, order in enumerate(results, start=1):
                pct = int((i / total) * 100)
                self.set_status(f"Processing order {i}/{total}...", pct)
                self.fetched_invoices.append(self._build_invoice_dict(order))
            self.root.after(0, self.populate_order_tree)
            self.set_status(f"Fetch complete. {len(self.fetched_invoices)} orders loaded.", 100)
        except Exception as e:
            messagebox.showerror("Fetch Error", str(e))
            self.set_status(f"Error: {e}", 0)
        finally:
         self._update_ui_state(is_busy=False)
    def _build_invoice_dict(self, order: Dict[str, Any]) -> Dict[str, Any]:
            def safe_float(value, default=0.0):
                try:
                    return float(value)
                except (TypeError, ValueError):
                    return default
            def safe_int(value, default=0):
                try:
                    return int(float(value))
                except (TypeError, ValueError):
                    return default
            order_ref = str(order.get("order_reference", "N/A"))
            order_date = (order.get("order_date") or "").split("T")[0]
            currency = order.get("currency_code", "AUD")
            items = []
            for i in order.get("items", []):
                items.append({
                    "name": str(i.get("name") or "Goods"),
                    "quantity": safe_int(i.get("quantity")),
                    "unit_cost_price": safe_float(i.get("unit_cost_price")),
                    "total": safe_float(i.get("total")),
                })
            supplier = order.get("supplier_data") or {}
            supplier_address = supplier.get("address") or {}
            bill_from = {
                "company_name": supplier.get("name") or supplier.get("company_name") or "Supplier",
                "address": {
                    "line_1": supplier_address.get("line_1", ""),
                    "city": supplier_address.get("city", ""),
                    "state": supplier_address.get("state", ""),
                    "postal_code": supplier_address.get("postal_code", ""),
                },
                "phone": supplier.get("phone", ""),
                "email": supplier.get("email", ""),
            }
            retailer = order.get("retailer_data") or {}
            bill_to = {
                "company_name": retailer.get("name", "Harvey Norman"),
                "address": {
                    "line_1": "145-151 Arthur Street",
                    "city": "Homebush West",
                    "state": "NSW",
                    "postal_code": "2140",
                },
                "phone": "+61 2 9999 8888",
                "email": "orders@harveynorman.com.au"
            }
            subtotal = safe_float(order.get("subtotal"))
            tax = safe_float(order.get("tax"))
            total = safe_float(order.get("total"))
            freight = 0.0
            invoice = {
                "invoice_number": f"INV-{order_ref[:11]}",
                "invoice_date": order_date,
                "currency": currency,
                "bill_from": bill_from,
                "bill_to": bill_to,
                "order_details": {
                    "order_reference": order_ref,
                    "dispatch_date": order_date,
                    "comment": order.get("comment", "")
                },
                "items": items,
                "totals": {
                    "subtotal": subtotal,
                    "freight": freight,
                    "tax": tax,
                    "grand_total": total + freight
                }
            }
            return invoice
    def populate_order_tree(self):
        for iid in self.order_tree.get_children():
            self.order_tree.delete(iid)
        for inv in self.fetched_invoices:
            values = (
                inv["order_details"]["order_reference"],
                inv["invoice_number"],
                inv["invoice_date"],
                inv["bill_to"]["company_name"],
                _money(inv["totals"]["grand_total"]),
                inv["currency"],
            )
            self.order_tree.insert("", tk.END, values=values)
        if self.fetched_invoices:
            self.bulk_export_btn.config(bootstyle="success")
    def edit_selected_order(self, _event=None):
        sel = self.order_tree.selection()
        if not sel:
            return
        order_ref = self.order_tree.item(sel[0], "values")[0]
        inv = next((x for x in self.fetched_invoices if x["order_details"]["order_reference"] == order_ref), None)
        if not inv:
            return
        dlg = EditInvoiceWindow(self.root, inv)
        self.root.wait_window(dlg)
        self.populate_order_tree()
    def delete_selected_orders(self):
        sel = self.order_tree.selection()
        if not sel:
            return
        if not messagebox.askyesno("Confirm Deletion", f"Delete {len(sel)} selected order(s)?"):
            return
        refs = {self.order_tree.item(i, "values")[0] for i in sel}
        self.fetched_invoices = [x for x in self.fetched_invoices if x["order_details"]["order_reference"] not in refs]
        self.populate_order_tree()
    def bulk_export_threaded(self):
        if not self.fetched_invoices:
            messagebox.showerror("No Data", "No invoices to export.")
            return
        folder = filedialog.askdirectory(title="Choose a folder to save all invoices")
        if not folder:
            return
        threading.Thread(target=self._bulk_export_task, args=(folder,), daemon=True).start()
    def _bulk_export_task(self, folder: str):
        self._update_ui_state(is_busy=True)
        total = len(self.fetched_invoices)
        ok, err = 0, 0
        for i, inv in enumerate(self.fetched_invoices, start=1):
            self.set_status(f"Exporting {i}/{total}...", int((i / total) * 100))
            path = os.path.join(folder, f"{inv['invoice_number']}.pdf")
            try:
                export_invoice_pdf_pro(inv, path)
                ok += 1
            except Exception:
                err += 1
        self.set_status(f"Bulk export finished. Success: {ok}, Failed: {err}.", 100)
        self._update_ui_state(is_busy=False)
    def _update_ui_state(self, is_busy: bool):
        def _apply():
            for btn in (self.fetch_btn, self.bulk_export_btn):
                btn.config(state="disabled" if is_busy else "normal")
        self.root.after(0, _apply)
    def set_status(self, text: str, progress_val: Optional[int] = None):
        def _apply():
            self.status.config(text=text)
            if progress_val is not None:
                self.progress.config(value=progress_val)
        self.root.after(0, _apply)
if __name__ == "__main__":
    app = InvoiceApp()
    app.root.mainloop()
