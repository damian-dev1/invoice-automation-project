import json
import unicodedata
import threading
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
HTTPX_OK = True
try:
    import httpx
except Exception:
    HTTPX_OK = False
from fpdf import FPDF
@dataclass
class Address:
    line_1: str = ""
    city: str = ""
    state: str = ""
    postal_code: str = ""
@dataclass
class BillTo:
    company_name: str = ""
    address: Address = field(default_factory=Address)
    phone: str = ""
    email: str = ""
@dataclass
class OrderDetails:
    order_reference: str = ""
    additional_order_reference: str = ""
    end_user_purchase_order_reference: str = ""
    promised_date: str = ""
    comment: str = ""
@dataclass
class LineItem:
    name: str
    quantity: float
    unit_cost_price: float
    tax: float
    total: float
@dataclass
class Totals:
    subtotal: float = 0.0
    freight: float = 0.0
    tax: float = 0.0
    grand_total: float = 0.0
@dataclass
class Invoice:
    invoice_number: str
    invoice_date: str
    currency: str
    bill_to: BillTo
    order_details: OrderDetails
    items: List[LineItem]
    totals: Totals
SMART_MAP = {
    0x2018: "'", 0x2019: "'",  # ‘ ’
    0x201C: '"', 0x201D: '"',  # “ ”
    0x2013: "-", 0x2014: "-",  # – —
    0x00A0: " ",
}
TRANS_TABLE = str.maketrans(SMART_MAP)
def _to_float(x, default=0.0) -> float:
    try:
        return float(x)
    except Exception:
        return default
def latin1_sanitize(s: str) -> str:
    if not isinstance(s, str):
        s = str(s)
    s = unicodedata.normalize("NFKC", s).translate(TRANS_TABLE)
    return s.encode("latin-1", "ignore").decode("latin-1")
def parse_vs_orders_payload(raw: Dict[str, Any]) -> Tuple[List[Dict[str, Any]], Optional[str], Optional[str]]:
    if isinstance(raw, dict) and "results" in raw and isinstance(raw["results"], list):
        return raw["results"], raw.get("next"), raw.get("previous")
    return [raw], None, None
def vs_order_to_invoice(order: Dict[str, Any]) -> Invoice:
    retailer = order.get("retailer_data", {}) or {}
    raddr = retailer.get("address", {}) or {}
    items = []
    for it in order.get("items", []) or []:
        qty = _to_float(it.get("quantity", 0))
        unit = _to_float(it.get("unit_cost_price", 0))
        tax = _to_float(it.get("tax", 0))
        total = _to_float(it.get("total", qty * unit + tax))
        items.append(LineItem(
            name=str(it.get("name") or it.get("part_number") or "")[:128],
            quantity=qty,
            unit_cost_price=unit,
            tax=tax,
            total=total
        ))
    subtotal = _to_float(order.get("subtotal", 0))
    tax = _to_float(order.get("tax", 0))
    grand = _to_float(order.get("total", subtotal + tax))
    bill_to = BillTo(
        company_name=retailer.get("name") or "Harvey Norman",
        address=Address(
            line_1=raddr.get("line_1", ""),
            city=raddr.get("city", ""),
            state=raddr.get("state", ""),
            postal_code=raddr.get("postal_code", ""),
        ),
        phone=retailer.get("phone",""),
        email=retailer.get("email",""),
    )
    promised = ""
    if order.get("items"):
        promised = (order["items"][0].get("promised_date") or "")[:10]
    order_details = OrderDetails(
        order_reference=order.get("order_reference",""),
        additional_order_reference=order.get("additional_order_reference","") or order.get("purchase_order_reference","") or "",
        end_user_purchase_order_reference=order.get("end_user_purchase_order_reference","") or "",
        promised_date=promised,
        comment=order.get("comment","") or "",
    )
    invoice_number = f"HN-INV-{order_details.order_reference or 'NA'}"
    invoice_date = (order.get("order_date","") or "")[:10] or f"{datetime.now():%Y-%m-%d}"
    currency = order.get("currency_code") or "AUD"
    return Invoice(
        invoice_number=invoice_number,
        invoice_date=invoice_date,
        currency=currency,
        bill_to=bill_to,
        order_details=order_details,
        items=items,
        totals=Totals(subtotal=subtotal, freight=0.0, tax=tax, grand_total=grand)
    )
def normalize_custom_invoice_payload(raw: Dict[str, Any]) -> Invoice:
    addr = ((raw.get("bill_to") or {}).get("address") or {})
    bill_to = BillTo(
        company_name=(raw.get("bill_to") or {}).get("company_name",""),
        address=Address(
            line_1=addr.get("line_1",""),
            city=addr.get("city",""),
            state=addr.get("state",""),
            postal_code=addr.get("postal_code",""),
        ),
        phone=(raw.get("bill_to") or {}).get("phone",""),
        email=(raw.get("bill_to") or {}).get("email",""),
    )
    od = raw.get("order_details", {}) or {}
    order_details = OrderDetails(
        order_reference=od.get("order_reference",""),
        additional_order_reference=od.get("additional_order_reference",""),
        end_user_purchase_order_reference=od.get("end_user_purchase_order_reference",""),
        promised_date=od.get("promised_date",""),
        comment=od.get("comment",""),
    )
    items = []
    for it in raw.get("items", []) or []:
        items.append(LineItem(
            name=str(it.get("name",""))[:128],
            quantity=_to_float(it.get("quantity",0)),
            unit_cost_price=_to_float(it.get("unit_cost_price",0)),
            tax=_to_float(it.get("tax",0)),
            total=_to_float(it.get("total",0)),
        ))
    t = raw.get("totals", {}) or {}
    totals = Totals(
        subtotal=_to_float(t.get("subtotal",0)),
        freight=_to_float(t.get("freight",0)),
        tax=_to_float(t.get("tax",0)),
        grand_total=_to_float(t.get("grand_total",0)),
    )
    return Invoice(
        invoice_number=raw.get("invoice_number","") or f"INV-{datetime.now():%Y%m%d%H%M%S}",
        invoice_date=raw.get("invoice_date","") or f"{datetime.now():%Y-%m-%d}",
        currency=raw.get("currency","AUD"),
        bill_to=bill_to,
        order_details=order_details,
        items=items,
        totals=totals
    )
def auto_normalize(raw: Dict[str, Any]) -> Tuple[Invoice, str]:
    if isinstance(raw, dict) and ("results" in raw or "order_reference" in raw):
        if "results" in raw:
            orders, _, _ = parse_vs_orders_payload(raw)
            inv = vs_order_to_invoice(orders[0]) if orders else normalize_custom_invoice_payload({})
            return inv, "VS"
        return vs_order_to_invoice(raw), "VS"
    return normalize_custom_invoice_payload(raw), "CUSTOM"
class InvoicePDF(FPDF):
    def __init__(self, ttf_path: Optional[str] = None):
        super().__init__()
        self._unicode_ok = False
        self._font_family = "Arial"
        candidates = []
        if ttf_path:
            candidates.append(ttf_path)
        candidates += [
            r"C:\Windows\Fonts\arial.ttf",
            r"C:\Windows\Fonts\Calibri.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
            "/Library/Fonts/Arial.ttf",
            "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
        ]
        for fp in candidates:
            if Path(fp).exists():
                try:
                    self.add_font("Fallback", "",   fp, uni=True)
                    self.add_font("Fallback", "B",  fp, uni=True)
                    self.add_font("Fallback", "I",  fp, uni=True)
                    self.add_font("Fallback", "BI", fp, uni=True)
                    self._font_family = "Fallback"
                    self._unicode_ok = True
                    break
                except Exception:
                    continue
        self.set_auto_page_break(auto=True, margin=15)
    def _safe(self, text: str) -> str:
        return text if self._unicode_ok else latin1_sanitize(text)
    def cell(self, w=0, h=0, txt="", border=0, ln=0, align="", fill=False, link=""):
        return super().cell(w, h, self._safe(txt), border, ln, align, fill, link)
    def multi_cell(self, w, h, txt="", border=0, align="J", fill=False):
        return super().multi_cell(w, h, self._safe(txt), border, align, fill)
    def header(self):
        self.set_font(self._font_family, "B", 14)
        self.cell(0, 8, "TAX INVOICE", ln=True, align="C")
        self.ln(2)
    def footer(self):
        self.set_y(-12)
        self.set_font(self._font_family, "I", 8)
        self.cell(0, 8, f"Page {self.page_no()}", align="C")
    def build(self, inv: Invoice, supplier_name="Supplier: XXX", supplier_abn="ABN: XX XXX XXX XXX"):
        self.add_page()
        self.set_font(self._font_family, "B", 10)
        self.cell(0, 6, f"Invoice Number: {inv.invoice_number}", ln=True)
        self.cell(0, 6, f"Invoice Date:   {inv.invoice_date}", ln=True)
        self.cell(0, 6, f"Currency:       {inv.currency}", ln=True)
        self.ln(2)
        self.set_font(self._font_family, "B", 10)
        self.cell(95, 6, "Bill From:", 0, 0)
        self.cell(95, 6, "Bill To:", 0, 1)
        self.set_font(self._font_family, "", 9)
        self.cell(95, 5, supplier_name, 0, 0)
        self.cell(95, 5, inv.bill_to.company_name, 0, 1)
        self.cell(95, 5, supplier_abn, 0, 0)
        self.cell(95, 5, inv.bill_to.address.line_1, 0, 1)
        self.cell(95, 5, "", 0, 0)
        self.cell(95, 5, f"{inv.bill_to.address.city}, {inv.bill_to.address.state} {inv.bill_to.address.postal_code}", 0, 1)
        self.cell(95, 5, "", 0, 0)
        self.cell(95, 5, inv.bill_to.phone, 0, 1)
        self.cell(95, 5, "", 0, 0)
        self.cell(95, 5, inv.bill_to.email, 0, 1)
        self.ln(2)
        self.set_font(self._font_family, "B", 10)
        self.cell(0, 6, "Order Details:", ln=True)
        self.set_font(self._font_family, "", 9)
        od = inv.order_details
        self.multi_cell(0, 5,
            f"Order Reference: {od.order_reference}\n"
            f"Additional Reference: {od.additional_order_reference}\n"
            f"End User PO Reference: {od.end_user_purchase_order_reference}\n"
            f"Promised Date: {od.promised_date}\n"
            f"Comment: {od.comment}"
        )
        self.ln(2)
        self.set_font(self._font_family, "B", 9)
        self.set_fill_color(220, 220, 220)
        self.cell(70, 7, "Item", 1, 0, "C", 1)
        self.cell(20, 7, "Qty", 1, 0, "C", 1)
        self.cell(30, 7, "Unit Price", 1, 0, "C", 1)
        self.cell(30, 7, "Tax", 1, 0, "C", 1)
        self.cell(30, 7, "Total", 1, 1, "C", 1)
        self.set_font(self._font_family, "", 9)
        for li in inv.items:
            self.cell(70, 7, (li.name or "")[:50], 1)
            self.cell(20, 7, f"{li.quantity:g}", 1, 0, "C")
            self.cell(30, 7, f"${li.unit_cost_price:,.2f}", 1, 0, "R")
            self.cell(30, 7, f"${li.tax:,.2f}", 1, 0, "R")
            self.cell(30, 7, f"${li.total:,.2f}", 1, 1, "R")
        self.ln(2)
        self.set_font(self._font_family, "B", 10)
        self.cell(150, 6, "", 0, 0)
        self.cell(30, 6, "Subtotal:", 0, 0, "R")
        self.cell(30, 6, f"${inv.totals.subtotal:,.2f}", 0, 1, "R")
        self.cell(150, 6, "", 0, 0)
        self.cell(30, 6, "Freight:", 0, 0, "R")
        self.cell(30, 6, f"${inv.totals.freight:,.2f}", 0, 1, "R")
        self.cell(150, 6, "", 0, 0)
        self.cell(30, 6, "Tax:", 0, 0, "R")
        self.cell(30, 6, f"${inv.totals.tax:,.2f}", 0, 1, "R")
        self.set_font(self._font_family, "B", 11)
        self.cell(150, 7, "", 0, 0)
        self.cell(30, 7, "Grand Total:", 0, 0, "R")
        self.cell(30, 7, f"${inv.totals.grand_total:,.2f}", 0, 1, "R")
class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Order → PDF Invoice (Notebook Portrait)")
        self.root.geometry("500x700")
        self.root.minsize(420, 500)
        default_font = ("Segoe UI", 9)
        self.root.option_add("*Font", default_font)
        self.root.option_add("*TCombobox*Listbox*Font", default_font)
        self.current_raw: Optional[Dict[str, Any]] = None
        self.current_invoice: Optional[Invoice] = None
        self.vs_orders: List[Dict[str, Any]] = []
        self.vs_next_url: Optional[str] = None
        self.vs_prev_url: Optional[str] = None
        self._build_ui()
    def _build_ui(self):
        self.nb = ttk.Notebook(self.root)
        self.nb.pack(fill="both", expand=True)
        self.tab_orders   = ttk.Frame(self.nb, padding=8)
        self.tab_json     = ttk.Frame(self.nb, padding=8)
        self.tab_invoice  = ttk.Frame(self.nb, padding=8)
        self.tab_items    = ttk.Frame(self.nb, padding=8)
        self.tab_export   = ttk.Frame(self.nb, padding=8)
        self.nb.add(self.tab_orders,  text="Orders")
        self.nb.add(self.tab_json,    text="JSON")
        self.nb.add(self.tab_invoice, text="Invoice")
        self.nb.add(self.tab_items,   text="Items")
        self.nb.add(self.tab_export,  text="Export")
        self._build_tab_orders()
        self._build_tab_json()
        self._build_tab_invoice()
        self._build_tab_items()
        self._build_tab_export()
        self.status_var = tk.StringVar(value="Ready.")
        status = ttk.Label(self.root, textvariable=self.status_var, anchor="w")
        status.pack(fill="x", padx=6, pady=(0,4))
    def _build_tab_orders(self):
        frm = self.tab_orders
        frm.columnconfigure(0, weight=1)
        ttk.Label(frm, text="API URL").grid(row=0, column=0, sticky="w")
        self.api_url_var = tk.StringVar(value="https://api.virtualstock.com/restapi/v4/orders/?limit=2&sort=desc&status=ORDER")
        ttk.Entry(frm, textvariable=self.api_url_var).grid(row=1, column=0, sticky="ew", pady=(0,6))
        auth_row = ttk.Frame(frm)
        auth_row.grid(row=2, column=0, sticky="ew", pady=(0,6))
        auth_row.columnconfigure(1, weight=1)
        ttk.Label(auth_row, text="Auth Key").grid(row=0, column=0, sticky="w")
        self.api_key_var = tk.StringVar()
        ttk.Entry(auth_row, textvariable=self.api_key_var, show="•").grid(row=0, column=1, sticky="ew", padx=(6,0))
        ttk.Label(auth_row, text="Timeout").grid(row=1, column=0, sticky="w", pady=(6,0))
        self.timeout_var = tk.StringVar(value="20")
        ttk.Entry(auth_row, textvariable=self.timeout_var, width=8).grid(row=1, column=1, sticky="w", padx=(6,0), pady=(6,0))
        btn_row = ttk.Frame(frm)
        btn_row.grid(row=3, column=0, sticky="ew", pady=(6,6))
        for i in range(6):
            btn_row.columnconfigure(i, weight=1)
        ttk.Button(btn_row, text="Fetch", command=self.on_fetch).grid(row=0, column=0, sticky="ew", padx=2)
        ttk.Button(btn_row, text="Load JSON", command=self.on_load_json).grid(row=0, column=1, sticky="ew", padx=2)
        ttk.Button(btn_row, text="Paste JSON", command=self.on_paste_json).grid(row=0, column=2, sticky="ew", padx=2)
        sel_row = ttk.Frame(frm)
        sel_row.grid(row=4, column=0, sticky="ew", pady=(6,0))
        sel_row.columnconfigure(1, weight=1)
        ttk.Label(sel_row, text="Order").grid(row=0, column=0, sticky="w")
        self.order_pick_var = tk.StringVar()
        self.order_pick = ttk.Combobox(sel_row, textvariable=self.order_pick_var, state="readonly")
        self.order_pick.grid(row=0, column=1, sticky="ew", padx=(6,6))
        self.order_pick.bind("<<ComboboxSelected>>", self.on_pick_order)
        self.btn_prev = ttk.Button(sel_row, text="◀", width=3, command=self.on_prev_page, state="disabled")
        self.btn_prev.grid(row=0, column=2, padx=2)
        self.btn_next = ttk.Button(sel_row, text="▶", width=3, command=self.on_next_page, state="disabled")
        self.btn_next.grid(row=0, column=3, padx=2)
        ttk.Label(frm, text="Tip: Fetch or Load/Paste JSON, pick an order, then move to other tabs.", foreground="#666")\
            .grid(row=5, column=0, sticky="w", pady=(8,0))
    def _build_tab_json(self):
        frm = self.tab_json
        frm.rowconfigure(0, weight=1)
        frm.columnconfigure(0, weight=1)
        self.json_text = tk.Text(frm, wrap="none", font=("Consolas", 10))
        self.json_text.grid(row=0, column=0, sticky="nsew")
        ys = ttk.Scrollbar(frm, orient="vertical", command=self.json_text.yview)
        xs = ttk.Scrollbar(frm, orient="horizontal", command=self.json_text.xview)
        self.json_text.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        ys.grid(row=0, column=1, sticky="ns")
        xs.grid(row=1, column=0, sticky="ew")
    def _build_tab_invoice(self):
        frm = self.tab_invoice
        for i in range(2):
            frm.columnconfigure(i, weight=1)
        ttk.Label(frm, text="Invoice Number").grid(row=0, column=0, sticky="w")
        self.invoice_number_var = tk.StringVar()
        ttk.Entry(frm, textvariable=self.invoice_number_var).grid(row=1, column=0, sticky="ew", pady=(0,6))
        ttk.Label(frm, text="Invoice Date (YYYY-MM-DD)").grid(row=0, column=1, sticky="w")
        self.invoice_date_var = tk.StringVar()
        ttk.Entry(frm, textvariable=self.invoice_date_var).grid(row=1, column=1, sticky="ew", pady=(0,6))
        ttk.Label(frm, text="Currency").grid(row=2, column=0, sticky="w")
        self.currency_var = tk.StringVar()
        ttk.Entry(frm, textvariable=self.currency_var).grid(row=3, column=0, sticky="ew", pady=(0,6))
        ttk.Label(frm, text="Order Reference").grid(row=2, column=1, sticky="w")
        self.order_ref_var = tk.StringVar()
        ttk.Entry(frm, textvariable=self.order_ref_var).grid(row=3, column=1, sticky="ew", pady=(0,6))
        ttk.Label(frm, text="Bill To (Company)").grid(row=4, column=0, sticky="w")
        self.bill_to_name_var = tk.StringVar()
        ttk.Entry(frm, textvariable=self.bill_to_name_var).grid(row=5, column=0, sticky="ew", pady=(0,6))
        ttk.Button(frm, text="Apply Header Edits", command=self._push_header_edits)\
            .grid(row=5, column=1, sticky="e")
    def _build_tab_items(self):
        frm = self.tab_items
        frm.rowconfigure(0, weight=1)
        frm.columnconfigure(0, weight=1)
        cols = ("name", "quantity", "unit_cost_price", "tax", "total")
        self.tree = ttk.Treeview(frm, columns=cols, show="headings")
        for c, w in zip(cols, (320, 70, 100, 90, 100)):
            self.tree.heading(c, text=c.replace("_"," ").title())
            self.tree.column(c, width=w, stretch=True)
        self.tree.grid(row=0, column=0, sticky="nsew")
        ys = ttk.Scrollbar(frm, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=ys.set)
        ys.grid(row=0, column=1, sticky="ns")
        ttk.Label(frm, text="Hint: Edit header details in the Invoice tab before exporting.", foreground="#666")\
            .grid(row=1, column=0, sticky="w", pady=(6,0))
    def _build_tab_export(self):
        frm = self.tab_export
        frm.columnconfigure(0, weight=1)
        row = 0
        ttk.Label(frm, text="Supplier Name").grid(row=row, column=0, sticky="w"); row += 1
        self.supplier_name_var = tk.StringVar(value="Supplier: XXX")
        ttk.Entry(frm, textvariable=self.supplier_name_var).grid(row=row, column=0, sticky="ew", pady=(0,6)); row += 1
        ttk.Label(frm, text="Supplier ABN").grid(row=row, column=0, sticky="w"); row += 1
        self.supplier_abn_var  = tk.StringVar(value="ABN: XX XXX XXX XXX")
        ttk.Entry(frm, textvariable=self.supplier_abn_var).grid(row=row, column=0, sticky="ew", pady=(0,6)); row += 1
        ttk.Label(frm, text="Optional TTF Font (Unicode)").grid(row=row, column=0, sticky="w"); row += 1
        ttf_row = ttk.Frame(frm)
        ttf_row.grid(row=row, column=0, sticky="ew", pady=(0,6)); row += 1
        ttf_row.columnconfigure(0, weight=1)
        self.ttf_path_var = tk.StringVar()
        ttk.Entry(ttf_row, textvariable=self.ttf_path_var).grid(row=0, column=0, sticky="ew")
        ttk.Button(ttf_row, text="Browse…", command=self.on_pick_ttf).grid(row=0, column=1, padx=(6,0))
        ttk.Button(frm, text="Generate PDF…", command=self.on_generate_pdf)\
            .grid(row=row, column=0, sticky="e", pady=(8,0)); row += 1
    def set_status(self, msg: str):
        self.status_var.set(msg)
        self.root.update_idletasks()
    def _update_from_invoice(self, inv: Invoice):
        self.invoice_number_var.set(inv.invoice_number)
        self.invoice_date_var.set(inv.invoice_date)
        self.currency_var.set(inv.currency)
        self.order_ref_var.set(inv.order_details.order_reference)
        self.bill_to_name_var.set(inv.bill_to.company_name)
        self.json_text.delete("1.0", "end")
        try:
            self.json_text.insert("1.0", json.dumps(self.current_raw or {}, indent=2))
        except Exception:
            self.json_text.insert("1.0", str(self.current_raw))
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        for li in inv.items:
            self.tree.insert("", "end", values=(
                li.name, f"{li.quantity:g}",
                f"{li.unit_cost_price:.2f}", f"{li.tax:.2f}", f"{li.total:.2f}"
            ))
    def _push_header_edits(self):
        if not self.current_invoice:
            return
        inv = self.current_invoice
        inv.invoice_number = (self.invoice_number_var.get().strip() or inv.invoice_number)
        inv.invoice_date   = (self.invoice_date_var.get().strip() or inv.invoice_date)
        inv.currency       = (self.currency_var.get().strip() or inv.currency)
        inv.order_details.order_reference = (self.order_ref_var.get().strip() or inv.order_details.order_reference)
        inv.bill_to.company_name = (self.bill_to_name_var.get().strip() or inv.bill_to.company_name)
    def _update_pager_buttons(self):
        self.btn_prev.config(state=("normal" if self.vs_prev_url else "disabled"))
        self.btn_next.config(state=("normal" if self.vs_next_url else "disabled"))
    def _load_vs_order(self, idx: int):
        try:
            order = self.vs_orders[idx]
        except Exception:
            return
        inv = vs_order_to_invoice(order)
        self.current_raw = order
        self.current_invoice = inv
        self._update_from_invoice(inv)
    def on_pick_ttf(self):
        fp = filedialog.askopenfilename(title="Pick .ttf font",
                                        filetypes=[("TrueType Font","*.ttf"), ("All files","*.*")])
        if fp:
            self.ttf_path_var.set(fp)
    def on_load_json(self):
        fp = filedialog.askopenfilename(title="Select JSON",
                                        filetypes=[("JSON files","*.json"),("All files","*.*")])
        if not fp: return
        try:
            with open(fp, "r", encoding="utf-8") as f:
                raw = json.load(f)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load JSON:\n{e}"); return
        self._ingest_raw(raw)
        self.set_status("Loaded JSON from file.")
    def on_paste_json(self):
        win = tk.Toplevel(self.root)
        win.title("Paste JSON"); win.geometry("800x520")
        txt = tk.Text(win, wrap="none", font=("Consolas", 10))
        ys = ttk.Scrollbar(win, orient="vertical", command=txt.yview)
        xs = ttk.Scrollbar(win, orient="horizontal", command=txt.xview)
        txt.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        txt.grid(row=0, column=0, sticky="nsew")
        ys.grid(row=0, column=1, sticky="ns")
        xs.grid(row=1, column=0, sticky="we")
        win.grid_rowconfigure(0, weight=1); win.grid_columnconfigure(0, weight=1)
        def ok():
            s = txt.get("1.0", "end").strip()
            try:
                raw = json.loads(s)
            except Exception as e:
                messagebox.showerror("Error", f"Invalid JSON:\n{e}"); return
            win.destroy(); self._ingest_raw(raw); self.set_status("Loaded JSON from paste.")
        ttk.Button(win, text="OK", command=ok).grid(row=2, column=0, sticky="e", padx=6, pady=6)
    def _ingest_raw(self, raw: Dict[str, Any]):
        self.current_raw = raw
        try:
            if isinstance(raw, dict) and "results" in raw:
                orders, nxt, prv = parse_vs_orders_payload(raw)
                self.vs_orders, self.vs_next_url, self.vs_prev_url = orders, nxt, prv
                refs = [(o.get("order_reference") or o.get("url","").rstrip("/").split("/")[-1]) for o in orders]
                self.order_pick["values"] = refs
                if refs:
                    self.order_pick.current(0); self._load_vs_order(0)
                self._update_pager_buttons()
                self.set_status("Loaded VS list JSON.")
            else:
                inv, _ = auto_normalize(raw)
                self.current_invoice = inv
                self._update_from_invoice(inv)
                self.order_pick["values"] = ()
                self.vs_next_url = self.vs_prev_url = None
                self._update_pager_buttons()
                self.set_status("Loaded single/custom JSON.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to parse payload:\n{e}")
    def on_fetch(self):
        if not HTTPX_OK:
            messagebox.showerror("Missing dependency", "httpx is not installed.\nRun: pip install httpx")
            return
        url = self.api_url_var.get().strip()
        if not url:
            messagebox.showwarning("Input required", "Provide API URL."); return
        try:
            timeout = float(self.timeout_var.get())
        except Exception:
            timeout = 20.0
        headers = {}
        if self.api_key_var.get().strip():
            headers["Authorization"] = f"Bearer {self.api_key_var.get().strip()}"
        self.set_status("Fetching…")
        def worker():
            try:
                with httpx.Client(timeout=timeout, follow_redirects=True, verify=True) as client:
                    r = client.get(url, headers=headers)
                    r.raise_for_status()
                    raw = r.json()
                self.root.after(0, self._after_fetch_ok, raw)
            except Exception as e:
                self.root.after(0, self._after_fetch_err, e)
    def _after_fetch_ok(self, raw):
        self._ingest_raw(raw); self.set_status("Fetch OK.")
    def _after_fetch_err(self, err):
        messagebox.showerror("Fetch failed", f"{err}"); self.set_status("Fetch failed.")
    def on_pick_order(self, _evt=None):
        sel = self.order_pick_var.get()
        for i, o in enumerate(self.vs_orders):
            ref = o.get("order_reference") or o.get("url","").rstrip("/").split("/")[-1]
            if ref == sel:
                self._load_vs_order(i); self.set_status(f"Selected {ref}")
                break
    def on_next_page(self):
        if not self.vs_next_url or not HTTPX_OK: return
        self._fetch_page(self.vs_next_url)
    def on_prev_page(self):
        if not self.vs_prev_url or not HTTPX_OK: return
        self._fetch_page(self.vs_prev_url)
    def _fetch_page(self, url: str):
        try:
            timeout = float(self.timeout_var.get())
        except Exception:
            timeout = 20.0
        headers = {}
        if self.api_key_var.get().strip():
            headers["Authorization"] = f"Bearer {self.api_key_var.get().strip()}"
        self.set_status("Fetching page…")
        def worker():
            try:
                with httpx.Client(timeout=timeout, follow_redirects=True, verify=True) as client:
                    r = client.get(url, headers=headers)
                    r.raise_for_status()
                    raw = r.json()
                self.root.after(0, lambda: self._after_fetch_ok(raw))
            except Exception as e:
                self.root.after(0, lambda: self._after_fetch_err(e))
        threading.Thread(target=worker, daemon=True).start()
    def on_generate_pdf(self):
        if not self.current_invoice:
            messagebox.showwarning("No data", "Load/paste/fetch an order first."); return
        self._push_header_edits()
        out = filedialog.asksaveasfilename(title="Save Invoice PDF",
                                           defaultextension=".pdf",
                                           filetypes=[("PDF files","*.pdf")],
                                           initialfile=f"{self.current_invoice.invoice_number}.pdf")
        if not out: return
        try:
            ttf = self.ttf_path_var.get().strip() or None
            pdf = InvoicePDF(ttf_path=ttf)
            pdf.build(
                self.current_invoice,
                supplier_name=self.supplier_name_var.get().strip() or "Supplier: XXX",
                supplier_abn=self.supplier_abn_var.get().strip() or "ABN: XX XXX XXX XXX",
            )
            pdf.output(out)
        except Exception as e:
            messagebox.showerror("Failed to generate PDF", f"{e}"); return
        self.set_status(f"Saved: {Path(out).name}")
        messagebox.showinfo("Success", f"Invoice generated:\n{out}")
if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
