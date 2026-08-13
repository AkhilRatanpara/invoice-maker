"use client";

import { useEffect, useMemo, useState } from "react";
import {
  Building2,
  Check,
  ChevronLeft,
  Database,
  Download,
  FileSpreadsheet,
  FileText,
  Lock,
  LogOut,
  Pencil,
  Plus,
  Printer,
  Save,
  Search,
  Settings,
  Shield,
  Trash2,
  UserPlus
} from "lucide-react";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import * as XLSX from "xlsx";
import {
  AlignmentType,
  BorderStyle,
  Document,
  HeadingLevel,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType
} from "docx";

type Business = {
  id: string;
  name: string;
  address: string;
  subtitle: string;
  phone: string;
  pin: string;
  template: "classic";
};

type Customer = {
  id: string;
  businessId: string;
  name: string;
  description: string;
  phone?: string;
};

type ItemPreset = {
  id: string;
  businessId: string;
  name: string;
  unit: string;
  rate: number;
  favorite: boolean;
  hidden?: boolean;
};

type InvoiceRow = {
  id: string;
  description: string;
  qty: number;
  unit: string;
  rate: number;
  amountOverride?: number;
};

type InvoiceGroup = {
  id: string;
  srNo: number;
  rows: InvoiceRow[];
};

type Invoice = {
  id: string;
  businessId: string;
  date: string;
  customerId?: string;
  customerName: string;
  customerDescription: string;
  groups: InvoiceGroup[];
  note: string;
  status: "draft" | "saved";
  lastEdited: string;
};

type AppData = {
  adminPin: string;
  businesses: Business[];
  customers: Customer[];
  presets: ItemPreset[];
  invoices: Invoice[];
};

const STORAGE_KEY = "madhav-invoice-maker-data-v1";
const SESSION_KEY = "madhav-invoice-maker-session-v1";

const today = () => new Date().toISOString().slice(0, 10);
const uid = () => Math.random().toString(36).slice(2, 10);
const money = (value: number) => Number.isFinite(value) ? value.toFixed(2) : "0.00";

const seedBusinessId = "madhav-electricals";

const seedData: AppData = {
  adminPin: "4321",
  businesses: [
    {
      id: seedBusinessId,
      name: "MADHAV ELECTRICALS",
      address:
        "58, Sagar Complex, At. Chhatral, Ta. Kalol, Dist. Gandhinagar-382729",
      subtitle: "(1 Phase, 3 Phase Motor Repairing & Rewinding)",
      phone: "99045 49013",
      pin: "1234",
      template: "classic"
    }
  ],
  customers: [
    {
      id: "raajratna-metal",
      businessId: seedBusinessId,
      name: "RAAJRATNA METAL INDUSTRIES Limited",
      description: "(Fine Wire Division, Bileshwarpura)"
    }
  ],
  presets: [
    preset("Rewinding CEILING FAN", 450, true),
    preset("2.5 MFD CAPACITOR", 50, true),
    preset("CAPACITOR CAMP", 10, true),
    preset("Rewinding 1-HP - 700 RPM TORQUE MOTOR", 1500, true),
    preset("Rewinding ALMONARD FAN", 1500, true),
    preset("4 MFD CAPACITOR", 70, false),
    preset("Rewinding 1-HP - 1440 RPM MOTOR", 1500, false),
    preset("Rewinding 2-HP - 1440 RPM MOTOR", 1920, false),
    preset("Bearing 6203", 163, false),
    preset("PVC Fan", 80, false),
    preset("Waterseal", 160, false),
    preset("Terminal Plate", 70, false)
  ],
  invoices: []
};

function preset(name: string, rate: number, favorite: boolean): ItemPreset {
  return {
    id: uid(),
    businessId: seedBusinessId,
    name,
    unit: "No",
    rate,
    favorite
  };
}

function newInvoice(businessId: string): Invoice {
  return {
    id: uid(),
    businessId,
    date: today(),
    customerName: "",
    customerDescription: "",
    groups: [
      {
        id: uid(),
        srNo: 1,
        rows: [emptyRow()]
      }
    ],
    note: "",
    status: "draft",
    lastEdited: new Date().toISOString()
  };
}

function emptyRow(): InvoiceRow {
  return {
    id: uid(),
    description: "",
    qty: 1,
    unit: "No",
    rate: 0
  };
}

function rowAmount(row: InvoiceRow) {
  return row.amountOverride ?? row.qty * row.rate;
}

function invoiceTotal(invoice: Invoice) {
  return invoice.groups.reduce(
    (sum, group) => sum + group.rows.reduce((rowSum, row) => rowSum + rowAmount(row), 0),
    0
  );
}

function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function customerFirstName(invoice: Invoice) {
  const first = invoice.customerName.trim().split(/\s+/)[0] || "Customer";
  return first.replace(/[^a-zA-Z0-9_-]/g, "");
}

function formattedDateForFile(date: string) {
  const [year, month, day] = date.split("-");
  return `${day}-${month}-${year}`;
}

function invoiceFilename(invoice: Invoice, extension: "pdf" | "xlsx" | "docx") {
  return `${customerFirstName(invoice)}_Invoice_${formattedDateForFile(invoice.date)}.${extension}`;
}

export default function Home() {
  const [data, setData] = useState<AppData>(seedData);
  const [ready, setReady] = useState(false);
  const [mode, setMode] = useState<"login" | "app" | "admin">("login");
  const [selectedBusinessId, setSelectedBusinessId] = useState(seedBusinessId);
  const [pin, setPin] = useState("");
  const [adminPin, setAdminPin] = useState("");
  const [pinError, setPinError] = useState("");
  const [activeInvoiceId, setActiveInvoiceId] = useState<string>("");
  const [lastSavedAt, setLastSavedAt] = useState("");
  const [activeStep, setActiveStep] = useState<"customer" | "items" | "preview">("customer");
  const [customerMode, setCustomerMode] = useState<"select" | "direct" | "new">("select");

  useEffect(() => {
    try {
      const stored = window.localStorage.getItem(STORAGE_KEY);
      const session = window.localStorage.getItem(SESSION_KEY);
      if (stored) {
        const parsedData = JSON.parse(stored) as AppData;
        if (parsedData && Array.isArray(parsedData.businesses) && parsedData.businesses.length > 0) {
          setData(parsedData);
        }
      }
      if (session) {
        const parsedSession = JSON.parse(session) as { businessId?: string; invoiceId?: string };
        if (parsedSession.businessId) {
          setSelectedBusinessId(parsedSession.businessId);
          setMode("app");
        }
        if (parsedSession.invoiceId) setActiveInvoiceId(parsedSession.invoiceId);
      }
    } catch (err) {
      console.warn("Storage parse error, resetting to seed data:", err);
      setData(seedData);
    } finally {
      setReady(true);
    }
  }, []);

  useEffect(() => {
    if (!ready) return;
    const timer = window.setTimeout(() => {
      window.localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
      setLastSavedAt(new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" }));
    }, 350);
    return () => window.clearTimeout(timer);
  }, [data, ready]);

  useEffect(() => {
    if (!ready || mode !== "app") return;
    window.localStorage.setItem(
      SESSION_KEY,
      JSON.stringify({ businessId: selectedBusinessId, invoiceId: activeInvoiceId })
    );
  }, [ready, mode, selectedBusinessId, activeInvoiceId]);

  const business = data.businesses?.find((item) => item.id === selectedBusinessId) ?? data.businesses?.[0] ?? seedData.businesses[0];
  const invoicesForBusiness = data.invoices?.filter((invoice) => invoice.businessId === business.id) ?? [];
  const activeInvoice =
    data.invoices?.find((invoice) => invoice.id === activeInvoiceId) ??
    invoicesForBusiness.find((invoice) => invoice.status === "draft") ??
    invoicesForBusiness[0] ??
    null;
  const customers = data.customers.filter((customer) => customer.businessId === business.id);
  const visiblePresets = data.presets.filter(
    (item) => item.businessId === business.id && !item.hidden
  );
  const favoritePresets = visiblePresets.filter((item) => item.favorite).slice(0, 8);

  useEffect(() => {
    if (!ready || mode !== "app") return;
    if (!activeInvoice) {
      const invoice = newInvoice(business.id);
      setData((current) => ({ ...current, invoices: [invoice, ...current.invoices] }));
      setActiveInvoiceId(invoice.id);
    } else if (!activeInvoiceId) {
      setActiveInvoiceId(activeInvoice.id);
    }
  }, [activeInvoice, activeInvoiceId, business.id, mode, ready]);

  const updateInvoice = (updater: (invoice: Invoice) => Invoice) => {
    if (!activeInvoice) return;
    setData((current) => ({
      ...current,
      invoices: current.invoices.map((invoice) =>
        invoice.id === activeInvoice.id
          ? { ...updater(invoice), lastEdited: new Date().toISOString() }
          : invoice
      )
    }));
  };

  const login = () => {
    if (pin === business.pin) {
      setMode("app");
      setPin("");
      setPinError("");
      return;
    }
    setPinError("Wrong PIN. Please ask admin.");
  };

  const adminLogin = () => {
    if (adminPin === data.adminPin) {
      setMode("admin");
      setAdminPin("");
      setPinError("");
      return;
    }
    setPinError("Wrong admin PIN.");
  };

  const logout = () => {
    window.localStorage.removeItem(SESSION_KEY);
    setMode("login");
    setPin("");
    setAdminPin("");
    setActiveInvoiceId("");
  };

  const createFreshInvoice = () => {
    const invoice = newInvoice(business.id);
    setData((current) => ({ ...current, invoices: [invoice, ...current.invoices] }));
    setActiveInvoiceId(invoice.id);
    setActiveStep("customer");
  };

  const deleteInvoice = (id: string) => {
    if (!confirm("Are you sure you want to delete this invoice draft?")) return;
    setData((current) => {
      const remaining = current.invoices.filter((inv) => inv.id !== id);
      const nextActive = activeInvoiceId === id
        ? remaining.find((inv) => inv.businessId === business.id)?.id || ""
        : activeInvoiceId;
      return { ...current, invoices: remaining };
    });
    if (activeInvoiceId === id) {
      const remaining = data.invoices.filter((inv) => inv.id !== id);
      const next = remaining.find((inv) => inv.businessId === business.id);
      if (next) {
        setActiveInvoiceId(next.id);
      } else {
        createFreshInvoice();
      }
    }
  };

  const applyPreset = (rowId: string, presetItem: ItemPreset) => {
    updateInvoice((invoice) => ({
      ...invoice,
      groups: invoice.groups.map((group) => ({
        ...group,
        rows: group.rows.map((row) =>
          row.id === rowId
            ? {
                ...row,
                description: presetItem.name,
                unit: presetItem.unit,
                rate: presetItem.rate,
                amountOverride: undefined
              }
            : row
        )
      }))
    }));
  };

  const saveCustomerFromInvoice = () => {
    if (!activeInvoice?.customerName.trim()) return;
    const customer: Customer = {
      id: uid(),
      businessId: business.id,
      name: activeInvoice.customerName.trim(),
      description: activeInvoice.customerDescription.trim()
    };
    setData((current) => ({ ...current, customers: [customer, ...current.customers] }));
    updateInvoice((invoice) => ({ ...invoice, customerId: customer.id }));
  };

  const addGroup = () => {
    updateInvoice((invoice) => ({
      ...invoice,
      groups: [
        ...invoice.groups,
        { id: uid(), srNo: invoice.groups.length + 1, rows: [emptyRow()] }
      ]
    }));
  };

  const addSubItem = (groupId: string) => {
    updateInvoice((invoice) => ({
      ...invoice,
      groups: invoice.groups.map((group) =>
        group.id === groupId ? { ...group, rows: [...group.rows, emptyRow()] } : group
      )
    }));
  };

  const deleteRow = (groupId: string, rowId: string) => {
    updateInvoice((invoice) => {
      const groups = invoice.groups
        .map((group) =>
          group.id === groupId
            ? { ...group, rows: group.rows.filter((row) => row.id !== rowId) }
            : group
        )
        .filter((group) => group.rows.length > 0)
        .map((group, index) => ({ ...group, srNo: index + 1 }));
      return { ...invoice, groups: groups.length ? groups : [{ id: uid(), srNo: 1, rows: [emptyRow()] }] };
    });
  };

  const exportPdf = () => {
    if (!activeInvoice) return;
    const doc = new jsPDF("p", "mm", "a4");
    const width = doc.internal.pageSize.getWidth();
    doc.setFont("helvetica", "bold");
    doc.setFontSize(20);
    doc.text(business.name, width / 2, 27, { align: "center" });
    doc.setFont("helvetica", "normal");
    doc.setFontSize(11);
    doc.text(`Address:- ${business.address}`, width / 2, 37, { align: "center" });
    doc.text(business.subtitle, width / 2, 46, { align: "center" });
    doc.text(`Mo:- ${business.phone}`, width / 2, 55, { align: "center" });
    doc.setLineWidth(0.7);
    doc.line(15, 66, 195, 66);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(16);
    doc.text("BILL / INVOICE", width / 2, 80, { align: "center" });
    doc.setFontSize(12);
    doc.text(`To: ${activeInvoice.customerName || "Customer Name"}`, 15, 96);
    const dateParts = activeInvoice.date.split("-");
    const displayDate = `${dateParts[2]}/${dateParts[1]}/${dateParts[0]}`;
    doc.text(`Date: ${displayDate}`, 195, 96, { align: "right" });
    doc.setFont("helvetica", "normal");
    if (activeInvoice.customerDescription) {
      doc.text(activeInvoice.customerDescription, 15, 105);
    }

    const body = activeInvoice.groups.flatMap((group) =>
      group.rows.map((row, rowIndex) => [
        rowIndex === 0 ? String(group.srNo) : "",
        row.description,
        `${row.qty || 0} ${row.unit}`,
        money(row.rate),
        money(rowAmount(row))
      ])
    );

    const tableStartY = activeInvoice.customerDescription ? 112 : 104;

    autoTable(doc, {
      startY: tableStartY,
      head: [["Sr. No.", "Particulars / Description", "Qty", "Rate (Rs.)", "Amount (Rs.)"]],
      body: [
        ...body,
        [
          {
            content: "TOTAL",
            colSpan: 4,
            styles: { halign: "right", fontStyle: "bold" }
          },
          { content: money(invoiceTotal(activeInvoice)), styles: { fontStyle: "bold" } }
        ]
      ],
      theme: "grid",
      styles: {
        font: "helvetica",
        fontSize: 11,
        cellPadding: 4,
        lineColor: [0, 0, 0],
        lineWidth: 0.25,
        textColor: [0, 0, 0]
      },
      headStyles: {
        fillColor: [255, 255, 255],
        textColor: [0, 0, 0],
        fontStyle: "bold",
        halign: "center",
        lineColor: [0, 0, 0],
        lineWidth: 0.25
      },
      columnStyles: {
        0: { cellWidth: 20, halign: "center" },
        1: { cellWidth: 88, halign: "left" },
        2: { cellWidth: 20, halign: "center" },
        3: { cellWidth: 24, halign: "right" },
        4: { cellWidth: 28, halign: "right" }
      },
      margin: { left: 15, right: 15 }
    });
    doc.save(invoiceFilename(activeInvoice, "pdf"));
  };

  const exportExcel = () => {
    if (!activeInvoice) return;
    const rows = activeInvoice.groups.flatMap((group) =>
      group.rows.map((row, rowIndex) => ({
        "Sr. No.": rowIndex === 0 ? group.srNo : "",
        "Particulars / Description": row.description,
        Qty: row.qty,
        Unit: row.unit,
        "Rate (Rs.)": row.rate,
        "Amount (Rs.)": rowAmount(row)
      }))
    );
    rows.push({
      "Sr. No.": "",
      "Particulars / Description": "TOTAL",
      Qty: "" as unknown as number,
      Unit: "",
      "Rate (Rs.)": "" as unknown as number,
      "Amount (Rs.)": invoiceTotal(activeInvoice)
    });
    const worksheet = XLSX.utils.json_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Invoice");
    XLSX.writeFile(workbook, invoiceFilename(activeInvoice, "xlsx"));
  };

  const exportWord = async () => {
    if (!activeInvoice) return;
    const border = { style: BorderStyle.SINGLE, size: 1, color: "000000" };
    const rows = [
      new TableRow({
        children: ["Sr. No.", "Particulars / Description", "Qty", "Rate (Rs.)", "Amount (Rs.)"].map(
          (text) =>
            new TableCell({
              children: [new Paragraph({ text, alignment: AlignmentType.CENTER })],
              borders: { top: border, bottom: border, left: border, right: border }
            })
        )
      }),
      ...activeInvoice.groups.flatMap((group) =>
        group.rows.map(
          (row, index) =>
            new TableRow({
              children: [
                index === 0 ? String(group.srNo) : "",
                row.description,
                `${row.qty} ${row.unit}`,
                money(row.rate),
                money(rowAmount(row))
              ].map(
                (text, cellIndex) =>
                  new TableCell({
                    children: [
                      new Paragraph({
                        text,
                        alignment: cellIndex === 1 ? AlignmentType.LEFT : AlignmentType.CENTER
                      })
                    ],
                    borders: { top: border, bottom: border, left: border, right: border }
                  })
              )
            })
        )
      ),
      new TableRow({
        children: [
          new TableCell({
            columnSpan: 4,
            children: [new Paragraph({ text: "TOTAL", alignment: AlignmentType.RIGHT })],
            borders: { top: border, bottom: border, left: border, right: border }
          }),
          new TableCell({
            children: [new Paragraph({ text: money(invoiceTotal(activeInvoice)), alignment: AlignmentType.CENTER })],
            borders: { top: border, bottom: border, left: border, right: border }
          })
        ]
      })
    ];
    const doc = new Document({
      sections: [
        {
          children: [
            new Paragraph({
              text: business.name,
              heading: HeadingLevel.TITLE,
              alignment: AlignmentType.CENTER
            }),
            new Paragraph({ text: `Address:- ${business.address}`, alignment: AlignmentType.CENTER }),
            new Paragraph({ text: business.subtitle, alignment: AlignmentType.CENTER }),
            new Paragraph({ text: `Mo:- ${business.phone}`, alignment: AlignmentType.CENTER }),
            new Paragraph({ text: "" }),
            new Paragraph({
              children: [new TextRun({ text: "BILL / INVOICE", bold: true, underline: {} })],
              alignment: AlignmentType.CENTER
            }),
            new Paragraph({ text: "" }),
            new Paragraph({
              children: [
                new TextRun({ text: `To: ${activeInvoice.customerName}`, bold: true }),
                new TextRun({ text: `          Date: ${activeInvoice.date}`, bold: true })
              ]
            }),
            new Paragraph({ text: activeInvoice.customerDescription }),
            new Paragraph({ text: "" }),
            new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows })
          ]
        }
      ]
    });
    const blob = await Packer.toBlob(doc);
    downloadBlob(blob, invoiceFilename(activeInvoice, "docx"));
  };

  if (!ready) {
    return <main className="center-screen">Loading invoice maker...</main>;
  }

  if (mode === "login") {
    return (
      <main className="login-shell">
        <section className="login-panel">
          <div className="brand-mark">
            <FileText size={36} />
          </div>
          <h1>Invoice Maker</h1>
          <p className="muted">Select business and enter 4 digit PIN.</p>

          <label className="field-label">Business</label>
          <select
            className="input"
            value={selectedBusinessId}
            onChange={(event) => setSelectedBusinessId(event.target.value)}
          >
            {data.businesses.map((item) => (
              <option key={item.id} value={item.id}>
                {item.name}
              </option>
            ))}
          </select>

          <label className="field-label">PIN</label>
          <input
            className="input pin-input"
            inputMode="numeric"
            maxLength={4}
            value={pin}
            onChange={(event) => setPin(event.target.value.replace(/\D/g, ""))}
            onKeyDown={(event) => event.key === "Enter" && login()}
            placeholder="0000"
          />
          {pinError && <p className="error">{pinError}</p>}
          <button className="primary wide" onClick={login}>
            <Lock size={18} />
            Open Business
          </button>

          <details className="admin-login">
            <summary>
              <Shield size={16} />
              Admin login
            </summary>
            <input
              className="input pin-input"
              inputMode="numeric"
              maxLength={4}
              value={adminPin}
              onChange={(event) => setAdminPin(event.target.value.replace(/\D/g, ""))}
              onKeyDown={(event) => event.key === "Enter" && adminLogin()}
              placeholder="Admin PIN"
            />
            <button className="secondary wide" onClick={adminLogin}>
              <Settings size={18} />
              Open Admin
            </button>
          </details>
        </section>
      </main>
    );
  }

  if (mode === "admin") {
    return (
      <AdminPanel
        data={data}
        setData={setData}
        selectedBusinessId={selectedBusinessId}
        setSelectedBusinessId={setSelectedBusinessId}
        logout={logout}
        lastSavedAt={lastSavedAt}
      />
    );
  }

  if (!activeInvoice) {
    return <main className="center-screen">Preparing invoice...</main>;
  }

  return (
    <main className="app-shell">
      <aside className="sidebar">
        <div>
          <div className="business-chip">
            <Building2 size={18} />
            <span>{business.name}</span>
          </div>
          <button className="primary wide" onClick={createFreshInvoice}>
            <Plus size={18} />
            New Invoice
          </button>
        </div>

        <nav className="step-list">
          {[
            ["customer", "Customer"],
            ["items", "Items"],
            ["preview", "Preview"]
          ].map(([id, label]) => (
            <button
              key={id}
              className={activeStep === id ? "step active" : "step"}
              onClick={() => setActiveStep(id as typeof activeStep)}
            >
              <span>{label}</span>
              {id === "preview" && <span>Rs. {money(invoiceTotal(activeInvoice))}</span>}
            </button>
          ))}
        </nav>

        <div className="saved-list">
          <p className="sidebar-title">Saved Drafts</p>
          {invoicesForBusiness.slice(0, 8).map((invoice) => (
            <div key={invoice.id} className="draft-row">
              <button
                className={invoice.id === activeInvoice.id ? "invoice-pill active" : "invoice-pill"}
                onClick={() => setActiveInvoiceId(invoice.id)}
              >
                <span>{invoice.customerName || "New customer"}</span>
                <small>{invoice.date}</small>
              </button>
              <button
                className="icon-button danger draft-delete"
                onClick={(e) => {
                  e.stopPropagation();
                  deleteInvoice(invoice.id);
                }}
                title="Delete Draft"
              >
                <Trash2 size={16} />
              </button>
            </div>
          ))}
        </div>

        <button className="ghost wide" onClick={logout}>
          <LogOut size={18} />
          Lock
        </button>
      </aside>

      <section className="workspace">
        <header className="topbar">
          <div>
            <p className="eyebrow">Auto saved {lastSavedAt || "now"}</p>
            <h2>{activeStep === "customer" ? "Customer Details" : activeStep === "items" ? "Add Work Items" : "Preview & Download"}</h2>
          </div>
          <div className="topbar-actions">
            {activeInvoice && (
              <button
                className="secondary danger-text"
                onClick={() => deleteInvoice(activeInvoice.id)}
                title="Delete current draft"
              >
                <Trash2 size={17} />
                Delete Draft
              </button>
            )}
            <button className="secondary" onClick={() => setMode("admin")}>
              <Settings size={18} />
              Admin
            </button>
          </div>
        </header>

        {activeStep === "customer" && (
          <section className="editor-panel">
            <div className="segmented">
              {[
                ["select", "Select"],
                ["direct", "Direct Fill"],
                ["new", "New Customer"]
              ].map(([id, label]) => (
                <button
                  key={id}
                  className={customerMode === id ? "active" : ""}
                  onClick={() => setCustomerMode(id as typeof customerMode)}
                >
                  {label}
                </button>
              ))}
            </div>

            {customerMode === "select" && (
              <div className="grid two">
                <label>
                  <span className="field-label">Saved customer</span>
                  <select
                    className="input"
                    value={activeInvoice.customerId ?? ""}
                    onChange={(event) => {
                      const customer = customers.find((item) => item.id === event.target.value);
                      if (!customer) return;
                      updateInvoice((invoice) => ({
                        ...invoice,
                        customerId: customer.id,
                        customerName: customer.name,
                        customerDescription: customer.description
                      }));
                    }}
                  >
                    <option value="">Choose customer</option>
                    {customers.map((customer) => (
                      <option key={customer.id} value={customer.id}>
                        {customer.name}
                      </option>
                    ))}
                  </select>
                </label>
                <label>
                  <span className="field-label">Invoice date</span>
                  <input
                    className="input"
                    type="date"
                    value={activeInvoice.date}
                    onChange={(event) =>
                      updateInvoice((invoice) => ({ ...invoice, date: event.target.value }))
                    }
                  />
                </label>
              </div>
            )}

            {(customerMode === "direct" || customerMode === "new") && (
              <div className="grid two">
                <label>
                  <span className="field-label">Customer name</span>
                  <input
                    className="input"
                    value={activeInvoice.customerName}
                    onChange={(event) =>
                      updateInvoice((invoice) => ({ ...invoice, customerName: event.target.value }))
                    }
                    placeholder="RAAJRATNA METAL INDUSTRIES Limited"
                  />
                </label>
                <label>
                  <span className="field-label">Invoice date</span>
                  <input
                    className="input"
                    type="date"
                    value={activeInvoice.date}
                    onChange={(event) =>
                      updateInvoice((invoice) => ({ ...invoice, date: event.target.value }))
                    }
                  />
                </label>
                <label className="span-two">
                  <span className="field-label">Customer description</span>
                  <input
                    className="input"
                    value={activeInvoice.customerDescription}
                    onChange={(event) =>
                      updateInvoice((invoice) => ({
                        ...invoice,
                        customerDescription: event.target.value
                      }))
                    }
                    placeholder="(Fine Wire Division, Bileshwarpura)"
                  />
                </label>
              </div>
            )}

            {customerMode === "new" && (
              <button className="secondary" onClick={saveCustomerFromInvoice}>
                <UserPlus size={18} />
                Save Customer
              </button>
            )}

            <button className="primary next-button" onClick={() => setActiveStep("items")}>
              Continue to Items
            </button>
          </section>
        )}

        {activeStep === "items" && (
          <section className="editor-panel item-editor">
            <div className="quick-row">
              {favoritePresets.map((item) => (
                <button
                  key={item.id}
                  className="quick-item"
                  onClick={() => {
                    const firstEmpty = activeInvoice.groups
                      .flatMap((group) => group.rows)
                      .find((row) => !row.description);
                    if (firstEmpty) {
                      applyPreset(firstEmpty.id, item);
                    } else {
                      updateInvoice((invoice) => ({
                        ...invoice,
                        groups: [
                          ...invoice.groups,
                          {
                            id: uid(),
                            srNo: invoice.groups.length + 1,
                            rows: [
                              {
                                id: uid(),
                                description: item.name,
                                qty: 1,
                                unit: item.unit,
                                rate: item.rate
                              }
                            ]
                          }
                        ]
                      }));
                    }
                  }}
                >
                  <Plus size={16} />
                  {item.name}
                </button>
              ))}
            </div>

            <datalist id="item-presets">
              {visiblePresets.map((item) => (
                <option key={item.id} value={item.name} />
              ))}
            </datalist>

            <div className="items-table">
              <div className="table-head">
                <span>Sr.</span>
                <span>Particulars / Description</span>
                <span>Qty</span>
                <span>Rate</span>
                <span>Total</span>
                <span />
              </div>
              {activeInvoice.groups.map((group) => (
                <div key={group.id} className="group-block">
                  {group.rows.map((row, rowIndex) => (
                    <div key={row.id} className="item-row">
                      <div className="sr-box">{rowIndex === 0 ? group.srNo : ""}</div>
                      <div className="description-cell">
                        <span className="mobile-label">Particulars / Description</span>
                        <div className="search-wrap">
                          <Search size={15} />
                          <input
                            list="item-presets"
                            value={row.description}
                            onBlur={() => {
                              const found = visiblePresets.find(
                                (item) => item.name.toLowerCase() === row.description.toLowerCase()
                              );
                              if (found) applyPreset(row.id, found);
                            }}
                            onChange={(event) =>
                              updateInvoice((invoice) => ({
                                ...invoice,
                                groups: invoice.groups.map((invoiceGroup) =>
                                  invoiceGroup.id === group.id
                                    ? {
                                        ...invoiceGroup,
                                        rows: invoiceGroup.rows.map((itemRow) =>
                                          itemRow.id === row.id
                                            ? { ...itemRow, description: event.target.value }
                                            : itemRow
                                        )
                                      }
                                    : invoiceGroup
                                )
                              }))
                            }
                            placeholder="Type or choose item"
                          />
                        </div>
                      </div>
                      <div className="qty-cell">
                        <span className="mobile-label">Qty & Unit</span>
                        <div className="qty-inputs">
                          <input
                            type="number"
                            min="0"
                            value={row.qty}
                            onChange={(event) =>
                              updateInvoice((invoice) => ({
                                ...invoice,
                                groups: invoice.groups.map((invoiceGroup) =>
                                  invoiceGroup.id === group.id
                                    ? {
                                        ...invoiceGroup,
                                        rows: invoiceGroup.rows.map((itemRow) =>
                                          itemRow.id === row.id
                                            ? { ...itemRow, qty: Number(event.target.value) }
                                            : itemRow
                                        )
                                      }
                                    : invoiceGroup
                                )
                              }))
                            }
                          />
                          <input
                            value={row.unit}
                            onChange={(event) =>
                              updateInvoice((invoice) => ({
                                ...invoice,
                                groups: invoice.groups.map((invoiceGroup) =>
                                  invoiceGroup.id === group.id
                                    ? {
                                        ...invoiceGroup,
                                        rows: invoiceGroup.rows.map((itemRow) =>
                                          itemRow.id === row.id
                                            ? { ...itemRow, unit: event.target.value }
                                            : itemRow
                                        )
                                      }
                                    : invoiceGroup
                                )
                              }))
                            }
                          />
                        </div>
                      </div>
                      <div className="rate-cell">
                        <span className="mobile-label">Rate (Rs.)</span>
                        <input
                          className="money-input"
                          type="number"
                          min="0"
                          value={row.rate}
                          onChange={(event) =>
                            updateInvoice((invoice) => ({
                              ...invoice,
                              groups: invoice.groups.map((invoiceGroup) =>
                                invoiceGroup.id === group.id
                                  ? {
                                      ...invoiceGroup,
                                      rows: invoiceGroup.rows.map((itemRow) =>
                                        itemRow.id === row.id
                                          ? {
                                              ...itemRow,
                                              rate: Number(event.target.value),
                                              amountOverride: undefined
                                            }
                                          : itemRow
                                      )
                                    }
                                  : invoiceGroup
                              )
                            }))
                          }
                        />
                      </div>
                      <div className="amount-cell">
                        <span className="mobile-label">Amount (Rs.)</span>
                        <input
                          className="money-input"
                          type="number"
                          min="0"
                          value={money(rowAmount(row))}
                          onChange={(event) =>
                            updateInvoice((invoice) => ({
                              ...invoice,
                              groups: invoice.groups.map((invoiceGroup) =>
                                invoiceGroup.id === group.id
                                  ? {
                                      ...invoiceGroup,
                                      rows: invoiceGroup.rows.map((itemRow) =>
                                        itemRow.id === row.id
                                          ? { ...itemRow, amountOverride: Number(event.target.value) }
                                          : itemRow
                                      )
                                    }
                                  : invoiceGroup
                              )
                            }))
                          }
                        />
                      </div>
                      <button className="icon-button danger" onClick={() => deleteRow(group.id, row.id)}>
                        <Trash2 size={17} />
                      </button>
                    </div>
                  ))}
                  <button className="small-action" onClick={() => addSubItem(group.id)}>
                    <Plus size={16} />
                    Add sub item under Sr. {group.srNo}
                  </button>
                </div>
              ))}
            </div>

            <div className="bottom-actions">
              <button className="secondary" onClick={addGroup}>
                <Plus size={18} />
                Add New Sr. No.
              </button>
              <div className="total-box">Total Rs. {money(invoiceTotal(activeInvoice))}</div>
              <button className="primary" onClick={() => setActiveStep("preview")}>
                Preview
              </button>
            </div>
          </section>
        )}

        {activeStep === "preview" && (
          <section className="preview-layout">
            <InvoicePreview business={business} invoice={activeInvoice} />
            <div className="download-panel">
              <button className="primary wide" onClick={exportPdf}>
                <Printer size={18} />
                Download PDF
              </button>
              <button className="secondary wide" onClick={exportExcel}>
                <FileSpreadsheet size={18} />
                Download Excel
              </button>
              <button className="secondary wide" onClick={exportWord}>
                <FileText size={18} />
                Download Word
              </button>
              <button className="ghost wide" onClick={() => setActiveStep("items")}>
                <ChevronLeft size={18} />
                Edit Items
              </button>
            </div>
          </section>
        )}
      </section>
    </main>
  );
}

function InvoicePreview({ business, invoice }: { business: Business; invoice: Invoice }) {
  const dateParts = invoice.date.split("-");
  const displayDate = `${dateParts[2]}/${dateParts[1]}/${dateParts[0]}`;
  return (
    <article className="invoice-paper">
      <header className="invoice-header">
        <h1>{business.name}</h1>
        <p>Address:- {business.address}</p>
        <p>{business.subtitle}</p>
        <p>Mo:- {business.phone}</p>
      </header>
      <div className="invoice-line" />
      <h2>BILL / INVOICE</h2>
      <div className="invoice-to-row">
        <div>
          <strong>To: {invoice.customerName || "Customer Name"}</strong>
          <p>{invoice.customerDescription}</p>
        </div>
        <strong>Date: {displayDate}</strong>
      </div>
      <table className="invoice-table">
        <thead>
          <tr>
            <th>Sr. No.</th>
            <th>Particulars / Description</th>
            <th>Qty</th>
            <th>Rate (Rs.)</th>
            <th>Amount (Rs.)</th>
          </tr>
        </thead>
        <tbody>
          {invoice.groups.map((group) =>
            group.rows.map((row, index) => (
              <tr key={row.id}>
                <td>{index === 0 ? group.srNo : ""}</td>
                <td>{row.description}</td>
                <td>{row.qty} {row.unit}</td>
                <td>{money(row.rate)}</td>
                <td>{money(rowAmount(row))}</td>
              </tr>
            ))
          )}
          <tr className="invoice-total">
            <td colSpan={4}>TOTAL</td>
            <td>{money(invoiceTotal(invoice))}</td>
          </tr>
        </tbody>
      </table>
    </article>
  );
}

function AdminPanel({
  data,
  setData,
  selectedBusinessId,
  setSelectedBusinessId,
  logout,
  lastSavedAt
}: {
  data: AppData;
  setData: React.Dispatch<React.SetStateAction<AppData>>;
  selectedBusinessId: string;
  setSelectedBusinessId: (id: string) => void;
  logout: () => void;
  lastSavedAt: string;
}) {
  const [tab, setTab] = useState<"business" | "items" | "customers" | "database">("business");
  const [dbStatus, setDbStatus] = useState<string>("Local Storage Mode (Offline / Free)");
  const [isSyncing, setIsSyncing] = useState<boolean>(false);
  const business = data.businesses.find((item) => item.id === selectedBusinessId) ?? data.businesses[0];
  const items = data.presets.filter((item) => item.businessId === business.id);
  const customers = data.customers.filter((customer) => customer.businessId === business.id);

  const checkDbConnection = async () => {
    setIsSyncing(true);
    try {
      const res = await fetch("/api/db");
      const result = await res.json();
      if (result.connected) {
        setDbStatus("🟢 Connected to Neon Postgres Database!");
      } else {
        setDbStatus(`🟡 Local Storage Mode: ${result.message || "DATABASE_URL not set"}`);
      }
    } catch {
      setDbStatus("🟡 Local Storage Mode (Offline)");
    } finally {
      setIsSyncing(false);
    }
  };

  const syncToNeon = async () => {
    setIsSyncing(true);
    try {
      const res = await fetch("/api/db", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(data)
      });
      const result = await res.json();
      if (result.success) {
        setDbStatus("✅ Sync Successful! All data saved in Neon Database.");
        alert("Data successfully synced to Neon Database!");
      } else {
        alert(result.message || "DATABASE_URL is not set on Vercel yet. Saving locally.");
      }
    } catch (err) {
      alert("Sync failed. Check connection or DATABASE_URL setting.");
    } finally {
      setIsSyncing(false);
    }
  };

  const updateBusiness = (patch: Partial<Business>) => {
    setData((current) => ({
      ...current,
      businesses: current.businesses.map((item) =>
        item.id === business.id ? { ...item, ...patch } : item
      )
    }));
  };

  const addBusiness = () => {
    const id = uid();
    setData((current) => ({
      ...current,
      businesses: [
        ...current.businesses,
        {
          id,
          name: "NEW BUSINESS",
          address: "",
          subtitle: "",
          phone: "",
          pin: "0000",
          template: "classic"
        }
      ]
    }));
    setSelectedBusinessId(id);
  };

  const addPreset = () => {
    setData((current) => ({
      ...current,
      presets: [
        {
          id: uid(),
          businessId: business.id,
          name: "New item",
          unit: "No",
          rate: 0,
          favorite: false
        },
        ...current.presets
      ]
    }));
  };

  const updatePreset = (id: string, patch: Partial<ItemPreset>) => {
    setData((current) => ({
      ...current,
      presets: current.presets.map((item) => (item.id === id ? { ...item, ...patch } : item))
    }));
  };

  const deletePreset = (id: string) => {
    setData((current) => ({
      ...current,
      presets: current.presets.filter((item) => item.id !== id)
    }));
  };

  const exportBackup = () => {
    downloadBlob(
      new Blob([JSON.stringify(data, null, 2)], { type: "application/json" }),
      `invoice-backup-${today()}.json`
    );
  };

  return (
    <main className="admin-shell">
      <aside className="sidebar">
        <div className="business-chip">
          <Shield size={18} />
          <span>Admin</span>
        </div>
        <select
          className="input"
          value={selectedBusinessId}
          onChange={(event) => setSelectedBusinessId(event.target.value)}
        >
          {data.businesses.map((item) => (
            <option key={item.id} value={item.id}>
              {item.name}
            </option>
          ))}
        </select>
        <button className="secondary wide" onClick={addBusiness}>
          <Plus size={18} />
          Add Business
        </button>
        <nav className="step-list">
          {[
            ["business", "Business"],
            ["items", "Items"],
            ["customers", "Customers"],
            ["database", "Database"]
          ].map(([id, label]) => (
            <button
              key={id}
              className={tab === id ? "step active" : "step"}
              onClick={() => setTab(id as typeof tab)}
            >
              {label}
            </button>
          ))}
        </nav>
        <button className="ghost wide" onClick={logout}>
          <LogOut size={18} />
          Lock
        </button>
      </aside>

      <section className="workspace">
        <header className="topbar">
          <div>
            <p className="eyebrow">Auto saved {lastSavedAt || "now"}</p>
            <h2>Admin Portal</h2>
          </div>
          <button className="secondary" onClick={exportBackup}>
            <Download size={18} />
            Backup
          </button>
        </header>

        {tab === "business" && (
          <section className="editor-panel">
            <div className="grid two">
              <label>
                <span className="field-label">Business name</span>
                <input
                  className="input"
                  value={business.name}
                  onChange={(event) => updateBusiness({ name: event.target.value })}
                />
              </label>
              <label>
                <span className="field-label">4 digit PIN</span>
                <input
                  className="input pin-input"
                  inputMode="numeric"
                  maxLength={4}
                  value={business.pin}
                  onChange={(event) => updateBusiness({ pin: event.target.value.replace(/\D/g, "") })}
                />
              </label>
              <label className="span-two">
                <span className="field-label">Address</span>
                <input
                  className="input"
                  value={business.address}
                  onChange={(event) => updateBusiness({ address: event.target.value })}
                />
              </label>
              <label>
                <span className="field-label">Business line</span>
                <input
                  className="input"
                  value={business.subtitle}
                  onChange={(event) => updateBusiness({ subtitle: event.target.value })}
                />
              </label>
              <label>
                <span className="field-label">Phone</span>
                <input
                  className="input"
                  value={business.phone}
                  onChange={(event) => updateBusiness({ phone: event.target.value })}
                />
              </label>
              <label>
                <span className="field-label">Admin PIN</span>
                <input
                  className="input pin-input"
                  inputMode="numeric"
                  maxLength={4}
                  value={data.adminPin}
                  onChange={(event) =>
                    setData((current) => ({
                      ...current,
                      adminPin: event.target.value.replace(/\D/g, "")
                    }))
                  }
                />
              </label>
            </div>
          </section>
        )}

        {tab === "items" && (
          <section className="editor-panel">
            <div className="panel-toolbar">
              <h3>Items & Rates</h3>
              <button className="primary" onClick={addPreset}>
                <Plus size={18} />
                Add Item
              </button>
            </div>
            <div className="admin-list">
              {items.map((item) => (
                <div key={item.id} className="admin-row">
                  <input
                    className="input"
                    value={item.name}
                    onChange={(event) => updatePreset(item.id, { name: event.target.value })}
                  />
                  <input
                    className="input compact"
                    value={item.unit}
                    onChange={(event) => updatePreset(item.id, { unit: event.target.value })}
                  />
                  <input
                    className="input compact"
                    type="number"
                    value={item.rate}
                    onChange={(event) => updatePreset(item.id, { rate: Number(event.target.value) })}
                  />
                  <button
                    className={item.favorite ? "icon-button selected" : "icon-button"}
                    onClick={() => updatePreset(item.id, { favorite: !item.favorite })}
                    title="Show as quick item"
                  >
                    <Check size={17} />
                  </button>
                  <button className="icon-button danger" onClick={() => deletePreset(item.id)}>
                    <Trash2 size={17} />
                  </button>
                </div>
              ))}
            </div>
          </section>
        )}

        {tab === "customers" && (
          <section className="editor-panel">
            <h3>Saved Customers</h3>
            <div className="admin-list">
              {customers.map((customer) => (
                <div key={customer.id} className="customer-card">
                  <strong>{customer.name}</strong>
                  <span>{customer.description}</span>
                </div>
              ))}
            </div>
          </section>
        )}

        {tab === "database" && (
          <section className="editor-panel database-card">
            <Database size={32} />
            <h3>Database & Multi-User Sync</h3>
            <div className="status-badge">
              <strong>Status:</strong> {dbStatus}
            </div>

            <div className="db-actions">
              <button className="secondary" onClick={checkDbConnection} disabled={isSyncing}>
                Check Connection
              </button>
              <button className="primary" onClick={syncToNeon} disabled={isSyncing}>
                Sync Now to Neon
              </button>
            </div>

            <div className="db-info">
              <h4>How Neon Database Works:</h4>
              <p>
                By default, your app runs <strong>100% free</strong> in browser LocalStorage. 
                If you want 2-3 users on different phones/computers to share the exact same customers and invoices:
              </p>
              <ol>
                <li>Create a free account at <a href="https://neon.tech" target="_blank" rel="noreferrer">neon.tech</a> (Free Postgres).</li>
                <li>Copy your Postgres Connection String (`postgresql://...`).</li>
                <li>Add variable `DATABASE_URL` in your <strong>Vercel Project Settings -&gt; Environment Variables</strong>.</li>
                <li>Your app will automatically detect Neon and sync all invoices across all devices!</li>
              </ol>
            </div>
          </section>
        )}
      </section>
    </main>
  );
}
