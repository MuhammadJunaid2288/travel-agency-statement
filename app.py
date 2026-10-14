#!/usr/bin/env python3
"""
Travel Agency - Multi-Agent Account Statement (Web Version)
Streamlit Web Application
"""

import streamlit as st
import sqlite3
from datetime import datetime, date
from pathlib import Path
import tempfile
import os

# Optional: PDF & Excel
try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, HRFlowable
    from reportlab.lib.enums import TA_CENTER
    REPORTLAB_OK = True
except ImportError:
    REPORTLAB_OK = False

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    from openpyxl.utils import get_column_letter
    OPENPYXL_OK = True
except ImportError:
    OPENPYXL_OK = False

# ==================== DATABASE ====================
DB_PATH = Path(__file__).parent / "travel_agency.db"

def get_conn():
    conn = sqlite3.connect(str(DB_PATH), check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS agents (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            agent_name TEXT NOT NULL,
            agency_city TEXT,
            contact_phone TEXT,
            statement_date TEXT,
            statement_no TEXT,
            is_active INTEGER DEFAULT 1
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS ledger (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            agent_id INTEGER NOT NULL,
            date TEXT,
            pnr_ref TEXT,
            category TEXT,
            passenger_name TEXT,
            sector_desc TEXT,
            qty REAL DEFAULT 0,
            unit_price REAL DEFAULT 0,
            total_amount REAL DEFAULT 0,
            payment_status TEXT,
            remarks TEXT
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS payments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            agent_id INTEGER NOT NULL,
            payment_date TEXT,
            receipt_ref TEXT,
            payment_mode TEXT,
            bank_details TEXT,
            amount REAL DEFAULT 0,
            remarks TEXT
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS app_settings (
            key TEXT PRIMARY KEY,
            value TEXT
        )
    """)
    cur.execute("SELECT COUNT(*) FROM agents")
    if cur.fetchone()[0] == 0:
        cur.execute("""
            INSERT INTO agents (agent_name, agency_city, contact_phone, statement_date, statement_no)
            VALUES (?, ?, ?, ?, ?)
        """, ("Zubair Dayu (Sub-Agent)", "Dayu Travels - Larkana", "+92 300 1234567",
              datetime.now().strftime("%d-%b-%Y"), "STM-2026-089"))
        aid = cur.lastrowid
        cur.execute("INSERT OR REPLACE INTO app_settings (key, value) VALUES ('current_agent_id', ?)", (str(aid),))
    conn.commit()
    conn.close()

def get_setting(key, default=None):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT value FROM app_settings WHERE key=?", (key,))
    row = cur.fetchone()
    conn.close()
    return row["value"] if row else default

def set_setting(key, value):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("INSERT OR REPLACE INTO app_settings (key, value) VALUES (?, ?)", (key, value))
    conn.commit()
    conn.close()

def get_all_agents():
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT * FROM agents WHERE is_active=1 ORDER BY agent_name")
    rows = cur.fetchall()
    conn.close()
    return rows

def get_agent(agent_id):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT * FROM agents WHERE id=?", (agent_id,))
    row = cur.fetchone()
    conn.close()
    return row

def add_agent(name, agency, phone, stmt_date, stmt_no):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO agents (agent_name, agency_city, contact_phone, statement_date, statement_no)
        VALUES (?, ?, ?, ?, ?)
    """, (name, agency, phone, stmt_date, stmt_no))
    aid = cur.lastrowid
    conn.commit()
    conn.close()
    return aid

def update_agent(agent_id, name, agency, phone, stmt_date, stmt_no):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        UPDATE agents SET agent_name=?, agency_city=?, contact_phone=?, statement_date=?, statement_no=?
        WHERE id=?
    """, (name, agency, phone, stmt_date, stmt_no, agent_id))
    conn.commit()
    conn.close()

def delete_agent(agent_id):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM ledger WHERE agent_id=?", (agent_id,))
    cur.execute("DELETE FROM payments WHERE agent_id=?", (agent_id,))
    cur.execute("UPDATE agents SET is_active=0 WHERE id=?", (agent_id,))
    conn.commit()
    conn.close()

def get_ledger(agent_id):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT * FROM ledger WHERE agent_id=? ORDER BY date, id", (agent_id,))
    rows = cur.fetchall()
    conn.close()
    return rows

def add_ledger(agent_id, data):
    total = float(data.get("qty", 0)) * float(data.get("unit_price", 0))
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO ledger (agent_id, date, pnr_ref, category, passenger_name, sector_desc,
                            qty, unit_price, total_amount, payment_status, remarks)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (agent_id, data["date"], data["pnr_ref"], data["category"], data["passenger_name"],
          data["sector_desc"], data["qty"], data["unit_price"], total, data["payment_status"], data["remarks"]))
    conn.commit()
    conn.close()

def update_ledger(row_id, data):
    total = float(data.get("qty", 0)) * float(data.get("unit_price", 0))
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        UPDATE ledger SET date=?, pnr_ref=?, category=?, passenger_name=?, sector_desc=?,
        qty=?, unit_price=?, total_amount=?, payment_status=?, remarks=? WHERE id=?
    """, (data["date"], data["pnr_ref"], data["category"], data["passenger_name"],
          data["sector_desc"], data["qty"], data["unit_price"], total,
          data["payment_status"], data["remarks"], row_id))
    conn.commit()
    conn.close()

def delete_ledger(row_id):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM ledger WHERE id=?", (row_id,))
    conn.commit()
    conn.close()

def get_payments(agent_id):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT * FROM payments WHERE agent_id=? ORDER BY payment_date, id", (agent_id,))
    rows = cur.fetchall()
    conn.close()
    return rows

def add_payment(agent_id, data):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO payments (agent_id, payment_date, receipt_ref, payment_mode, bank_details, amount, remarks)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """, (agent_id, data["payment_date"], data["receipt_ref"], data["payment_mode"],
          data["bank_details"], data["amount"], data["remarks"]))
    conn.commit()
    conn.close()

def update_payment(row_id, data):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        UPDATE payments SET payment_date=?, receipt_ref=?, payment_mode=?, bank_details=?, amount=?, remarks=?
        WHERE id=?
    """, (data["payment_date"], data["receipt_ref"], data["payment_mode"],
          data["bank_details"], data["amount"], data["remarks"], row_id))
    conn.commit()
    conn.close()

def delete_payment(row_id):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM payments WHERE id=?", (row_id,))
    conn.commit()
    conn.close()

def get_summary(agent_id):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*), COALESCE(SUM(total_amount),0) FROM ledger WHERE agent_id=? AND category='Ticket'", (agent_id,))
    t_count, t_amount = cur.fetchone()
    cur.execute("SELECT COUNT(*), COALESCE(SUM(total_amount),0) FROM ledger WHERE agent_id=? AND category='Package'", (agent_id,))
    p_count, p_amount = cur.fetchone()
    cur.execute("SELECT COUNT(*), COALESCE(SUM(total_amount),0) FROM ledger WHERE agent_id=? AND category='Visa/Hotel'", (agent_id,))
    v_count, v_amount = cur.fetchone()
    cur.execute("SELECT COALESCE(SUM(total_amount),0) FROM ledger WHERE agent_id=?", (agent_id,))
    gross = cur.fetchone()[0]
    cur.execute("SELECT COALESCE(SUM(amount),0) FROM payments WHERE agent_id=?", (agent_id,))
    payments = cur.fetchone()[0]
    conn.close()
    return {
        "tickets_count": t_count or 0, "tickets_amount": t_amount or 0,
        "packages_count": p_count or 0, "packages_amount": p_amount or 0,
        "visa_count": v_count or 0, "visa_amount": v_amount or 0,
        "gross": gross or 0, "payments": payments or 0,
        "outstanding": (gross or 0) - (payments or 0)
    }

def clear_agent_data(agent_id):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM ledger WHERE agent_id=?", (agent_id,))
    cur.execute("DELETE FROM payments WHERE agent_id=?", (agent_id,))
    conn.commit()
    conn.close()

# ==================== PDF ====================
def generate_pdf(agent_id, path):
    agent = get_agent(agent_id)
    summary = get_summary(agent_id)
    ledger = get_ledger(agent_id)
    payments = get_payments(agent_id)

    doc = SimpleDocTemplate(path, pagesize=A4,
                            rightMargin=15*mm, leftMargin=15*mm, topMargin=12*mm, bottomMargin=12*mm)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('T', parent=styles['Heading1'], fontSize=15, alignment=TA_CENTER,
                                 textColor=colors.HexColor('#1a365d'), spaceAfter=3, fontName='Helvetica-Bold')
    subtitle_style = ParagraphStyle('S', parent=styles['Normal'], fontSize=9, alignment=TA_CENTER,
                                    textColor=colors.HexColor('#718096'), spaceAfter=8)
    small = ParagraphStyle('Sm', parent=styles['Normal'], fontSize=8)

    elements = []
    elements.append(Paragraph("TRAVEL AGENCY - SUB-AGENT ACCOUNT STATEMENT", title_style))
    elements.append(Paragraph(f"Agent: {agent['agent_name']}  |  {agent['agency_city']}", subtitle_style))
    elements.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#1a365d'), spaceAfter=8))

    agent_data = [
        [Paragraph("<b>AGENT INFORMATION</b>", ParagraphStyle('h', fontSize=9, textColor=colors.white))],
        [Paragraph(f"<b>Agent Name:</b>  {agent['agent_name']}", small)],
        [Paragraph(f"<b>Agency / City:</b>  {agent['agency_city']}", small)],
        [Paragraph(f"<b>Contact / Phone:</b>  {agent['contact_phone']}", small)],
        [Paragraph(f"<b>Statement Date:</b>  {agent['statement_date']}", small)],
        [Paragraph(f"<b>Statement No:</b>  {agent['statement_no']}", small)],
    ]
    finance_data = [
        [Paragraph("<b>FINANCIAL SUMMARY OVERVIEW</b>", ParagraphStyle('h', fontSize=9, textColor=colors.white))],
        [Paragraph(f"<b>Total Tickets Amount:</b>  Rs. {summary['tickets_amount']:,.0f}", small)],
        [Paragraph(f"<b>Total Packages Amount:</b>  Rs. {summary['packages_amount']:,.0f}", small)],
        [Paragraph(f"<b>Total Gross Booking Amount:</b>  Rs. {summary['gross']:,.0f}", small)],
        [Paragraph(f"<b>Total Payments Received:</b>  Rs. {summary['payments']:,.0f}", small)],
        [Paragraph(f"<b>NET OUTSTANDING BALANCE:</b>  <font color='#c53030'><b>Rs. {summary['outstanding']:,.0f}</b></font>", small)],
    ]
    t1 = Table(agent_data, colWidths=[90*mm])
    t1.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a365d')),
        ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#ebf8ff')),
        ('BOX', (0, 0), (-1, -1), 0.5, colors.HexColor('#2b6cb0')),
        ('TOPPADDING', (0, 0), (-1, -1), 3), ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('LEFTPADDING', (0, 0), (-1, -1), 5),
    ]))
    t2 = Table(finance_data, colWidths=[90*mm])
    t2.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a365d')),
        ('BACKGROUND', (0, 1), (-1, -2), colors.HexColor('#f0fff4')),
        ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#fefcbf')),
        ('BOX', (0, 0), (-1, -1), 0.5, colors.HexColor('#2b6cb0')),
        ('TOPPADDING', (0, 0), (-1, -1), 3), ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('LEFTPADDING', (0, 0), (-1, -1), 5),
    ]))
    side = Table([[t1, t2]], colWidths=[95*mm, 95*mm])
    side.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'TOP')]))
    elements.append(side)
    elements.append(Spacer(1, 8))

    elements.append(Paragraph("<b>SUMMARY BY SERVICE TYPE</b>", ParagraphStyle(
        'sec', fontSize=10, textColor=colors.HexColor('#1a365d'), spaceBefore=3, spaceAfter=3)))
    svc = [['Service Category', 'Count', 'Total Amount (PKR)', 'Agent Balance (PKR)'],
           ['Airline Tickets', str(summary['tickets_count']), f"{summary['tickets_amount']:,.0f}", f"{summary['tickets_amount']:,.0f}"],
           ['Umrah / Holiday Packages', str(summary['packages_count']), f"{summary['packages_amount']:,.0f}", f"{summary['packages_amount']:,.0f}"],
           ['Hotel & Visa Services', str(summary['visa_count']), f"{summary['visa_amount']:,.0f}", f"{summary['visa_amount']:,.0f}"],
           ['TOTAL ALL SERVICES', str(summary['tickets_count']+summary['packages_count']+summary['visa_count']),
            f"{summary['gross']:,.0f}", f"{summary['gross']:,.0f}"]]
    st_table = Table(svc, colWidths=[55*mm, 25*mm, 45*mm, 45*mm])
    st_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2b6cb0')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('ALIGN', (1, 0), (-1, -1), 'CENTER'),
        ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#edf2f7')),
        ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
        ('GRID', (0, 0), (-1, -1), 0.4, colors.HexColor('#cbd5e0')),
        ('TOPPADDING', (0, 0), (-1, -1), 3), ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ]))
    elements.append(st_table)
    elements.append(Spacer(1, 10))

    elements.append(Paragraph("<b>DETAILED TICKETS & PACKAGES LEDGER</b>", ParagraphStyle(
        'sec', fontSize=10, textColor=colors.HexColor('#1a365d'), spaceBefore=3, spaceAfter=3)))
    led = [['Sr', 'Date', 'PNR/Ref', 'Category', 'Passenger', 'Sector', 'Qty', 'Unit Price', 'Total', 'Status']]
    for i, r in enumerate(ledger, 1):
        led.append([str(i), str(r['date'] or '')[:10], str(r['pnr_ref'] or '')[:12],
                    str(r['category'] or ''), str(r['passenger_name'] or '')[:16],
                    str(r['sector_desc'] or '')[:14], str(int(r['qty']) if r['qty'] else 0),
                    f"{r['unit_price']:,.0f}", f"{r['total_amount']:,.0f}", str(r['payment_status'] or '')])
    if not ledger:
        led.append(['--', 'No bookings', '', '', '', '', '', '', '', ''])
    led.append(['', '', '', '', '', '', '', 'TOTAL', f"{summary['gross']:,.0f}", ''])
    lt = Table(led, colWidths=[8*mm, 17*mm, 20*mm, 17*mm, 30*mm, 26*mm, 10*mm, 20*mm, 20*mm, 16*mm])
    ls = [('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a365d')),
          ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
          ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
          ('FONTSIZE', (0, 0), (-1, -1), 7),
          ('ALIGN', (0, 0), (-1, 0), 'CENTER'), ('ALIGN', (6, 1), (8, -1), 'RIGHT'),
          ('GRID', (0, 0), (-1, -2), 0.3, colors.HexColor('#cbd5e0')),
          ('TOPPADDING', (0, 0), (-1, -1), 2), ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
          ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#edf2f7')),
          ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
          ('LINEABOVE', (0, -1), (-1, -1), 1, colors.HexColor('#1a365d'))]
    for i in range(1, len(led)-1):
        if i % 2 == 0:
            ls.append(('BACKGROUND', (0, i), (-1, i), colors.HexColor('#f7fafc')))
    lt.setStyle(TableStyle(ls))
    elements.append(lt)
    elements.append(Spacer(1, 10))

    elements.append(Paragraph("<b>PAYMENTS & RECEIPTS LOG</b>", ParagraphStyle(
        'sec', fontSize=10, textColor=colors.HexColor('#1a365d'), spaceBefore=3, spaceAfter=3)))
    pay = [['Sr', 'Date', 'Receipt/Ref', 'Mode', 'Bank/Account', 'Amount (PKR)', 'Remarks']]
    for i, r in enumerate(payments, 1):
        pay.append([str(i), str(r['payment_date'] or '')[:10], str(r['receipt_ref'] or '')[:13],
                    str(r['payment_mode'] or ''), str(r['bank_details'] or '')[:20],
                    f"{r['amount']:,.0f}", str(r['remarks'] or '')[:16]])
    if not payments:
        pay.append(['--', 'No payments', '', '', '', '', ''])
    pay.append(['', '', '', '', 'TOTAL', f"{summary['payments']:,.0f}", ''])
    pt = Table(pay, colWidths=[8*mm, 20*mm, 26*mm, 22*mm, 38*mm, 24*mm, 28*mm])
    ps = [('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a365d')),
          ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
          ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
          ('FONTSIZE', (0, 0), (-1, -1), 7),
          ('ALIGN', (0, 0), (-1, 0), 'CENTER'), ('ALIGN', (5, 1), (5, -1), 'RIGHT'),
          ('GRID', (0, 0), (-1, -2), 0.3, colors.HexColor('#cbd5e0')),
          ('TOPPADDING', (0, 0), (-1, -1), 2), ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
          ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#f0fff4')),
          ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
          ('LINEABOVE', (0, -1), (-1, -1), 1, colors.HexColor('#276749'))]
    for i in range(1, len(pay)-1):
        if i % 2 == 0:
            ps.append(('BACKGROUND', (0, i), (-1, i), colors.HexColor('#f7fafc')))
    pt.setStyle(TableStyle(ps))
    elements.append(pt)

    elements.append(Spacer(1, 12))
    elements.append(HRFlowable(width="100%", thickness=0.8, color=colors.HexColor('#cbd5e0')))
    elements.append(Paragraph(
        f"Generated: {datetime.now().strftime('%d-%b-%Y %H:%M')}  |  Statement: {agent['statement_no']}  |  Computer generated",
        ParagraphStyle('f', fontSize=7, alignment=TA_CENTER, textColor=colors.HexColor('#718096'), spaceBefore=3)))
    doc.build(elements)

# ==================== STREAMLIT UI ====================
st.set_page_config(
    page_title="Travel Agency - Multi-Agent Statement",
    page_icon="✈️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
<style>
    .main-header { font-size: 1.6rem; font-weight: 700; color: #1a365d; margin-bottom: 0; }
    .sub-header { color: #718096; font-size: 0.9rem; margin-top: 0; }
    .metric-card { background: white; border: 1px solid #e2e8f0; border-radius: 8px; padding: 12px; text-align: center; }
    .stMetric label { font-size: 0.85rem !important; }
</style>
""", unsafe_allow_html=True)

init_db()

# Session state
if "current_agent_id" not in st.session_state:
    val = get_setting("current_agent_id")
    st.session_state.current_agent_id = int(val) if val else None

agents = get_all_agents()
if not agents:
    st.error("No agents found. Please add an agent.")
    st.stop()

# Ensure current agent is valid
agent_ids = [a["id"] for a in agents]
if st.session_state.current_agent_id not in agent_ids:
    st.session_state.current_agent_id = agents[0]["id"]
    set_setting("current_agent_id", str(st.session_state.current_agent_id))

# ==================== SIDEBAR ====================
with st.sidebar:
    st.markdown("### ✈️ Travel Agency")
    st.markdown("**Multi-Agent Statement**")
    st.divider()

    agent_options = {f"{a['agent_name']} ({a['agency_city'] or 'N/A'})": a["id"] for a in agents}
    current_display = None
    for disp, aid in agent_options.items():
        if aid == st.session_state.current_agent_id:
            current_display = disp
            break
    if current_display is None:
        current_display = list(agent_options.keys())[0]

    selected = st.selectbox("Current Agent", list(agent_options.keys()),
                            index=list(agent_options.keys()).index(current_display))
    new_id = agent_options[selected]
    if new_id != st.session_state.current_agent_id:
        st.session_state.current_agent_id = new_id
        set_setting("current_agent_id", str(new_id))
        st.rerun()

    st.caption(f"Total Agents: {len(agents)}")
    st.divider()

    page = st.radio("Navigation", [
        "📊 Dashboard",
        "🎫 Tickets & Packages",
        "💰 Payments",
        "👥 All Agents",
        "⚙️ Agent Settings"
    ], label_visibility="collapsed")

# Current agent
agent = get_agent(st.session_state.current_agent_id)
summary = get_summary(st.session_state.current_agent_id)

# ==================== DASHBOARD ====================
if page == "📊 Dashboard":
    st.markdown(f'<p class="main-header">Dashboard — {agent["agent_name"]}</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="sub-header">{agent["agency_city"]} | {agent["statement_no"]} | {agent["statement_date"]}</p>', unsafe_allow_html=True)
    st.divider()

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Tickets Amount", f"Rs. {summary['tickets_amount']:,.0f}")
    c2.metric("Packages Amount", f"Rs. {summary['packages_amount']:,.0f}")
    c3.metric("Gross Booking", f"Rs. {summary['gross']:,.0f}")
    c4.metric("Payments Received", f"Rs. {summary['payments']:,.0f}")
    c5.metric("NET OUTSTANDING", f"Rs. {summary['outstanding']:,.0f}")

    st.subheader("Summary by Service Type")
    import pandas as pd
    svc_df = pd.DataFrame([
        {"Service Category": "Airline Tickets", "Count": summary["tickets_count"],
         "Total Amount (PKR)": summary["tickets_amount"], "Agent Balance": summary["tickets_amount"]},
        {"Service Category": "Umrah / Holiday Packages", "Count": summary["packages_count"],
         "Total Amount (PKR)": summary["packages_amount"], "Agent Balance": summary["packages_amount"]},
        {"Service Category": "Hotel & Visa Services", "Count": summary["visa_count"],
         "Total Amount (PKR)": summary["visa_amount"], "Agent Balance": summary["visa_amount"]},
        {"Service Category": "TOTAL ALL SERVICES",
         "Count": summary["tickets_count"]+summary["packages_count"]+summary["visa_count"],
         "Total Amount (PKR)": summary["gross"], "Agent Balance": summary["gross"]},
    ])
    st.dataframe(svc_df, use_container_width=True, hide_index=True)

    st.subheader("Export / Print")
    col1, col2, col3 = st.columns(3)

    with col1:
        if REPORTLAB_OK:
            if st.button("📄 Download PDF Statement", use_container_width=True):
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                    generate_pdf(st.session_state.current_agent_id, tmp.name)
                    with open(tmp.name, "rb") as f:
                        st.download_button("⬇️ Click to Download PDF", f.read(),
                                           file_name=f"Statement_{agent['statement_no']}.pdf",
                                           mime="application/pdf", use_container_width=True)
                    os.unlink(tmp.name)
        else:
            st.warning("Install reportlab for PDF")

    with col2:
        if OPENPYXL_OK:
            if st.button("📤 Export Excel", use_container_width=True):
                st.info("Excel export — use desktop version for full Excel, or install openpyxl and extend.")
        else:
            st.caption("Excel: pip install openpyxl")

    with col3:
        if st.button("🗑️ Clear This Agent Data", use_container_width=True):
            st.session_state.confirm_clear = True
        if st.session_state.get("confirm_clear"):
            if st.button("⚠️ Yes, Clear All Data"):
                clear_agent_data(st.session_state.current_agent_id)
                st.session_state.confirm_clear = False
                st.success("Data cleared!")
                st.rerun()

# ==================== LEDGER ====================
elif page == "🎫 Tickets & Packages":
    st.markdown(f'<p class="main-header">Tickets & Packages — {agent["agent_name"]}</p>', unsafe_allow_html=True)
    st.divider()

    with st.expander("➕ Add New Booking", expanded=False):
        with st.form("add_booking"):
            c1, c2 = st.columns(2)
            with c1:
                b_date = st.date_input("Date", value=date.today())
                b_pnr = st.text_input("PNR / Ref # *")
                b_cat = st.selectbox("Category", ["Ticket", "Package", "Visa/Hotel"])
                b_pass = st.text_input("Passenger Name / Group")
            with c2:
                b_sector = st.text_input("Sector / Description")
                b_qty = st.number_input("Qty / Seats", min_value=0.0, value=1.0, step=1.0)
                b_price = st.number_input("Unit Price (PKR)", min_value=0.0, value=0.0, step=100.0)
                b_status = st.selectbox("Payment Status", ["Pending", "Partial", "Paid", "Cancelled"])
            b_remarks = st.text_input("Remarks")
            if st.form_submit_button("💾 Save Booking", use_container_width=True):
                if not b_pnr.strip():
                    st.error("PNR / Ref # required")
                else:
                    add_ledger(st.session_state.current_agent_id, {
                        "date": str(b_date), "pnr_ref": b_pnr.strip(), "category": b_cat,
                        "passenger_name": b_pass, "sector_desc": b_sector,
                        "qty": b_qty, "unit_price": b_price, "payment_status": b_status, "remarks": b_remarks
                    })
                    st.success("Booking added!")
                    st.rerun()

    ledger = get_ledger(st.session_state.current_agent_id)
    if ledger:
        import pandas as pd
        df = pd.DataFrame([dict(r) for r in ledger])
        df_show = df[["id", "date", "pnr_ref", "category", "passenger_name", "sector_desc",
                      "qty", "unit_price", "total_amount", "payment_status", "remarks"]].copy()
        df_show.columns = ["ID", "Date", "PNR/Ref", "Category", "Passenger", "Sector",
                           "Qty", "Unit Price", "Total", "Status", "Remarks"]
        st.dataframe(df_show, use_container_width=True, hide_index=True)
        st.markdown(f"**TOTAL BOOKINGS: Rs. {summary['gross']:,.0f}**")

        st.subheader("Delete Booking")
        del_id = st.number_input("Enter ID to delete", min_value=0, step=1, key="del_led")
        if st.button("🗑️ Delete"):
            if del_id > 0:
                delete_ledger(int(del_id))
                st.success("Deleted!")
                st.rerun()
    else:
        st.info("No bookings yet. Add one above.")

# ==================== PAYMENTS ====================
elif page == "💰 Payments":
    st.markdown(f'<p class="main-header">Payments Received — {agent["agent_name"]}</p>', unsafe_allow_html=True)
    st.divider()

    with st.expander("➕ Add New Payment", expanded=False):
        with st.form("add_payment"):
            c1, c2 = st.columns(2)
            with c1:
                p_date = st.date_input("Payment Date", value=date.today())
                p_ref = st.text_input("Receipt / Ref #")
                p_mode = st.selectbox("Payment Mode", ["Cash", "Bank Transfer", "JazzCash", "EasyPaisa", "Cheque", "Other"])
            with c2:
                p_bank = st.text_input("Bank / Account Details")
                p_amt = st.number_input("Amount Received (PKR)", min_value=0.0, value=0.0, step=100.0)
                p_remarks = st.text_input("Remarks / Notes")
            if st.form_submit_button("💾 Save Payment", use_container_width=True):
                if p_amt <= 0:
                    st.error("Amount must be greater than 0")
                else:
                    add_payment(st.session_state.current_agent_id, {
                        "payment_date": str(p_date), "receipt_ref": p_ref, "payment_mode": p_mode,
                        "bank_details": p_bank, "amount": p_amt, "remarks": p_remarks
                    })
                    st.success("Payment added!")
                    st.rerun()

    payments = get_payments(st.session_state.current_agent_id)
    if payments:
        import pandas as pd
        df = pd.DataFrame([dict(r) for r in payments])
        df_show = df[["id", "payment_date", "receipt_ref", "payment_mode", "bank_details", "amount", "remarks"]].copy()
        df_show.columns = ["ID", "Date", "Receipt/Ref", "Mode", "Bank/Account", "Amount", "Remarks"]
        st.dataframe(df_show, use_container_width=True, hide_index=True)
        st.markdown(f"**TOTAL PAYMENTS: Rs. {summary['payments']:,.0f}**")

        st.subheader("Delete Payment")
        del_id = st.number_input("Enter ID to delete", min_value=0, step=1, key="del_pay")
        if st.button("🗑️ Delete Payment"):
            if del_id > 0:
                delete_payment(int(del_id))
                st.success("Deleted!")
                st.rerun()
    else:
        st.info("No payments yet. Add one above.")

# ==================== ALL AGENTS ====================
elif page == "👥 All Agents":
    st.markdown('<p class="main-header">All Agents Overview</p>', unsafe_allow_html=True)
    st.divider()

    import pandas as pd
    rows = []
    for a in agents:
        s = get_summary(a["id"])
        rows.append({
            "ID": a["id"], "Agent Name": a["agent_name"], "Agency": a["agency_city"],
            "Phone": a["contact_phone"], "Statement No": a["statement_no"],
            "Gross (PKR)": s["gross"], "Payments (PKR)": s["payments"],
            "Outstanding (PKR)": s["outstanding"]
        })
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    st.subheader("Switch Agent")
    switch_opts = {f"{a['agent_name']}": a["id"] for a in agents}
    chosen = st.selectbox("Select agent to switch", list(switch_opts.keys()))
    if st.button("Switch to this Agent"):
        st.session_state.current_agent_id = switch_opts[chosen]
        set_setting("current_agent_id", str(switch_opts[chosen]))
        st.success(f"Switched to {chosen}")
        st.rerun()

# ==================== AGENT SETTINGS ====================
elif page == "⚙️ Agent Settings":
    st.markdown('<p class="main-header">Agent Settings</p>', unsafe_allow_html=True)
    st.divider()

    tab1, tab2 = st.tabs(["Edit Current Agent", "Add New Agent"])

    with tab1:
        with st.form("edit_agent"):
            e_name = st.text_input("Agent Name *", value=agent["agent_name"])
            e_agency = st.text_input("Agency / City", value=agent["agency_city"] or "")
            e_phone = st.text_input("Contact / Phone", value=agent["contact_phone"] or "")
            e_date = st.text_input("Statement Date", value=agent["statement_date"] or "")
            e_no = st.text_input("Statement No", value=agent["statement_no"] or "")
            if st.form_submit_button("💾 Save Changes"):
                if not e_name.strip():
                    st.error("Name required")
                else:
                    update_agent(st.session_state.current_agent_id, e_name.strip(), e_agency, e_phone, e_date, e_no)
                    st.success("Updated!")
                    st.rerun()

        st.divider()
        if len(agents) > 1:
            if st.button("🗑️ Delete Current Agent", type="secondary"):
                delete_agent(st.session_state.current_agent_id)
                st.session_state.current_agent_id = None
                st.success("Agent deleted")
                st.rerun()
        else:
            st.caption("Cannot delete — at least one agent must remain.")

    with tab2:
        with st.form("add_agent"):
            n_name = st.text_input("Agent Name *", key="new_name")
            n_agency = st.text_input("Agency / City", key="new_agency")
            n_phone = st.text_input("Contact / Phone", key="new_phone")
            n_date = st.text_input("Statement Date", value=datetime.now().strftime("%d-%b-%Y"), key="new_date")
            n_no = st.text_input("Statement No", value=f"STM-{datetime.now().strftime('%Y')}-{len(agents)+1:03d}", key="new_no")
            if st.form_submit_button("➕ Add Agent"):
                if not n_name.strip():
                    st.error("Name required")
                else:
                    new_id = add_agent(n_name.strip(), n_agency, n_phone, n_date, n_no)
                    st.session_state.current_agent_id = new_id
                    set_setting("current_agent_id", str(new_id))
                    st.success(f"Agent '{n_name}' added!")
                    st.rerun()

st.markdown("---")
st.caption("Travel Agency Multi-Agent Statement Software | Web Version (Streamlit)")
