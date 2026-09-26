import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="{top}" w:type="dxa"/>'
        f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:left w:w="{left}" w:type="dxa"/>'
        f'<w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)

def create_document():
    doc = docx.Document()

    # Page Margins: 1 inch everywhere
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.9)
        section.right_margin = Inches(0.9)

    # Palette
    COLOR_PRIMARY = RGBColor(18, 52, 102)     # Deep Navy
    COLOR_SECONDARY = RGBColor(184, 34, 42)   # Metfone Red
    COLOR_TEXT = RGBColor(40, 40, 40)         # Charcoal
    COLOR_MUTED = RGBColor(100, 100, 100)

    # Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("METFONE EXPRESS AUTOMATION BOT")
    r_title.bold = True
    r_title.font.size = Pt(22)
    r_title.font.color.rgb = COLOR_PRIMARY

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run("Executive Reference Manual: Operational Commands & Business Formulas")
    r_sub.font.size = Pt(13)
    r_sub.font.italic = True
    r_sub.font.color.rgb = COLOR_SECONDARY

    p_meta = doc.add_paragraph()
    p_meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_meta = p_meta.add_run("Prepared for Management Review | System Version 2.4 | Nationwide 36 Branches")
    r_meta.font.size = Pt(9.5)
    r_meta.font.color.rgb = COLOR_MUTED

    doc.add_paragraph() # Spacer

    # ─────────────────────────────────────────────────────────────────────────────
    # SECTION 1: EXECUTIVE OVERVIEW
    # ─────────────────────────────────────────────────────────────────────────────
    h1 = doc.add_heading(level=1)
    r_h1 = h1.add_run("1. Executive Overview & Business Value")
    r_h1.font.color.rgb = COLOR_PRIMARY

    p_lead = doc.add_paragraph()
    p_lead.add_run(
        "This Telegram Operations Bot automates end-to-end monitoring, operational dispatching, "
        "delivery speed auditing, stagnant inventory penalization, and executive reporting across all "
        "36 Metfone Express branches and 5 Zones in Cambodia. It replaces manual spreadsheet tracking "
        "with real-time, automated intelligence delivered straight to Telegram management groups."
    )

    # Bullet summary
    bullets = [
        ("Massive Scalability: ", "Processes ~68,000 active shipment records per cycle from Metfone TMS in seconds."),
        ("Automated Triple-Schedule: ", "Generates and pushes network-wide audit reports automatically at 08:00, 14:00, and 16:00."),
        ("Real-Time SLA Enforcement: ", "Tracks delivery speed and automatically calculates courier bonus commissions or delay fines."),
        ("Executive Integration: ", "Automates Microsoft Excel COM engine to generate populated Master Daily Reports and swipeable image albums to senior leadership.")
    ]
    for b_title, b_desc in bullets:
        bp = doc.add_paragraph(style='List Bullet')
        r1 = bp.add_run(b_title)
        r1.bold = True
        r1.font.color.rgb = COLOR_PRIMARY
        bp.add_run(b_desc)

    doc.add_paragraph()

    # ─────────────────────────────────────────────────────────────────────────────
    # SECTION 2: COMMAND INVENTORY (TOTAL: 50 COMMANDS / ALIASES)
    # ─────────────────────────────────────────────────────────────────────────────
    h2 = doc.add_heading(level=1)
    r_h2 = h2.add_run("2. Comprehensive Bot Command Inventory")
    r_h2.font.color.rgb = COLOR_PRIMARY

    p_cmd_intro = doc.add_paragraph()
    p_cmd_intro.add_run(
        "The bot supports 50 command aliases grouped into 6 core operational modules. "
        "Commands can be triggered either with a leading slash (e.g. /total) or as natural text commands."
    )

    table_data = [
        # (Module, Commands, Description, Target / Output)
        (
            "Operations & Push",
            "push, /push [handle/all]\n/total, /pending\n/dvc, /dvczone\n/tomorrow [all/zone/branch]\n/vs, /vs2",
            "Generates real-time Pickup, Delivery, and Pending counts. Pushes targeted HD operational images to branch and zone Telegram groups. Tracks DVC container handover and forecasts incoming MEGA hub arrivals.",
            "Group dispatch & Private DM with HD Images + XLSX"
        ),
        (
            "SLA & Speed Tracking",
            "/speed\n/speed all\n/speed [zone/branch]",
            "Audits door-to-door VTT delivery performance. Calculates time-to-deliver from physical arrival scan, classifies orders into 4 tiers (<2h, 2-4h, 4-8h, >8h), and computes individual courier commissions.",
            "Executive Speed Dashboard Image + Audit Excel"
        ),
        (
            "Stagnant Inventory & Penalty",
            "/penalty\n/penalty all\n/penalty [zone/branch]",
            "Audits packages stuck in Handover or Delivery. Evaluates parcel aging against SLA thresholds, applies $0.10 or $0.40 fines, excuses customer appointment delays, and attributes penalties to the exact physical post office.",
            "SLA Penalty Dashboard Image + Complete Audit Table"
        ),
        (
            "Executive Reporting",
            "/dailyreport [date]\n/report\n/deletereport",
            "Spawns Microsoft Excel engine to populate Master Daily Report across 8 sheets. Runs formula calculation, captures 5 high-res dashboard images (Day, Service Point, Agent, Showroom, New Customer), and forwards summary text to Report Alert group (-5587688944).",
            "Text Summary + 5-Photo Album + Populated Master Excel"
        ),
        (
            "Bill Investigation",
            "/find <bill_id>\n/check <bill_id>\n/trace <bill_id>\n/ask <question>\n/qr <bill_id>\n/export [date]",
            "Instant lifecycle timeline lookup querying live Metfone TMS API. Inspects exact courier actions, scans, timestamps, and customer remarks. Generates physical scannable barcodes/QR codes.",
            "Interactive Telegram cards & Barcode images"
        ),
        (
            "Exceptions & System Control",
            "/delay <bill_id> <days>\n/undelay <bill_id>\n/delaylist, /delayed, /backlog\n/pause, /resume\n/register, /groups\n/settoken <token>",
            "Exempts specific bills from SLA penalty tracking (customer appointment, pending investigation). Controls automated scheduled pushes. Manages Telegram group permissions and updates Metfone API authentication tokens.",
            "Instant administrative state persistence"
        )
    ]

    table = doc.add_table(rows=1, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    # Header
    hdr_cells = table.rows[0].cells
    headers = ["Functional Area", "Commands / Aliases", "Operational Purpose", "Output Deliverables"]
    widths = [Inches(1.5), Inches(1.8), Inches(2.2), Inches(1.5)]

    for idx, (cell, h_text) in enumerate(zip(hdr_cells, headers)):
        cell.width = widths[idx]
        p = cell.paragraphs[0]
        r = p.add_run(h_text)
        r.bold = True
        r.font.size = Pt(10)
        r.font.color.rgb = RGBColor(255, 255, 255)
        set_cell_background(cell, "123466") # Navy
        set_cell_margins(cell, top=120, bottom=120, left=140, right=140)

    for row_idx, (mod, cmds, desc, out) in enumerate(table_data):
        row_cells = table.add_row().cells
        bg_col = "F8FAFC" if row_idx % 2 == 0 else "FFFFFF"
        for idx, (cell, val) in enumerate(zip(row_cells, [mod, cmds, desc, out])):
            cell.width = widths[idx]
            p = cell.paragraphs[0]
            r = p.add_run(val)
            r.font.size = Pt(8.5)
            if idx == 1:
                r.font.name = "Consolas"
                r.bold = True
            set_cell_background(cell, bg_col)
            set_cell_margins(cell, top=100, bottom=100, left=120, right=120)

    doc.add_paragraph()

    # ─────────────────────────────────────────────────────────────────────────────
    # SECTION 3: BUSINESS FORMULAS & SLA MATHEMATICS
    # ─────────────────────────────────────────────────────────────────────────────
    h3 = doc.add_heading(level=1)
    r_h3 = h3.add_run("3. Business Formulas & Operational SLA Rules")
    r_h3.font.color.rgb = COLOR_PRIMARY

    # --- 3.1 SPEED & COURIER COMMISSION ---
    h3_1 = doc.add_heading(level=2)
    r_h3_1 = h3_1.add_run("3.1 Delivery Speed & Courier Commission (/speed all)")
    r_h3_1.font.color.rgb = COLOR_SECONDARY

    p_sp = doc.add_paragraph()
    p_sp.add_run("• Target Population: ").bold = True
    p_sp.add_run("All delivered orders (Status 410) with Door-to-Door Delivery service (VAS = VTT).\n")
    p_sp.add_run("• Elapsed Duration Formula: ").bold = True
    p_sp.add_run("Duration (Hours) = Time Delivered (410) - Time Arrived at Branch (306/309/Hub Scan)\n")
    p_sp.add_run("• Customer Appointment Protection: ").bold = True
    p_sp.add_run(
        "If a parcel was delayed due to customer appointment (Status 420, 472, SHIP_AGAIN, POSTPONE, or notes 'hẹn/reschedule'), "
        "the arrival baseline is automatically adjusted to the delivery day's morning scan (08:00 AM) so the courier is not unfairly penalized."
    )

    # Speed Tier Table
    s_table = doc.add_table(rows=1, cols=4)
    s_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    s_headers = ["Delivery Duration", "Speed Classification", "Courier Rate / Commission", "Performance Incentive"]
    s_widths = [Inches(1.8), Inches(1.8), Inches(1.8), Inches(1.6)]

    for idx, (cell, h) in enumerate(zip(s_table.rows[0].cells, s_headers)):
        cell.width = s_widths[idx]
        p = cell.paragraphs[0]
        r = p.add_run(h)
        r.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(255, 255, 255)
        set_cell_background(cell, "1E3A8A")
        set_cell_margins(cell, top=100, bottom=100, left=120, right=120)

    s_rows = [
        ("< 2 Hours", "Fast Delivery (Elite)", "$0.30 / order", "+50% Bonus Commission"),
        ("2 – 4 Hours", "Standard Delivery", "$0.25 / order", "+25% Bonus Commission"),
        ("4 – 8 Hours", "Normal Delivery", "$0.20 / order", "Standard SLA Baseline (100%)"),
        ("> 8 Hours", "Delayed Delivery", "$0.15 / order", "-25% SLA Delay Fine")
    ]
    for row_idx, r_vals in enumerate(s_rows):
        row_cells = s_table.add_row().cells
        bg_col = "ECFDF5" if row_idx == 0 else ("EFF6FF" if row_idx == 1 else ("FFFFFF" if row_idx == 2 else "FEF2F2"))
        for idx, (cell, val) in enumerate(zip(row_cells, r_vals)):
            cell.width = s_widths[idx]
            p = cell.paragraphs[0]
            r = p.add_run(val)
            r.font.size = Pt(9)
            if idx == 1 or idx == 2:
                r.bold = True
            set_cell_background(cell, bg_col)
            set_cell_margins(cell, top=90, bottom=90, left=110, right=110)

    p_sp_dash = doc.add_paragraph()
    p_sp_dash.add_run("\nKey Dashboard Summary Formulas:\n").bold = True
    p_sp_dash.add_run("  • % Within 8h = ((Orders < 2h + Orders 2-4h + Orders 4-8h) / Total Delivered) × 100%\n")
    p_sp_dash.add_run("  • % Over 8h = (Orders > 8h / Total Delivered) × 100%\n")
    p_sp_dash.add_run("  • Total Branch Commission = Σ (Delivered Orders × Speed Rate)")

    # --- 3.2 STAGNANT INVENTORY & PENALTY ---
    h3_2 = doc.add_heading(level=2)
    r_h3_2 = h3_2.add_run("3.2 Stagnant Inventory & SLA Penalty (/penalty all)")
    r_h3_2.font.color.rgb = COLOR_SECONDARY

    p_pen = doc.add_paragraph()
    p_pen.add_run("• Target Population: ").bold = True
    p_pen.add_run("All active packages in custody at branches/hubs divided into Handover stream and Delivery stream.\n")
    p_pen.add_run("• Attribution Rule: ").bold = True
    p_pen.add_run(
        "Penalties are attributed strictly to the latest scanned Post Office where the parcel is physically sitting. "
        "Agents and Showrooms are mapped back to their parent branch."
    )

    pen_table = doc.add_table(rows=1, cols=4)
    pen_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    pen_headers = ["Parcel Age in Custody", "Risk Classification", "Fine Amount ($)", "SLA Action"]
    pen_widths = [Inches(1.8), Inches(1.8), Inches(1.6), Inches(1.8)]

    for idx, (cell, h) in enumerate(zip(pen_table.rows[0].cells, pen_headers)):
        cell.width = pen_widths[idx]
        p = cell.paragraphs[0]
        r = p.add_run(h)
        r.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(255, 255, 255)
        set_cell_background(cell, "7F1D1D") # Deep Red
        set_cell_margins(cell, top=100, bottom=100, left=120, right=120)

    pen_rows = [
        ("≤ 1 Day", "Safe (Normal SLA)", "$0.00", "No fine applied"),
        ("> 1 Day (2 Days)", "Backlog / Stagnant", "$0.10 / bill", "Standard Delay Deduction"),
        ("≥ 3 Days", "Urgent / Critical Backlog", "$0.40 / bill", "Severe Stagnant Fine")
    ]
    for row_idx, r_vals in enumerate(pen_rows):
        row_cells = pen_table.add_row().cells
        bg_col = "F0FDF4" if row_idx == 0 else ("FEF3C7" if row_idx == 1 else "FEE2E2")
        for idx, (cell, val) in enumerate(zip(row_cells, r_vals)):
            cell.width = pen_widths[idx]
            p = cell.paragraphs[0]
            r = p.add_run(val)
            r.font.size = Pt(9)
            if idx in (1, 2):
                r.bold = True
            set_cell_background(cell, bg_col)
            set_cell_margins(cell, top=90, bottom=90, left=110, right=110)

    p_pen_dash = doc.add_paragraph()
    p_pen_dash.add_run("\nExclusions & Compliance Ratios:\n").bold = True
    p_pen_dash.add_run("  • Customer Problem Exemption: Bills with status 420, 471, 472, 480 (customer reschedule / address verification) are 100% EXCUSED ($0.00 fine).\n")
    p_pen_dash.add_run("  • % RIGHT Handover = ((Total Handover - Penalty Handover) / Total Handover) × 100%\n")
    p_pen_dash.add_run("  • % RIGHT Delivery = ((Total Delivery - Penalty Delivery) / Total Delivery) × 100%\n")
    p_pen_dash.add_run("  • Total Branch Fine = Σ (Penalty Bills × Fine Amount)")

    # --- 3.3 TOMORROW FORECAST ---
    h3_3 = doc.add_heading(level=2)
    r_h3_3 = h3_3.add_run("3.3 Next-Day Incoming Forecast (/tomorrow all)")
    r_h3_3.font.color.rgb = COLOR_SECONDARY

    p_tmr = doc.add_paragraph()
    p_tmr.add_run(
        "Forecasts incoming parcel volume headed to each branch for next-day dispatch using strict container and trip validation:\n"
    )
    p_tmr.add_run("  1. Status Filter: ").bold = True
    p_tmr.add_run("Order must be in Status 306 (Storage/Container Handover).\n")
    p_tmr.add_run("  2. Hub Transit Verification: ").bold = True
    p_tmr.add_run("Order must have been physically handed over at MEGA1 central hub.\n")
    p_tmr.add_run("  3. Courier Handover (DV Driver): ").bold = True
    p_tmr.add_run("The most recent custody scan must be with a long-haul transfer driver (DVCMEGA1 / DVC...).\n")
    p_tmr.add_run("  4. Exclusions: ").bold = True
    p_tmr.add_run("Orders already received at the destination post office (e.g. at KRAP001) or originating from local shops (Status 302/310) are strictly excluded.")

    # --- 3.4 DAILY REPORT ---
    h3_4 = doc.add_heading(level=2)
    r_h3_4 = h3_4.add_run("3.4 Master Daily Executive Report (/dailyreport)")
    r_h3_4.font.color.rgb = COLOR_SECONDARY

    p_dr = doc.add_paragraph()
    p_dr.add_run(
        "Automates senior management daily reporting via headless Microsoft Excel COM automation:\n"
    )
    p_dr.add_run("  • Executive Volume & Variance: ").bold = True
    p_dr.add_run("Total Daily Volume = Total Orders (Today at Cutoff) vs Total Same Cutoff Last Week (Variance: ± Δ Orders).\n")
    p_dr.add_run("  • Channel Breakdown: ").bold = True
    p_dr.add_run("Classifies volume into Service Points (xxxP...), Authorized Agents (xxxA...), and Showrooms (xxxS...).\n")
    p_dr.add_run("  • Zero-Order Outlets: ").bold = True
    p_dr.add_run("Detects and alerts branches and post offices with 0 generated bills in-day.\n")
    p_dr.add_run("  • Multi-Sheet High-Res Renders: ").bold = True
    p_dr.add_run("Captures and uploads 5 sheet images: Day Report, Service Point, Agent, Showroom, and New Customer Development.")

    doc.add_paragraph()

    # ─────────────────────────────────────────────────────────────────────────────
    # SECTION 4: SYSTEM ARCHITECTURE & PERFORMANCE
    # ─────────────────────────────────────────────────────────────────────────────
    h4 = doc.add_heading(level=1)
    r_h4 = h4.add_run("4. System Architecture, CPU & Performance")
    r_h4.font.color.rgb = COLOR_PRIMARY

    p_arch = doc.add_paragraph()
    p_arch.add_run(
        "The system is architected for asynchronous, multi-threaded high throughput. "
        "The table below outlines real-world CPU consumption and processing latency across live operations:"
    )

    perf_table = doc.add_table(rows=1, cols=4)
    perf_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    perf_headers = ["Operation", "Underlying Engine", "CPU Consumption", "Total Execution Time"]
    perf_widths = [Inches(1.8), Inches(1.8), Inches(1.8), Inches(1.6)]

    for idx, (cell, h) in enumerate(zip(perf_table.rows[0].cells, perf_headers)):
        cell.width = perf_widths[idx]
        p = cell.paragraphs[0]
        r = p.add_run(h)
        r.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(255, 255, 255)
        set_cell_background(cell, "0F172A") # Slate dark
        set_cell_margins(cell, top=100, bottom=100, left=120, right=120)

    perf_rows = [
        ("Push / /total", "pandas + openpyxl + Pillow", "15% – 35% of 1 core", "4 – 8 seconds"),
        ("/speed all", "calamine + pandas + Pillow", "20% – 40% of 1 core", "5 – 10 seconds"),
        ("/penalty all", "calamine + openpyxl", "20% – 40% of 1 core", "6 – 10 seconds"),
        ("/tomorrow all", "openpyxl + API prefetch", "15% – 30% of 1 core", "4 – 8 seconds"),
        ("/dailyreport", "MS Excel COM (EXCEL.EXE)", "80% – 100% of 1–2 cores", "18 – 32 seconds")
    ]
    for row_idx, r_vals in enumerate(perf_rows):
        row_cells = perf_table.add_row().cells
        bg_col = "F8FAFC" if row_idx % 2 == 0 else "FFFFFF"
        for idx, (cell, val) in enumerate(zip(row_cells, r_vals)):
            cell.width = perf_widths[idx]
            p = cell.paragraphs[0]
            r = p.add_run(val)
            r.font.size = Pt(9)
            if idx == 0 or idx == 3:
                r.bold = True
            set_cell_background(cell, bg_col)
            set_cell_margins(cell, top=90, bottom=90, left=110, right=110)

    # Footer note
    doc.add_paragraph()
    p_foot = doc.add_paragraph()
    p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_foot = p_foot.add_run("— End of Metfone Express Automation Bot Technical Documentation —")
    r_foot.font.size = Pt(9.5)
    r_foot.font.italic = True
    r_foot.font.color.rgb = COLOR_MUTED

    target_path = r"c:\Users\DELL\Desktop\daily_push\METFONE_EXPRESS_TELEGRAM_BOT_DOCUMENTATION.docx"
    doc.save(target_path)
    print(f"Document successfully created at: {target_path}")

if __name__ == "__main__":
    create_document()
