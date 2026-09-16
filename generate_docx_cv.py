import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def create_tech_lead_cv(output_path):
    doc = docx.Document()

    # Standard A4 page size (210mm x 297mm = 8.27 in x 11.69 in)
    for section in doc.sections:
        section.page_width = Inches(8.27)
        section.page_height = Inches(11.69)
        section.top_margin = Inches(0.42)
        section.bottom_margin = Inches(0.42)
        section.left_margin = Inches(0.52)
        section.right_margin = Inches(0.52)

    # Color Palette: Clean, professional, grounded
    NAVY = RGBColor(15, 23, 42)       # #0f172a
    BLUE = RGBColor(37, 99, 235)      # #2563eb
    SLATE = RGBColor(71, 85, 105)     # #475569
    DARK = RGBColor(30, 41, 59)       # #1e293b
    MUTED = RGBColor(100, 116, 139)   # #64748b

    def set_cell_background(cell, fill_hex):
        tcPr = cell._tc.get_or_add_tcPr()
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
        tcPr.append(shd)

    def clear_cell_borders(cell):
        tcPr = cell._tc.get_or_add_tcPr()
        tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/></w:tcBorders>')
        tcPr.append(tcBorders)

    def set_cell_margins(cell, top=25, bottom=25, left=40, right=40):
        tcPr = cell._tc.get_or_add_tcPr()
        tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
        tcPr.append(tcMar)

    # 1. HEADER SECTION
    header_table = doc.add_table(rows=1, cols=2)
    header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    header_table.autofit = False

    cell_left = header_table.cell(0, 0)
    cell_right = header_table.cell(0, 1)
    cell_left.width = Inches(4.65)
    cell_right.width = Inches(2.58)
    clear_cell_borders(cell_left)
    clear_cell_borders(cell_right)
    set_cell_margins(cell_left, top=10, bottom=20, left=0, right=20)
    set_cell_margins(cell_right, top=10, bottom=20, left=20, right=0)

    # Left: Name & Title
    p_name = cell_left.paragraphs[0]
    p_name.paragraph_format.space_after = Pt(1)
    p_name.paragraph_format.space_before = Pt(0)
    r_name = p_name.add_run("Md Niaj Shahriar Shishir")
    r_name.font.name = "Arial"
    r_name.font.size = Pt(18)
    r_name.font.bold = True
    r_name.font.color.rgb = NAVY

    p_title = cell_left.add_paragraph()
    p_title.paragraph_format.space_after = Pt(2)
    p_title.paragraph_format.space_before = Pt(0)
    r_title = p_title.add_run("Technical Lead | Odoo ERP & System Integrations")
    r_title.font.name = "Arial"
    r_title.font.size = Pt(9.5)
    r_title.font.bold = True
    r_title.font.color.rgb = BLUE

    p_cred = cell_left.add_paragraph()
    p_cred.paragraph_format.space_after = Pt(0)
    p_cred.paragraph_format.space_before = Pt(0)
    r_cred = p_cred.add_run("Certified Odoo Consultant | Head of Technology | 7+ Years Experience")
    r_cred.font.name = "Calibri"
    r_cred.font.size = Pt(8.5)
    r_cred.font.color.rgb = SLATE

    # Right: Contact Info
    contact_lines = [
        ("Portfolio: ", "nsrshishir.github.io/nsrshishir", True),
        ("Email: ", "niaj@metamorphosis.com.bd", False),
        ("Alt Email: ", "nsrshishir550@gmail.com", False),
        ("Phone: ", "+880 1827-512928 | Dhaka, BD", False),
        ("LinkedIn: ", "linkedin.com/in/nsrshishir", False),
        ("GitHub: ", "github.com/nsrshishir", False),
    ]

    p_contact = cell_right.paragraphs[0]
    for idx, (label, val, highlight) in enumerate(contact_lines):
        if idx > 0:
            p_contact = cell_right.add_paragraph()
        p_contact.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_contact.paragraph_format.space_after = Pt(0.5)
        p_contact.paragraph_format.space_before = Pt(0)

        rl = p_contact.add_run(label)
        rl.font.name = "Calibri"
        rl.font.size = Pt(7.8)
        rl.font.bold = True
        rl.font.color.rgb = MUTED

        rv = p_contact.add_run(val)
        rv.font.name = "Calibri"
        rv.font.size = Pt(7.8)
        if highlight:
            rv.font.bold = True
            rv.font.color.rgb = BLUE
        else:
            rv.font.color.rgb = DARK

    # Top Divider Line
    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_before = Pt(3)
    p_div.paragraph_format.space_after = Pt(4)
    p_div_border = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="10" w:space="1" w:color="2563EB"/></w:pBdr>')
    p_div._p.get_or_add_pPr().append(p_div_border)

    # 2. KEY METRICS CALLOUT STRIP
    stats_table = doc.add_table(rows=1, cols=4)
    stats_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    stats_table.autofit = False
    
    stats_data = [
        ("900+ Modules", "HAND-CODED IN PYTHON & XML"),
        ("SAP & 7 Gateways", "BUILT IN 1-MONTH SPRINTS"),
        ("14+ Upgrades", "ODOO 13 THROUGH 19"),
        ("Team Delivery", "3-MONTH SCOPE IN 2-3 WEEKS")
    ]
    
    col_w = Inches(1.80)
    for i, (val, label) in enumerate(stats_data):
        c = stats_table.cell(0, i)
        c.width = col_w
        set_cell_background(c, "F1F5F9")
        clear_cell_borders(c)
        set_cell_margins(c, top=40, bottom=40, left=30, right=30)
        
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0.5)
        rv = p.add_run(val)
        rv.font.name = "Arial"
        rv.font.size = Pt(8.8)
        rv.font.bold = True
        rv.font.color.rgb = NAVY

        p2 = c.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(0)
        rl = p2.add_run(label)
        rl.font.name = "Calibri"
        rl.font.size = Pt(6.4)
        rl.font.bold = True
        rl.font.color.rgb = BLUE

    # Section Heading Helper
    def add_section_heading(title_text, space_before=5):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(2.5)
        p.paragraph_format.keep_with_next = True
        
        bdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="6" w:space="1" w:color="CBD5E1"/></w:pBdr>')
        p._p.get_or_add_pPr().append(bdr)

        r = p.add_run(title_text.upper())
        r.font.name = "Arial"
        r.font.size = Pt(9.2)
        r.font.bold = True
        r.font.color.rgb = NAVY

    # Bullet helper
    def add_bullet(bold_prefix="", body_text="", space_after=1.2):
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.10
        
        if bold_prefix:
            rb = p.add_run(bold_prefix + " ")
            rb.font.name = "Calibri"
            rb.font.size = Pt(8.5)
            rb.font.bold = True
            rb.font.color.rgb = DARK
        
        if body_text:
            rt = p.add_run(body_text)
            rt.font.name = "Calibri"
            rt.font.size = Pt(8.5)
            rt.font.color.rgb = DARK
        return p

    # 3. EXECUTIVE SUMMARY
    add_section_heading("Executive Summary", space_before=4.5)
    p_sum = doc.add_paragraph()
    p_sum.paragraph_format.space_before = Pt(1)
    p_sum.paragraph_format.space_after = Pt(3)
    p_sum.paragraph_format.line_spacing = 1.12
    p_sum.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    r_sum = p_sum.add_run(
        "Technical Lead and Certified Odoo Consultant with 7+ years of hands-on experience building custom Odoo ERP "
        "solutions and integrating Odoo with external and proprietary systems. Hand-coded more than 900 custom modules in "
        "Python, JavaScript (OWL 2), and XML for 85+ enterprise clients. Experienced in delivering complex integrations under tight "
        "deadlines, including a complete bidirectional sync with SAP (customers, sales orders, deliveries, payments) in 1 month, and "
        "7 New Zealand and Australian payment gateways in a single month. Delivered integrations for mobile Sales Force Automation "
        "(SFA), telecom SMS marketing gateways, and cross-border retail POS systems. Completed 14+ database upgrades across Odoo 13 "
        "to 19 with zero data loss. Leads development teams using modern CLI agents and local AI models to cut delivery cycles from "
        "3 months down to 2–3 weeks while enforcing clean code standards via the Agentic E3 Framework."
    )
    r_sum.font.name = "Calibri"
    r_sum.font.size = Pt(8.5)
    r_sum.font.color.rgb = DARK

    # 4. CORE TECHNICAL & INTEGRATION COMPETENCIES
    add_section_heading("Core Technical & Integration Competencies", space_before=4)

    comp_table = doc.add_table(rows=3, cols=2)
    comp_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    comp_table.autofit = False

    comp_items = [
        ("Odoo Customization:", "900+ custom modules; custom models, business logic, OWL 2 reactive components, QWeb XML views, inheritances, compute field caching, automated server actions, cron tasks."),
        ("System & API Integrations:", "Bidirectional sync with SAP, external mobile SFA, WMS, and legacy SQL databases. REST, JSON-RPC, XML-RPC APIs, Webhook event pipelines, ETL data mapping."),
        ("Data Reliability & Sync:", "Idempotent API calls, queue-based retries, rate limits, audit logs, conflict resolution, automated stock reconciliation, and zero data divergence."),
        ("Performance Tuning:", "ORM batching, prefetching, eliminating N+1 queries, PostgreSQL index design, _read_group aggregations, Gunicorn worker isolation for API vs back-office traffic."),
        ("Database Upgrades:", "14+ major upgrades (v13 to v19), structural data retention scripts, legacy business logic extraction, and financial ledger reconciliation without data loss."),
        ("Technical Leadership:", "Team leadership, developer mentorship, sprint planning, client stakeholder management, 6-milestone delivery model, CI/CD, Docker deployment automation.")
    ]

    for idx, (head, desc) in enumerate(comp_items):
        row = idx // 2
        col = idx % 2
        cell = comp_table.cell(row, col)
        cell.width = Inches(3.61)
        clear_cell_borders(cell)
        set_cell_margins(cell, top=14, bottom=18, left=20, right=30)
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0.5)
        p.paragraph_format.line_spacing = 1.09
        
        rh = p.add_run("• " + head + " ")
        rh.font.name = "Calibri"
        rh.font.size = Pt(8.3)
        rh.font.bold = True
        rh.font.color.rgb = BLUE

        rd = p.add_run(desc)
        rd.font.name = "Calibri"
        rd.font.size = Pt(8.3)
        rd.font.color.rgb = DARK

    # 5. KEY INTEGRATION & CUSTOMIZATION DELIVERABLES
    add_section_heading("Key Integration & Customization Deliverables", space_before=4)

    add_bullet("SAP ERP Bidirectional Sync (Built in 1 Month):", 
               "Built an end-to-end integration between Odoo ERP and SAP. Synchronized customer accounts, sales orders, delivery fulfillment statuses, and payment records with automated retry queues and reconciliation within a 1-month deadline.")

    add_bullet("7 NZ & AUS Payment Gateways (1-Month Milestone):", 
               "Integrated and deployed 7 New Zealand and Australian payment gateways and BNPL platforms within a single 1-month milestone. Automated invoice reconciliation, payment webhooks, and PCI-DSS compliance.")

    add_bullet("Mobile Sales Force Automation (SFA) Integration:", 
               "Built a real-time sync connector between an external mobile SFA system and Odoo Enterprise for an FMCG manufacturer. Handled bulk field sales order import, stock reservation, and automated delivery order creation.")

    add_bullet("New Zealand Marketplaces (Kogan & Tradevine):", 
               "Built real-time API integrations connecting Odoo with Kogan and Tradevine. Automated product catalog mapping, multi-warehouse stock sync, and order fulfillment webhooks.")

    add_bullet("High-Throughput Telecom SMS Gateways:", 
               "Built custom SMS gateway modules connecting Odoo directly to telecommunication carrier APIs for automated order updates, OTP verification, and bulk customer notification campaigns.")

    add_bullet("Cross-Border Omnichannel POS & Loyalty System:", 
               "Deployed retail POS systems across nationwide store chains and an international branch in India. Built custom multi-tier prepaid membership cards, customer balance wallets, and offline-to-online sync.")

    add_bullet("Odoo Database Model Context Protocol (MCP) Server:", 
               "Built and deployed an MCP server exposing Odoo ORM schemas and models to AI coding tools (Claude Code, Antigravity, Cline, Puku) for automated schema inspection and task execution.")

    add_bullet("ZATCA Phase 2 & Tax Invoicing Compliance:", 
               "Built tax compliance modules connecting Odoo accounting with regulatory portals: Saudi Arabia ZATCA Phase 2 (XML signing, ECDSA cryptographic hashing, QR codes) and Bangladesh NBR tax engines.")

    # Explicit page break to guarantee pristine 2-page split
    doc.add_page_break()

    # 6. PROFESSIONAL LEADERSHIP & CAREER HISTORY (PAGE 2)
    add_section_heading("Professional Leadership & Engineering Experience", space_before=0)

    # Role 1
    p_r1 = doc.add_paragraph()
    p_r1.paragraph_format.space_before = Pt(2.5)
    p_r1.paragraph_format.space_after = Pt(1)
    p_r1.paragraph_format.keep_with_next = True
    
    r_role = p_r1.add_run("Head of Technology  |  Metamorphosis Ltd.")
    r_role.font.name = "Arial"
    r_role.font.size = Pt(9.8)
    r_role.font.bold = True
    r_role.font.color.rgb = NAVY
    
    r_date = p_r1.add_run("  [Feb 2026 – Present]")
    r_date.font.name = "Calibri"
    r_date.font.size = Pt(8.8)
    r_date.font.bold = True
    r_date.font.color.rgb = BLUE

    add_bullet("Technical Leadership for ERP Customizations:", 
               "Leads engineering teams delivering Odoo customization projects for 85+ global and domestic enterprise clients (Bangladesh, New Zealand, Saudi Arabia, Spain, South Sudan). Guides core architecture, custom themes, and proprietary integrations.")

    add_bullet("Fast-Track Enterprise Integrations (SAP, SFA, 7 Payment Gateways):", 
               "Led the delivery of major integrations on tight timelines: full SAP sync (customers, orders, deliveries, payments) in 1 month, SFA field order import for FMCG clients, and 7 NZ & AUS payment gateways in a 1-month milestone.")

    add_bullet("Fast-Paced Team Delivery (3-Month Scope in 2–3 Weeks):", 
               "Established a delivery workflow combining modular Odoo patterns, automated test suites, and local AI tools (Claude Code, Antigravity, local LLMs), cutting project cycle times by up to 75% with high first-pass UAT acceptance.")

    add_bullet("Odoo Enterprise Database Upgrades (v13.0 to v19.0):", 
               "Led 14+ major database upgrades across multiple version jumps (v13-v17, v14-v19, v16-v19). Used custom data retention scripts to migrate years of custom business data while retiring outdated code.")

    add_bullet("Nginx Server Worker Isolation:", 
               "Configured Nginx upstream routing into separate worker pools, isolating heavy external API sync and e-commerce traffic from internal back-office users to eliminate slowdowns.")

    # Role 2
    p_r2 = doc.add_paragraph()
    p_r2.paragraph_format.space_before = Pt(4.5)
    p_r2.paragraph_format.space_after = Pt(1)
    p_r2.paragraph_format.keep_with_next = True
    
    r_role2 = p_r2.add_run("Sr. Odoo Systems Engineer & DevOps Lead  |  Metamorphosis Ltd.")
    r_role2.font.name = "Arial"
    r_role2.font.size = Pt(9.3)
    r_role2.font.bold = True
    r_role2.font.color.rgb = NAVY
    
    r_date2 = p_r2.add_run("  [Oct 2023 – Jan 2026]")
    r_date2.font.name = "Calibri"
    r_date2.font.size = Pt(8.8)
    r_date2.font.bold = True
    r_date2.font.color.rgb = BLUE

    add_bullet("Marketplace & Payment Integrations:", 
               "Built real-time sync modules connecting Odoo to international e-commerce channels (Kogan, Tradevine, BNPL platforms). Built automated reconciliation scripts to prevent stock mismatches.")

    add_bullet("Custom Module Development & SFA Sync:", 
               "Wrote custom Odoo modules for business automations, discrete manufacturing assembly lines, multi-currency accounting, and mobile SFA order processing.")

    add_bullet("Deployment & Infrastructure Automation:", 
               "Built Docker and Bash deployment pipelines managing 300+ live Odoo instances, cutting onboarding deployment time by 85%.")

    # Role 3
    p_r3 = doc.add_paragraph()
    p_r3.paragraph_format.space_before = Pt(4.5)
    p_r3.paragraph_format.space_after = Pt(1)
    p_r3.paragraph_format.keep_with_next = True
    
    r_role3 = p_r3.add_run("Full Stack Odoo Developer  |  Metamorphosis Ltd.")
    r_role3.font.name = "Arial"
    r_role3.font.size = Pt(9.3)
    r_role3.font.bold = True
    r_role3.font.color.rgb = NAVY
    
    r_date3 = p_r3.add_run("  [2022 – Sep 2023]")
    r_date3.font.name = "Calibri"
    r_date3.font.size = Pt(8.8)
    r_date3.font.bold = True
    r_date3.font.color.rgb = BLUE

    add_bullet("Custom Modules Built to Business Requirements:", 
               "Developed hundreds of custom Odoo modules matching specific client business workflows, custom reporting, and customer portal features.")

    add_bullet("Payment Gateways & PostgreSQL Tuning:", 
               "Integrated local payment gateways (cards, mobile banking) and optimized complex PostgreSQL queries and indexing for high-volume transactions.")

    # Role 4 & 5
    p_r4 = doc.add_paragraph()
    p_r4.paragraph_format.space_before = Pt(4.5)
    p_r4.paragraph_format.space_after = Pt(1)
    p_r4.paragraph_format.keep_with_next = True
    
    r_role4 = p_r4.add_run("Software Engineer & Systems Specialist  |  MetaKave")
    r_role4.font.name = "Arial"
    r_role4.font.size = Pt(9.3)
    r_role4.font.bold = True
    r_role4.font.color.rgb = NAVY
    
    r_date4 = p_r4.add_run("  [2020 – 2021]")
    r_date4.font.name = "Calibri"
    r_date4.font.size = Pt(8.8)
    r_date4.font.bold = True
    r_date4.font.color.rgb = BLUE

    add_bullet("Magento 2+ Backend & Mobile APIs:", 
               "Built backend web services and API endpoints for Magento 2+ e-commerce platforms connected to mobile apps. Managed cloud servers and SSL/Nginx reverse proxies.")

    p_r5 = doc.add_paragraph()
    p_r5.paragraph_format.space_before = Pt(4.5)
    p_r5.paragraph_format.space_after = Pt(1)
    p_r5.paragraph_format.keep_with_next = True
    
    r_role5 = p_r5.add_run("IT Executive  |  Aptech Industries")
    r_role5.font.name = "Arial"
    r_role5.font.size = Pt(9.3)
    r_role5.font.bold = True
    r_role5.font.color.rgb = NAVY
    
    r_date5 = p_r5.add_run("  [Jun 2017 – Sep 2019]")
    r_date5.font.name = "Calibri"
    r_date5.font.size = Pt(8.8)
    r_date5.font.bold = True
    r_date5.font.color.rgb = BLUE

    add_bullet("Enterprise Infrastructure Administration:", 
               "Managed enterprise IT operations, Windows Server administration (Active Directory, DNS, IIS, Group Policy up to 2019), network security, and hardware/software deployments.")

    # 7. EDUCATION, CERTIFICATIONS & RESEARCH
    add_section_heading("Education, Certifications & Research", space_before=4)

    add_bullet("Certified Odoo Consultant:", "Metamorphosis Ltd. (Official Odoo Silver Partner). Extensive expertise in Odoo standard architecture, ORM, customization guidelines, and business process modeling.")
    add_bullet("Master of Information Technology (MIT):", "Jahangirnagar University (2018 – 2019). Focus on Enterprise Systems, Advanced Databases, and Distributed Architectures.")
    add_bullet("Bachelor of Science (B.Sc):", "Shahjalal University of Science and Technology (SUST). 1st Major in Mathematics, 2nd Major in Computer Science & Engineering (CSE).")
    add_bullet("IEEE Research Author (2021):", "Published paper on data processing optimization in Hadoop MapReduce, introducing the IMC algorithm (65% efficiency gain). Applied these distributed processing principles to high-volume Odoo database operations.")

    # Footer note
    p_foot = doc.add_paragraph()
    p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_foot.paragraph_format.space_before = Pt(5)
    p_foot.paragraph_format.space_after = Pt(0)
    rf = p_foot.add_run("Md Niaj Shahriar Shishir  |  Technical Lead — Odoo & Integrations  |  Portfolio: nsrshishir.github.io/nsrshishir")
    rf.font.name = "Calibri"
    rf.font.size = Pt(7.2)
    rf.font.color.rgb = MUTED

    doc.save(output_path)
    print(f"Successfully generated clean, no-fluff CV: {output_path}")

if __name__ == "__main__":
    output_docx = "/home/niaj/Documents/nsrshishir/Md_Niaj_Shahriar_Shishir_Tech_Lead_Odoo_Customization_Integrations.docx"
    create_tech_lead_cv(output_docx)
