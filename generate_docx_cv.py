import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def create_tech_lead_cv(output_path):
    doc = docx.Document()

    # Set standard A4 page size (210mm x 297mm = 8.27 in x 11.69 in)
    for section in doc.sections:
        section.page_width = Inches(8.27)
        section.page_height = Inches(11.69)
        section.top_margin = Inches(0.48)
        section.bottom_margin = Inches(0.48)
        section.left_margin = Inches(0.55)
        section.right_margin = Inches(0.55)

    # Color Palette
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

    def set_cell_margins(cell, top=40, bottom=40, left=60, right=60):
        tcPr = cell._tc.get_or_add_tcPr()
        tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
        tcPr.append(tcMar)

    # 1. HEADER SECTION
    header_table = doc.add_table(rows=1, cols=2)
    header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    header_table.autofit = False

    cell_left = header_table.cell(0, 0)
    cell_right = header_table.cell(0, 1)
    cell_left.width = Inches(4.5)
    cell_right.width = Inches(2.67)
    clear_cell_borders(cell_left)
    clear_cell_borders(cell_right)
    set_cell_margins(cell_left, top=20, bottom=40, left=0, right=40)
    set_cell_margins(cell_right, top=20, bottom=40, left=40, right=0)

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
    r_title = p_title.add_run("TECH LEAD — ODOO CUSTOMIZATION & PROPRIETARY INTEGRATIONS")
    r_title.font.name = "Arial"
    r_title.font.size = Pt(8.8)
    r_title.font.bold = True
    r_title.font.color.rgb = BLUE

    p_cred = cell_left.add_paragraph()
    p_cred.paragraph_format.space_after = Pt(0)
    p_cred.paragraph_format.space_before = Pt(0)
    r_cred = p_cred.add_run("Certified Odoo Consultant | Odoo Solution Architect by Practice | Head of Technology")
    r_cred.font.name = "Calibri"
    r_cred.font.size = Pt(8.2)
    r_cred.font.color.rgb = SLATE

    # Right: Contact Info
    contact_lines = [
        ("Live Portfolio: ", "nsrshishir.github.io/nsrshishir", True),
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
        rl.font.size = Pt(8)
        rl.font.bold = True
        rl.font.color.rgb = MUTED

        rv = p_contact.add_run(val)
        rv.font.name = "Calibri"
        rv.font.size = Pt(8)
        if highlight:
            rv.font.bold = True
            rv.font.color.rgb = BLUE
        else:
            rv.font.color.rgb = DARK

    # Top Divider Line
    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_before = Pt(4)
    p_div.paragraph_format.space_after = Pt(5)
    p_div_border = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="10" w:space="1" w:color="2563EB"/></w:pBdr>')
    p_div._p.get_or_add_pPr().append(p_div_border)

    # 2. KEY METRICS CALLOUT STRIP
    stats_table = doc.add_table(rows=1, cols=4)
    stats_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    stats_table.autofit = False
    
    stats_data = [
        ("900+ Custom Modules", "HAND-CODED ODOO APPS"),
        ("14+ Enterprise Upgrades", "V13.0 THROUGH V19.0"),
        ("Proprietary & 3rd-Party", "CUSTOM API & ETL SYNC"),
        ("3 Mo. Scope in 2-3 Wks", "HYPER-VELOCITY DELIVERY")
    ]
    
    col_w = Inches(1.79)
    for i, (val, label) in enumerate(stats_data):
        c = stats_table.cell(0, i)
        c.width = col_w
        set_cell_background(c, "F1F5F9")
        clear_cell_borders(c)
        set_cell_margins(c, top=50, bottom=50, left=40, right=40)
        
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0.5)
        rv = p.add_run(val)
        rv.font.name = "Arial"
        rv.font.size = Pt(9)
        rv.font.bold = True
        rv.font.color.rgb = NAVY

        p2 = c.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(0)
        rl = p2.add_run(label)
        rl.font.name = "Calibri"
        rl.font.size = Pt(6.5)
        rl.font.bold = True
        rl.font.color.rgb = BLUE

    # Section Heading Helper
    def add_section_heading(title_text, space_before=7):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.keep_with_next = True
        
        bdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="6" w:space="1" w:color="CBD5E1"/></w:pBdr>')
        p._p.get_or_add_pPr().append(bdr)

        r = p.add_run(title_text.upper())
        r.font.name = "Arial"
        r.font.size = Pt(9.5)
        r.font.bold = True
        r.font.color.rgb = NAVY

    # Bullet helper
    def add_bullet(bold_prefix="", body_text="", space_after=1.5):
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.12
        
        if bold_prefix:
            rb = p.add_run(bold_prefix + " ")
            rb.font.name = "Calibri"
            rb.font.size = Pt(8.8)
            rb.font.bold = True
            rb.font.color.rgb = DARK
        
        if body_text:
            rt = p.add_run(body_text)
            rt.font.name = "Calibri"
            rt.font.size = Pt(8.8)
            rt.font.color.rgb = DARK
        return p

    # 3. EXECUTIVE SUMMARY
    add_section_heading("Executive Summary", space_before=6)
    p_sum = doc.add_paragraph()
    p_sum.paragraph_format.space_before = Pt(1)
    p_sum.paragraph_format.space_after = Pt(4)
    p_sum.paragraph_format.line_spacing = 1.14
    p_sum.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    r_sum = p_sum.add_run(
        "High-impact Technical Lead and Certified Odoo Consultant (Solution Architect by Practice) with 7+ years of hands-on "
        "engineering mastery leading complex Odoo customization projects and architecting bidirectional integrations between Odoo "
        "and third-party proprietary enterprise systems. Proven authority in custom Odoo development, having hand-coded 900+ custom "
        "modules spanning custom business logic, automated workflow engines, OWL 2 interfaces, and bespoke financial/operational reports "
        "for 85+ global and domestic enterprise clients. Track record includes leading 14+ major multi-version Odoo Enterprise database "
        "upgrades (v13.0 through v19.0) with zero data loss. Expert in designing resilient, fault-tolerant integration architectures "
        "(REST, XML-RPC, JSON-RPC, Webhooks, ETL queues) linking Odoo with legacy back-offices, proprietary WMS/TMS, regional marketplaces "
        "(Kogan, Tradevine), and government compliance portals (ZATCA Phase 2, NBR). Accomplished engineering leader driving hyper-velocity "
        "teams that compress 3-month complete project scopes into 2–3 weeks while enforcing zero-defect code via the Agentic E3 Framework."
    )
    r_sum.font.name = "Calibri"
    r_sum.font.size = Pt(8.8)
    r_sum.font.color.rgb = DARK

    # 4. CORE TECHNICAL & INTEGRATION COMPETENCIES
    add_section_heading("Core Technical & Integration Competencies", space_before=5)

    comp_table = doc.add_table(rows=3, cols=2)
    comp_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    comp_table.autofit = False

    comp_items = [
        ("Odoo Customization Mastery:", "900+ hand-coded modules; custom data models, business logic engines, OWL 2 reactive widgets, QWeb XML views, inheritances, compute field caching, automated server actions, cron tasks."),
        ("Proprietary & 3rd-Party Integration:", "Custom RESTful APIs, XML-RPC / JSON-RPC endpoints, Webhook consumers, event-driven sync, ETL data pipelines, token authentication, payload mapping & transformation."),
        ("Fault Tolerance & Data Reconciliation:", "Transactional idempotency, queued retries, rate limiting, error logging/alerting, conflict resolution, bidirectional catalog/inventory/order sync, zero data loss."),
        ("Performance & Scale Optimization:", "ORM batching & prefetching, eliminating N+1 queries, PostgreSQL index design, _read_group SQL aggregations, Gunicorn worker isolation (back-office vs API/web traffic)."),
        ("Multi-Version Database Migrations:", "14+ Enterprise upgrades (v13-v17, v14-v19, v16-v19), structural data retention, legacy workflow extraction, custom migration scripts, financial ledger reconciliation."),
        ("Tech Leadership & Project Delivery:", "Team leadership & developer mentorship, technical solution architecture, client stakeholder management, 6-milestone delivery model, CI/CD, Docker deployment automation.")
    ]

    for idx, (head, desc) in enumerate(comp_items):
        row = idx // 2
        col = idx % 2
        cell = comp_table.cell(row, col)
        cell.width = Inches(3.58)
        clear_cell_borders(cell)
        set_cell_margins(cell, top=20, bottom=25, left=30, right=40)
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = 1.12
        
        rh = p.add_run("• " + head + " ")
        rh.font.name = "Calibri"
        rh.font.size = Pt(8.5)
        rh.font.bold = True
        rh.font.color.rgb = BLUE

        rd = p.add_run(desc)
        rd.font.name = "Calibri"
        rd.font.size = Pt(8.5)
        rd.font.color.rgb = DARK

    # 5. KEY INTEGRATION & CUSTOMIZATION DELIVERABLES
    add_section_heading("Key Integration & Customization Deliverables", space_before=5)

    add_bullet("Proprietary & Legacy System Connector Engines:", 
               "Architected bidirectional integration bridges linking Odoo ERP to clients' in-house proprietary software, warehouse management systems (WMS), and legacy SQL databases. Designed asynchronous queue handlers and data reconcilers handling catalog, on-hand inventory quants, and live order streaming with zero data divergence.")
    
    add_bullet("Multichannel Marketplace Integration Suite (Kogan & Tradevine):",
               "Spearheaded real-time API integrations connecting Odoo ERP with leading New Zealand marketplaces (Kogan and Tradevine). Handled automated product mapping, multi-warehouse stock reservation sync, order fulfillment webhooks, and tracking notification dispatch.")
    
    add_bullet("Universal Odoo Database Model Context Protocol (MCP) Server:",
               "Architected and deployed a universal MCP server exposing Odoo ORM schemas and database models safely to external AI agent platforms (Claude Code, Gemini Enterprise, Antigravity, Cline, Puku). Enabled automated external natural language querying, schema inspection, and task orchestration.")
    
    add_bullet("ZATCA Phase 2 & NBR Tax Cryptographic Invoicing Integrations:",
               "Engineered mission-critical tax compliance modules interfacing Odoo's accounting engine with government regulatory portals (ZATCA Phase 2 e-invoicing in Saudi Arabia with XML signing, ECDSA cryptographic hashing, QR codes, and Bangladesh NBR tax engines).")

    add_bullet("High-Performance Mobile & E-Commerce Integration (Flutter & Magento 2+):",
               "Built custom REST API layers in Odoo to power a cross-platform Flutter field mobile application (OpenStreetMap route tracking, offline queueing, push notifications) and previously architected high-performance mobile APIs on Magento 2+ backend systems.")

    # Explicit page break here to create a clean, dedicated 2-page split
    doc.add_page_break()

    # 6. PROFESSIONAL LEADERSHIP & CAREER HISTORY (PAGE 2)
    add_section_heading("Professional Leadership & Engineering Experience", space_before=0)

    # Role 1
    p_r1 = doc.add_paragraph()
    p_r1.paragraph_format.space_before = Pt(3)
    p_r1.paragraph_format.space_after = Pt(1)
    p_r1.paragraph_format.keep_with_next = True
    
    r_role = p_r1.add_run("Head of Technology & Lead Architect  |  Metamorphosis Ltd.")
    r_role.font.name = "Arial"
    r_role.font.size = Pt(10)
    r_role.font.bold = True
    r_role.font.color.rgb = NAVY
    
    r_date = p_r1.add_run("  [Feb 2026 – Present]")
    r_date.font.name = "Calibri"
    r_date.font.size = Pt(9)
    r_date.font.bold = True
    r_date.font.color.rgb = BLUE

    add_bullet("Technical Leadership for Complex ERP Customizations:", 
               "Directing engineering teams delivering high-consequence Odoo ERP customization projects for 85+ global and domestic enterprise clients (Bangladesh, New Zealand, Saudi Arabia, Spain, South Sudan). Personally architecting core R&D workflows, custom themes, and 3rd-party integrations.")

    add_bullet("Hyper-Velocity Team Delivery (3-Month Scope in 2–3 Weeks):", 
               "Pioneered a delivery methodology combining modular Odoo design patterns, automated QA test suites, and local AI agentic tooling (Claude Code, Antigravity, vLLM/Ollama with DeepSeek-R1/Qwen2.5), cutting project cycle times by up to 75% with near-100% first-pass UAT acceptance.")

    add_bullet("Proprietary Systems & 3rd-Party Integration Architecture:", 
               "Lead technical architect for all custom interface layers connecting Odoo with third-party proprietary systems, banking APIs, regional logistics providers, and external e-commerce engines.")

    add_bullet("Odoo Enterprise Multi-Version Upgrades (v13.0 to v19.0):", 
               "Supervised 14+ database upgrade programs across multi-version jumps (v13-v17, v14-v19, v16-v19). Developed proprietary structural data retention scripts to carry over decades of custom business data while retiring technical debt.")

    add_bullet("High-Availability Server Worker Isolation:", 
               "Architected Nginx upstream routing dividing server workers into dedicated pools (isolating CPU-intensive external API/e-commerce sync from internal administrative back-office Gunicorn workers), eliminating downtime and latency spikes.")

    # Role 2
    p_r2 = doc.add_paragraph()
    p_r2.paragraph_format.space_before = Pt(5)
    p_r2.paragraph_format.space_after = Pt(1)
    p_r2.paragraph_format.keep_with_next = True
    
    r_role2 = p_r2.add_run("Sr. Odoo Systems Engineer & DevOps Lead  |  Metamorphosis Ltd.")
    r_role2.font.name = "Arial"
    r_role2.font.size = Pt(9.5)
    r_role2.font.bold = True
    r_role2.font.color.rgb = NAVY
    
    r_date2 = p_r2.add_run("  [Oct 2023 – Jan 2026]")
    r_date2.font.name = "Calibri"
    r_date2.font.size = Pt(9)
    r_date2.font.bold = True
    r_date2.font.color.rgb = BLUE

    add_bullet("Proprietary & Marketplace API Engineering:", 
               "Designed and implemented real-time sync modules linking Odoo ERP to international e-commerce channels (Kogan, Tradevine, BNPL platforms). Built robust reconciliation jobs to auto-correct inventory drift.")

    add_bullet("Custom Odoo Module Architecture:", 
               "Hand-coded custom modules for business process automations, manufacturing shop floor operations, multi-currency accounting, and tailored client business units.")

    add_bullet("Odoo Deployment & Infrastructure Automation:", 
               "Architected automated Docker and Bash deployment pipelines managing 300+ live Odoo database instances, reducing deployment onboarding time by 85%.")

    # Role 3
    p_r3 = doc.add_paragraph()
    p_r3.paragraph_format.space_before = Pt(5)
    p_r3.paragraph_format.space_after = Pt(1)
    p_r3.paragraph_format.keep_with_next = True
    
    r_role3 = p_r3.add_run("Full Stack Odoo Developer  |  Metamorphosis Ltd.")
    r_role3.font.name = "Arial"
    r_role3.font.size = Pt(9.5)
    r_role3.font.bold = True
    r_role3.font.color.rgb = NAVY
    
    r_date3 = p_r3.add_run("  [2022 – Sep 2023]")
    r_date3.font.name = "Calibri"
    r_date3.font.size = Pt(9)
    r_date3.font.bold = True
    r_date3.font.color.rgb = BLUE

    add_bullet("Bespoke Customization to Business Requirements:", 
               "Developed hundreds of custom Odoo modules tailored strictly to client business logic, custom operational workflows, dynamic PDF reporting, and specialized portal features.")

    add_bullet("Payment Gateways & Database Tuning:", 
               "Engineered secure payment gateway integrations (cards, mobile financial services) and optimized multi-table PostgreSQL queries and indexing for high-volume transactions.")

    # Role 4 & 5
    p_r4 = doc.add_paragraph()
    p_r4.paragraph_format.space_before = Pt(5)
    p_r4.paragraph_format.space_after = Pt(1)
    p_r4.paragraph_format.keep_with_next = True
    
    r_role4 = p_r4.add_run("Software Engineer & Systems Specialist  |  MetaKave")
    r_role4.font.name = "Arial"
    r_role4.font.size = Pt(9.5)
    r_role4.font.bold = True
    r_role4.font.color.rgb = NAVY
    
    r_date4 = p_r4.add_run("  [2020 – 2021]")
    r_date4.font.name = "Calibri"
    r_date4.font.size = Pt(9)
    r_date4.font.bold = True
    r_date4.font.color.rgb = BLUE

    add_bullet("Magento 2+ Integration & Mobile API:", 
               "Engineered backend web services and API endpoints for Magento 2+ e-commerce architectures powering high-performance mobile clients. Managed cloud servers and SSL/Nginx reverse proxies.")

    p_r5 = doc.add_paragraph()
    p_r5.paragraph_format.space_before = Pt(5)
    p_r5.paragraph_format.space_after = Pt(1)
    p_r5.paragraph_format.keep_with_next = True
    
    r_role5 = p_r5.add_run("IT Executive  |  Aptech Industries")
    r_role5.font.name = "Arial"
    r_role5.font.size = Pt(9.5)
    r_role5.font.bold = True
    r_role5.font.color.rgb = NAVY
    
    r_date5 = p_r5.add_run("  [Jun 2017 – Sep 2019]")
    r_date5.font.name = "Calibri"
    r_date5.font.size = Pt(9)
    r_date5.font.bold = True
    r_date5.font.color.rgb = BLUE

    add_bullet("Enterprise Infrastructure Administration:", 
               "Managed enterprise IT operations, Windows Server administration (Active Directory, DNS, IIS, Group Policy up to 2019), network security, and hardware/software deployments.")

    # 7. EDUCATION, CERTIFICATIONS & RESEARCH
    add_section_heading("Education, Certifications & Research", space_before=5)

    add_bullet("Certified Odoo Consultant:", "Metamorphosis Ltd. (Official Odoo Silver Partner). Extensive expertise in Odoo standard architecture, ORM, customization guidelines, and business process modeling.")
    add_bullet("Master of Information Technology (MIT):", "Jahangirnagar University (2018 – 2019). Focus on Enterprise Systems, Advanced Databases, and Distributed Architectures.")
    add_bullet("Bachelor of Science (B.Sc):", "Shahjalal University of Science and Technology (SUST). 1st Major in Mathematics, 2nd Major in Computer Science & Engineering (CSE).")
    add_bullet("IEEE Research Author (2021):", "Published paper on 'Performance Enhancement of Hadoop MapReduce' introducing the IMC algorithm (65% data processing efficiency improvement). Applied core distributed computing principles to optimize high-volume Odoo database operations.")

    # Footer note
    p_foot = doc.add_paragraph()
    p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_foot.paragraph_format.space_before = Pt(6)
    p_foot.paragraph_format.space_after = Pt(0)
    rf = p_foot.add_run("Md Niaj Shahriar Shishir  |  Tech Lead — Odoo Customization & Integrations  |  Portfolio: nsrshishir.github.io/nsrshishir")
    rf.font.name = "Calibri"
    rf.font.size = Pt(7.5)
    rf.font.color.rgb = MUTED

    doc.save(output_path)
    print(f"Successfully generated: {output_path}")

if __name__ == "__main__":
    output_docx = "/home/niaj/Documents/nsrshishir/Md_Niaj_Shahriar_Shishir_Tech_Lead_Odoo_Customization_Integrations.docx"
    create_tech_lead_cv(output_docx)
