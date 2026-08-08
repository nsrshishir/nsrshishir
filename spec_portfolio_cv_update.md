# Design Specification: Portfolio Website & CV Update for Md Niaj Shahriar Shishir

**Date:** 2026-08-08  
**Author:** Antigravity AI  
**Target Directory:** `/home/niaj/Documents/nsrshishir`

---

## 1. Executive Summary

This specification outlines the complete refresh of Niaj Shishir's personal portfolio website (`index.html`) and print-ready CV (`cv_final.html` -> `Md_Niaj_Shahriar_Shishir_Odoo_FullStack_Architect.pdf`). The update reflects recent key milestones:
* **Head of Technology & Senior Software Engineer** leadership at Metamorphosis Ltd.
* **AI Agentic Workflows & Tooling**: Creation of custom skills (`odoo-structural-data-retention`, `odoo-server-tuner`, `odoo-bd-accounting`, `odoo-19`), subagent orchestration, AGY CLI usage, prompt engineering, and token efficiency optimization.
* **Odoo 19 & Middle-Ground Migrations**: Legacy (v14-v18) to v19 structural data retention strategy (retaining models, custom fields, historical reporting, and QWeb views while dropping deprecated Python workflows).
* **Enterprise Quotation & Milestone Architecture**: 6-milestone operational delivery model (40/12/12/12/12/12 payment split, Go-Live at M4) across Bangladesh, New Zealand, Saudi Arabia, and South Sudan.
* **High-Performance Infrastructure & WeasyPrint PDF Pipeline**: Single-server tuning (4GB RAM memory recycling, deadlock prevention) and automated document rendering.

---

## 2. Portfolio Website Redesign (`index.html`)

### 2.1 Aesthetic & UI/UX Principles
* **Dual Theme Switcher**: Toggle seamlessly between a sleek Glassmorphism Dark Mode (Navy `#0f172a`, Deep Slate `#1e293b`, Electric Blue `#38bdf8`, Violet Accent `#818cf8`) and a Clean Light Mode (`#ffffff`, Light Slate `#f8fafc`, Royal Blue `#2563eb`). Choice saved in `localStorage`.
* **Typography & Styling**: Google Fonts (`Inter` & `Outfit`), clean font hierarchy, glowing glass cards, subtle micro-animations (hover elevation, gradient glows).
* **Responsive Layout**: Mobile-first responsive navigation with drawer menu.

### 2.2 Section Breakdown
1. **Header & Navigation**:
   * Brand Logo (`NIAJ.SHISHIR`), navigation links (About, Core Expertise, AI & Workflow Tools, Featured Projects, Career Timeline, Contact).
   * Theme Toggle switch (Sun/Moon icon).
2. **Hero Section**:
   * Headline: **Head of Technology & Odoo/AI Architect**
   * Subtitle: *"Pioneering AI-driven ERP Architectures, Odoo v12-v19 Structural Migrations, & Scalable Enterprise Integrations."*
   * Action buttons: `Explore Portfolio`, `Download CV (PDF)` (linking to fresh WeasyPrint compiled PDF).
3. **Interactive Metrics / Stats Bar**:
   * `135+` ERP Implementations
   * `85+` Global Enterprise Clients
   * `15+` Version Upgrades (v12-v19)
   * `10+` Custom AI Skills & Workflows
4. **About & Leadership Vision**:
   * Profile highlight (using `image_niaj_blue_suit.png`).
   * Breakdown of technology leadership at Metamorphosis Ltd (Bangladesh Odoo Silver Partner), international client reach (BD, NZ, KSA, South Sudan), and AI integration philosophy.
5. **Core Technical Competencies**:
   * **Odoo 19 & Structural Migration**: `models.Index`, `models.Constraint`, OWL 2 components, dummy-module retention.
   * **AI Agentic Systems**: AGY CLI subagent pipelines, custom skill development, automated code review & token efficiency.
   * **Global Fintech & Compliance**: ZATCA Phase 2 (KSA), NBR/DataSoft (BD), 13+ Payment Gateways (Afterpay, Zip, Laybuy, SSLCommerz, bKash).
   * **DevOps & Performance Tuning**: Low-spec single-server tuning, memory recycling, WeasyPrint PDF generation pipelines.
6. **Featured Projects (Filterable & Detailed)**:
   * **AI Agent & Odoo Migration Suite**: Structural Data Retention tool & dummy-module generator.
   * **ZATCA & NBR Regulatory Engine**: Phase 2 Saudi E-invoicing & BD tax compliance.
   * **Global Fintech & BNPL Hub**: Multi-gateway payment orchestration for NZ & global retail.
   * **Single-Server Performance Tuner**: Odoo 4GB RAM optimization & WeasyPrint automation.
   * **MERN + Odoo Headless eCommerce**: Sub-second React storefront.
7. **Career Timeline**:
   * **Head of Technology & Lead Architect** | Metamorphosis Ltd (2024 - Present)
   * **Sr. Odoo Systems Engineer & DevOps Lead** | Metamorphosis Ltd (2023 - 2024)
   * **Full Stack Developer** | Metamorphosis Ltd (2022 - 2023)
   * **Software Engineer & Server Specialist** | MetaKave (2020 - 2021)
8. **Footer & Contact**:
   * Contact details (`nsrshishir550@gmail.com`, `info@metamorphosis.com.bd`), LinkedIn, GitHub.

---

## 3. Print-Ready CV Redesign (`cv_final.html` -> PDF)

### 3.1 Template Architecture
* Strictly formatted 2-page A4 document using WeasyPrint-compatible CSS (`@page { size: A4; margin: 0; }`).
* Executive palette (Navy, Accent Blue, Cool Slate).
* Header with photo (`image_niaj_blue_suit.png`), contact info, role title (**Head of Technology & Odoo/AI Architect**).

### 3.2 PDF Generation Workflow
* Render `cv_final.html` to `Md_Niaj_Shahriar_Shishir_Odoo_FullStack_Architect.pdf` using `/home/niaj/odoo/venv17/bin/python -m weasyprint cv_final.html Md_Niaj_Shahriar_Shishir_Odoo_FullStack_Architect.pdf`.
* Verify exact 2-page layout without overflow or blank pages.

---

## 4. Implementation Steps

1. Update `cv_final.html` with new title, summary, AI agent workflow experience, Odoo 19 structural migration details, and updated metrics.
2. Compile `cv_final.html` into `Md_Niaj_Shahriar_Shishir_Odoo_FullStack_Architect.pdf`.
3. Rewrite `index.html` with new dual-theme glassmorphism design, interactive features, responsive navigation, updated content, and direct PDF download link.
4. Verify both files in browser / WeasyPrint build.
