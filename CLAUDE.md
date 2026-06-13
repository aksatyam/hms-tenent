# HMS / HMSPro — CLAUDE.md

## Project
**HMSPro Product Family** is a multi-tenant SaaS Hospital Management System by **VSJ AI Labs Private Limited**. One shared core platform; **five editions** (Hospital · Clinic/CMS · Pharmacy SME · Pathology SME · Radiology SME) composed from **13 modules (M01–M13)** plus a locked-on core (Identity/Auth, Registry/UHID, M04 Finance, M09 Audit, M12 Platform Admin, Notifications). Editions are entitlement bundles (not forks); upgrade Pharmacy→Clinic→Hospital with zero data migration.

> **Attribution:** Author = Ashish Kumar Satyam (Founder & Lead Architect); for VSJ AI Labs Pvt. Ltd.; "Confidential — VSJ AI Labs Internal". **Never reference TechDigital WishTree / TGWT** (per global rule). The earlier "by TechDigital WishTree / 12 modules / TGWT-HMS-ENT doc IDs" wording was stale and has been corrected.

## Tech Stack
- **Backend:** Spring Boot + Apache Kafka (event-driven, partition key = `tenantId`) + PostgreSQL
- **Tenancy:** Tiered isolation — Starter=Pool/RLS · Professional=Schema+RLS · Enterprise=Dedicated DB; RLS fail-closed in all tiers; entitlements enforced at 4 layers (gateway→JWT→service-guard→UI)
- **Frontend:** React + PWA (Patient Portal)
- **Infrastructure:** AWS EKS, Terraform, Ansible, Redis, Elasticsearch
- **Security:** JWT + MFA, 6-tier RBAC, DPDP Act 2023/ABDM/NABH compliant; dual breach clocks (CERT-In 6h / DPDP 72h); append-only hash-chained audit

## Repository
- **Remote:** `github-aksatyam:aksatyam/hms-tenent.git` (private)
- **Branch:** `main` (direct push, no PR workflow)
- **SSH:** Uses `github-aksatyam` host alias

## Key Conventions
- Marketing/wireframe HTML: Interactive, visually polished, navy (#1B3A5C) / teal (#0D7377) / gold (#C49A2A) palette
- **Architecture/spec docs** (BRD/FSD/HLD/LLD): interactive self-contained HTML in the `HMS-Workflow-Enhanced` house style — fonts Sora / JetBrains Mono / Crimson Pro; palette navy `#0D1F35` / teal `#0E8E8E`+`#14B8B8` / gold `#C49A2A`; fixed sidebar + hero + arch-strip; inline SVG for C4 / sequence / state-machine diagrams. Keep HLD and LLD as **separate** documents.
- Spec doc-ID scheme: `XXX-003-HMSPRO-2026` (e.g. BRD-003, FSD-003, HLD-003, LLD-003); marking "Confidential — VSJ AI Labs Internal"
- Char hygiene: never use `§` (write `Sec.`); never emoji circles in DOCX; verify no TGWT references before shipping
- Marketing references: "13 integrated modules", ABDM/ABHA compliant
- Languages: English (primary), Hindi (regional marketing)
- Always commit and push completed artifacts immediately

## File Structure
- `HMS-BRD-WorkFlow-FSD/` — canonical spec set: BRD-003, FSD-003, **HLD-003** (C4 + tenancy + entitlement + events + ADRs), **LLD-003** (schema + per-module state/API/event/rules + sequences + error catalog + RTM), HMS-Workflow-Enhanced v2.2. Referenced-external: TAD-003 (tenancy primitives), per-module OpenAPI (next artefact).
- `docs/enterprise/` — SOW, project plans, build plans
- `docs/research/` — PRD, TSD, API specs, product strategy
- `docs/presentations/` — Executive decks, pitch decks
- `wireframes/` — Interactive HTML wireframes and workflows (64+ screens)
- `marketing/` — WhatsApp templates, campaign assets
- `marketing/assets/` — Marketing images (WhatsApp, hero, etc.)
- `scripts/` — Python document generators (python-docx based; only `generate_workflow_document.py` remains)
- `archives/` — Zip backups
