# HMS Enterprise — CLAUDE.md

## Project
HMS Enterprise is a multi-tenant SaaS Hospital Management System by TGWT Solutions (TechDigital WishTree). 12 integrated modules covering OPD, Pharmacy, Lab, Finance, HR, Inventory, BI, Patient Portal, Audit, IPD, Appointments, and Platform Admin.

## Tech Stack
- **Backend:** Spring Boot + Apache Kafka (event-driven) + PostgreSQL (schema-isolated multi-tenancy)
- **Frontend:** React + PWA (Patient Portal)
- **Infrastructure:** AWS EKS, Terraform, Ansible, Redis, Elasticsearch
- **Security:** JWT + MFA, 6-tier RBAC, DISHA/ABDM/NABH compliant

## Repository
- **Remote:** `github-aksatyam:aksatyam/hms-tenent.git` (private)
- **Branch:** `main` (direct push, no PR workflow)
- **SSH:** Uses `github-aksatyam` host alias

## Key Conventions
- HTML artifacts: Interactive, visually polished, navy (#1B3A5C) / teal (#0D7377) / gold (#C49A2A) palette
- Documents marked: TGWT-HMS-ENT-2026-XXX, "CONFIDENTIAL — Enterprise SaaS"
- Marketing references: "50+ hospitals", "12 integrated modules", "64 screens", ABDM/ABHA compliant
- Languages: English (primary), Hindi (regional marketing)
- Always commit and push completed artifacts immediately

## File Structure
- `docs/enterprise/` — SOW, project plans, build plans
- `docs/research/` — PRD, TSD, API specs, product strategy
- `docs/presentations/` — Executive decks, pitch decks
- `wireframes/` — Interactive HTML wireframes and workflows (64+ screens)
- `marketing/` — WhatsApp templates, campaign assets
- `marketing/assets/` — Marketing images (WhatsApp, hero, etc.)
- `scripts/` — Python document generators (python-docx based)
- `archives/` — Zip backups
