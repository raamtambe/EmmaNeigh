# Law Firm Acquisition Leads

This folder contains a first-pass acquisition target list sourced from public firm websites and search results, then cross-checked against existing Dropbox sourcing files.

## What is in here

- `law_firm_acquisition_targets.csv`: scored target list with public-source notes.
- `hubspot_public_email_candidates.csv`: subset with public email addresses that can be screened for HubSpot import.
- `outreach_drafts.md`: draft outreach emails to review before any send.

## Scoring logic

- Practice fit: heavier weight on estate planning, probate, family law, bankruptcy, real estate, and criminal defense.
- Regulatory fit: Arizona scores better because outside-capital / non-lawyer ownership structures are more workable there than in Florida or California.
- Founder transition readiness: higher score for older founders or founders with long public biographies suggesting retirement-age timing.
- Firm size: stronger scores for multi-attorney boutiques that look big enough to support roughly 1M of EBITDA but not so large that they are obviously outside the target.
- Founder ownership: founder still active and prominently associated with the brand.
- Contactability: public email available for direct CRM and outreach use.

## Existing Dropbox overlaps already found

- `Karp Law Firm` already appears in prior Dropbox sourcing files. One older internal row estimated 32 attorneys and about 4.1M of EBITDA, which does not match the smaller public team listing and should be reconciled before import.
- `The Law Offices of C.R. Abrams P.C.` already appears in prior Dropbox California probate exports and should be checked for prior outreach history before re-import.

## Regulatory note

As of April 2, 2026, Arizona is the easiest of the three preferred states for a direct acquisition thesis involving non-lawyer ownership. Florida and California remain more constrained and likely require a lawyer-owned or otherwise compliant structure.

Useful official links:

- Arizona ABS overview: https://www.azcourts.gov/cld/Alternative-Business-Structures
- California Rule 5.4: https://www.calbar.ca.gov/Portals/0/documents/rules/Rule_5.4-Exec_Summary-Redline.pdf
- Florida Rule 4-5.4 materials: https://www.floridabar.org/the-florida-bar-news/nonlawyer-owned-law-firms-a-likely-ethics-trap-for-florida-lawyers/
