# Security and Privacy Policy

## EmmaNeigh - Transaction Management & Document Automation

**Last Updated:** March 23, 2026

## Overview

EmmaNeigh is a desktop application for transaction management, document automation, email-assisted workflows, and AI-assisted task execution. The app performs most document work on the local machine, but the current product can also use configured network services for updates, telemetry, access control, AI inference, and document-management integrations.

This document is intended for:
- IT security teams
- knowledge management teams
- risk and compliance teams
- deployment administrators
- end users who want to understand how the app handles data

## What Happens Locally

The following categories of work are designed to run on the user's machine:
- document parsing and redline orchestration
- Litera desktop invocation
- Outlook desktop automation
- Word desktop automation
- local file packaging, zipping, export, and output generation
- optional local AI inference when a local runtime is configured

EmmaNeigh stores some operational data locally, including:
- local app settings
- session state
- local usage history and retry queues
- feedback log cache
- generated outputs and temporary working files created during workflows

## What May Use the Network

Depending on configuration, EmmaNeigh may communicate with:
- GitHub Releases for app update checks and downloads
- a telemetry ingest backend or Firebase Firestore for login/activity tracking
- an access-policy backend for remote allow/block decisions and kill-switch behavior
- configured AI providers such as Anthropic, OpenAI, Harvey, Ollama, LM Studio, or other local-compatible runtimes
- other deployment-specific services explicitly configured by the administrator

EmmaNeigh does **not** claim to be an air-gapped or fully offline application.

## Data Categories

EmmaNeigh may process or store the following categories of data during normal use:
- user email address and display name for login identity and analytics
- usage telemetry such as feature used, timestamps, durations, model/provider, and counts
- feedback submitted through the app
- email metadata and message content when email-driven workflows are used
- document names, file paths, temporary copies, and generated outputs
- AI prompts, attachments, and extracted content for workflows that use an external model or proxy

EmmaNeigh is not intended to transmit document contents to third parties unless the deployment administrator or end user explicitly configures an AI provider, telemetry service, or other network integration that requires it.

## Authentication and Access Control

Current EmmaNeigh builds support lightweight email-based identity for ordinary users. This is intended primarily for user tracking and operational gating.

Admin capabilities should be managed through deployment configuration and remote access policy, not by user-entered local settings. Deployments should use:
- remote access policy for allow/block decisions
- managed admin claims or allowlists
- version enforcement and remote disable capability
- enterprise SSO or equivalent identity in future production deployments

## Telemetry and Analytics

EmmaNeigh can record:
- login events
- feature usage events
- session heartbeats
- feedback submissions
- selected operational diagnostics

Telemetry can be configured in one of these ways:
- managed deployment configuration bundled with the app or set by environment
- local administrator override in unmanaged deployments
- disabled, if no telemetry backend is configured

Telemetry content should be limited to operational metadata unless the deployment intentionally enables richer logging.

## AI Providers and Data Egress

EmmaNeigh is model-agnostic. If an AI workflow is routed to:
- a local runtime, prompts remain on the local machine except for update/download traffic related to the runtime itself
- a remote provider, the prompt payload and any attached workflow context sent to that provider leave the local machine

Deployment administrators should explicitly control:
- which providers are allowed
- which models are allowed
- whether cloud inference is permitted at all
- whether prompt logging is enabled

## Secret Handling

EmmaNeigh uses Electron's OS-backed secure storage when available for:
- API keys
- proxy tokens
- telemetry tokens
- access-policy tokens
- other sensitive secrets

If secure OS-backed secret storage is unavailable, enterprise deployments should treat that machine as unsupported for secret-backed features.

## Local Files and Temporary Data

EmmaNeigh creates local working files for some workflows. Those may include:
- exported documents
- redline outputs
- packet folders
- generated checklists or punchlists
- temporary intermediate files produced by local processors or desktop integrations

Administrators should apply their own workstation controls for:
- disk encryption
- endpoint protection
- local retention policy
- profile separation on shared machines

## Recommended Enterprise Controls

For enterprise deployment, EmmaNeigh should be operated with:
- managed configuration for telemetry and access policy
- remote access control and kill-switch support
- a current supported Electron runtime
- code-signed installers and notarized releases where applicable
- strict AI provider allowlists
- centralized audit logging
- workstation disk encryption and endpoint security
- enterprise identity/SSO for admin and privileged actions

## Current Limitations

EmmaNeigh integrates with several desktop applications and local environments. Reliability and permissions can vary by firm environment, including:
- Outlook COM availability
- Litera desktop installation and policy
- iManage COM, Drive, or API availability
- local AI runtime availability
- firewall or proxy restrictions

Some enterprise controls depend on deployment setup outside the app itself.

## Incident Response and Support

If a security or privacy issue is identified:
1. stop using the affected workflow if the issue could expose confidential data
2. preserve logs or screenshots needed for diagnosis
3. notify the deployment administrator or maintainer
4. rotate any affected secrets or API keys
5. review the configured telemetry, AI, and policy endpoints involved

## Source and Review

Repository:
- https://github.com/raamtambe/EmmaNeigh

Key files to review:
- `desktop-app/main.js`
- `desktop-app/preload.js`
- `desktop-app/firebase-config.js`
- `desktop-app/index.html`
- `desktop-app/python/`

