# D365 Copilot Toolbox

A community-driven framework for enabling **multi-agent workflows** in **Dynamics 365 Finance & Operations**. This repository provides solutions, examples, and integrations that bring AI agents into D365 F&O to enhance business processes and user productivity.

[![Build Status](../../actions/workflows/build.yml/badge.svg)](../../actions/workflows/build.yml)
![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)
![Platform: D365 F&O](https://img.shields.io/badge/Platform-D365%20F%26O-blue.svg)
![Version: 10.0.45](https://img.shields.io/badge/Minimum%20Version-10.0.45-lightgrey.svg)

> **By the community, for the community!**
>
> Inspired by the [D365FOAdminToolkit](https://github.com/ameyer505/D365FOAdminToolkit) project.
> Initial idea by Lone Gru (lone.gru@arineo.com)

---

## Overview

The D365 Copilot Toolbox is designed to enable **multi-agent workflows** within D365 Finance & Operations. The project includes multiple solutions and examples that demonstrate how to integrate AI agents into ERP business processes.

### Core Solution: Copilot Studio Agent Integration

The primary solution provides a ready-to-use extensible control (`COTXCopilotHostControl`) that embeds Microsoft Copilot Studio agents directly inside any D365 F&O form — either as a **global side panel** or as a **local tab embedded on a specific form**. The control handles:

- **MSAL.js browser-side authentication** — application tokens with admin consent, no server-side secrets
- **Copilot Studio Agent SDK** — connection via `@microsoft/agents-copilotstudio-client`
- **Automatic ERP context injection** — legal entity, current form, record, navigation info sent as channel data
- **Tool call visualization** — optional Adaptive Card display of agent tool execution
- **Event-driven responses** — X++ delegates let form code react to agent replies
- **Multi-agent routing** — configure different agents for different application areas

## Key Capabilities

### Multi-Agent Workflows
- **Application Area Routing** — Configure separate agents per business area (Sales, Inventory, Finance, etc.)
- **Context-Aware Agents** — Agents receive ERP context (legal entity, form, record, user info)
- **Agent Orchestration** — Coordinate multiple agents working together on complex workflows
- **Extensible Architecture** — Add new agents and application areas without modifying core code

### Integration Features

| Feature | Description |
|---------|-------------|
| **Side Panel Agent** | Global Copilot pane available from Settings menu across all forms |
| **Embedded Agent** | Drop the control onto any form tab with local context scope |
| **Global Context Tracking** | Side panel automatically tracks form navigation and sends context |
| **Local Context Tracking** | Embedded controls subscribe to their form's data source changes |
| **Role-Based Security** | Admin and User roles with dedicated duties and privileges |
| **Tool Call Visualization** | See which tools agents invoke and their reasoning |

## Current Solutions

This repository currently includes:

1. **Copilot Studio Agent Integration** — The core framework for embedding Microsoft Copilot Studio agents into D365 F&O forms (detailed above)
2. **SalesTable Example** — Reference implementation showing agent integration on the Sales Order form

### Future Solutions & Examples

The project is designed to grow with contributions from the community. Future solutions may include:
- Agent orchestration patterns for complex workflows
- Pre-built agents for common ERP tasks
- Integration with other AI services and platforms
- Industry-specific agent templates

## Quick Start

See the [Getting Started Guide](docs/getting-started.md) for prerequisites, installation, and initial configuration.

## Documentation

| Document | Description |
|----------|-------------|
| [Getting Started](docs/getting-started.md) | Prerequisites, installation, and first-time setup |
| [Architecture](docs/architecture.md) | Technical design, control lifecycle, and data flow |
| [Configuration](docs/configuration.md) | Agent parameter setup, Entra ID, Copilot Studio connection |
| [Extending](docs/extending.md) | Adding agents to custom forms, new application areas, handling responses |
| [Security](docs/security.md) | Roles, duties, privileges, and Entra ID app registration |
| [Contributing](CONTRIBUTING.md) | How to contribute, coding standards, and PR guidelines |
| [Changelog](CHANGELOG.md) | Version history and release notes |

## Project Structure

```
D365CopilotToolbox/
├── Metadata/
│   ├── CopilotToolbox/              # Main model (module: CopilotToolbox)
│   │   └── Copilot Toolbox/
│   │       ├── AxClass/             # X++ classes (host control, context tracking)
│   │       ├── AxTable/             # Agent parameters & application area tables
│   │       ├── AxForm/              # Parameters form & side panel
│   │       ├── AxEdt/               # Extended data types
│   │       ├── AxEnum/              # Application area & context scope enums
│   │       ├── AxResource/          # HTML, JS, CSS for the browser control
│   │       ├── AxSecurityRole/      # Admin & User roles
│   │       ├── AxSecurityDuty/      # Admin & User duties
│   │       ├── AxSecurityPrivilege/ # Parameter & side panel privileges
│   │       ├── AxMenuExtension/     # Settings menu & System Administration integrations
│   │       ├── AxMenuItemDisplay/   # Menu items for forms
│   │       └── AxLabelFile/         # en-US localizable labels
│   └── CopilotToolboxExamples/      # Example model (module: CopilotToolboxExamples)
│       └── Copilot Toolbox Examples/
│           ├── AxEnumExtension/     # Example: SalesTable application area
│           └── AxFormExtension/     # Example: Copilot on SalesTable form
├── Project/
│   ├── CopilotAgentHost/            # VS project for main model
│   └── CopilotHostExamples/         # VS project for examples model
├── Scripts/
│   ├── RegisterSymbolicLinks.ps1    # Create symlinks for local development
│   └── UnregisterSymbolicLinks.ps1  # Remove symlinks
├── Templates/                       # (Reserved for future templates)
├── .github/                         # CI/CD build workflows (FSC-PS)
└── .FSC-PS/                         # FSC-PS configuration
```

## Models

The repository is organized into models that contain different solutions and examples:

| Model | Module | Description | Dependencies |
|-------|--------|-------------|--------------|
| **Copilot Toolbox** | `CopilotToolbox` | Core framework for Copilot Studio agent integration — control, tables, security, resources | ApplicationFoundation, ApplicationPlatform, ApplicationSuite |
| **Copilot Toolbox Examples** | `CopilotToolboxExamples` | Example implementations and reference patterns (currently includes SalesTable integration) | ApplicationPlatform, ApplicationSuite, CopilotToolbox |

Additional models for future solutions will be added as the project grows.

## Contact & Contribution

Want to contribute? We'd **LOVE** to have you! See the [Contributing Guide](CONTRIBUTING.md) for details.

There are several ways to contribute:
- Add new features
- Report bugs
- Request new features
- Documentation

Feel free to reach out to **copilot@erpilots.com** for any questions or feature suggestions.

## Legal

This project is licensed under the [MIT License](LICENSE).

All features are provided **as-is**, without support or warranty of any kind.

## Build

The GitHub build action is based on [FSC-PS for GitHub](https://github.com/fscpscollaborative/fscps).

**Target version:** 10.0.45 GA
