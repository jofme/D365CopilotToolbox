# Getting Started

This guide walks you through the prerequisites, installation, and initial configuration of the D365 Copilot Toolbox.

The Copilot Toolbox enables **multi-agent workflows** in D365 Finance & Operations. This guide focuses on the core Copilot Studio agent integration solution, which provides the foundation for embedding AI agents into your ERP forms and processes.

## Prerequisites

### D365 Finance & Operations
- **Version:** 10.0.45 or later
- **Platform:** Cloud-hosted or locally deployed development environment
- Access to **System Administration** module

### Microsoft Copilot Studio
- A published **Copilot Studio agent** (formerly Power Virtual Agent)
- The agent must be deployed to a **Dataverse environment**
- You will need:
  - **Dataverse environment GUID** (found in Power Platform Admin Center or Copilot Studio agent settings)
  - **Agent schema name** (found in Copilot Studio under agent settings)

### Microsoft Entra ID (Azure AD)
- An **App Registration** configured as a **Single Page Application (SPA)** — public client, no client secret required
- The app registration needs:
  - **Redirect URI:** `https://<your-d365-environment-url>` (the origin of your D365 instance)
  - **API Permission:** Power Platform API > **Application** > `CopilotStudio.Copilots.Invoke` (requires admin consent)
- You will need:
  - **Application (client) ID**
  - **Tenant ID**

## Installation

### Deploy via Deployable Package

1. Download the latest release package from the [Releases](../../releases) page
2. Deploy the package to your D365 environment using LCS or PPAC
3. After deployment, the models **Copilot Toolbox** and (optionally) **Copilot Toolbox Examples** will be available

## Initial Configuration

### 1. Assign Security Roles

Navigate to **System Administration > Users** and assign the appropriate roles:

| Role | Purpose |
|------|---------|
| **Copilot Administrator** (`COTXCopilotAdminRole`) | Configure agent parameters and use the side panel |
| **Copilot User** (`COTXCopilotUserRole`) | Use the Copilot side panel |

### 2. Configure Agent Parameters

1. Navigate to **System Administration > Setup > Copilot Toolbox > Agent Parameters**
2. Click **New** to create a parameter record
3. Fill in the following fields:

| Field | Description | Example |
|-------|-------------|---------|
| **Name** | A unique identifier for this configuration | `Production Agent` |
| **Description** | Human-readable description | `Main Copilot Studio agent for FSCM` |
| **Entra ID Tenant** | Your Azure AD tenant ID | `12345678-abcd-...` |
| **Entra ID App Registration** | The SPA app registration client ID | `87654321-dcba-...` |
| **Agent Schema Name** | The Copilot Studio agent schema name | `cr123_myAgent` |
| **Dataverse Environment** | The Dataverse environment GUID | `a1b2c3d4-1234-5678-90ab-cdef12345678` |
| **Send Global FSCM Context** | Enable to send navigation context to the agent | `Yes` |
| **Show tool usage** | Show tool call cards in the chat UI | `Yes` |
| **Show thoughts** | Display the agent's chain-of-thought reasoning in the chat | `No` |
| **Keep connection alive** | Preserve the connection to the Copilot Studio agent. Reasoning will be showed in the console  | `No` |

### 3. Map Application Areas

On the **Available In** tab of the Agent Parameters form:

1. Click **New**
2. Select an **Application Area** (e.g., `Side Panel`, `Fallback`)
3. Repeat for each area where this agent should be active

> **Tip:** The `Fallback` application area is used when no specific mapping exists for a requested area. Always configure at least a Fallback mapping.

### 4. Test the Side Panel

1. Open any form in D365 (e.g., **All Sales Orders**)
2. Click the **Settings** gear icon in the navigation bar
3. Select **Copilot Agent**
4. The Copilot chat panel should appear on the right side
5. Type a message to verify the connection

## Troubleshooting

| Issue | Cause | Solution |
|-------|-------|---------|
| Side panel is empty / no chat | Missing or incorrect agent parameters | Verify all fields in Agent Parameters |
| Authentication popup appears | Expected on first use | Sign in; subsequent requests use silent token acquisition |
| "AADSTS..." error in popup | App registration misconfigured | Check redirect URI matches your D365 origin, verify API permissions |
| No context sent to agent | `Send Global FSCM Context` is disabled | Enable it in Agent Parameters |
| Control doesn't appear on form | Missing security role | Assign `Copilot User` or `Copilot Administrator` role |

## Next Steps

- [Architecture](architecture.md) — understand how the control works internally
- [Configuration](configuration.md) — detailed parameter reference
- [Extending](extending.md) — add agents to your own forms
- [Security](security.md) — security model details
