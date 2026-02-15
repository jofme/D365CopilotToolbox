# Configuration

This document provides a detailed reference for configuring the D365 Copilot Toolbox.

## Agent Parameters

Navigate to **System Administration > Setup > Copilot Toolbox > Agent Parameters** to manage agent configurations.

### Copilot Studio Tab

#### Entra ID Settings

| Field | EDT | Required | Description |
|-------|-----|----------|-------------|
| **Entra ID Tenant** | `COTXCopilotEntraIdTenantId` | Yes | The Azure AD / Entra ID tenant GUID. Found in Azure Portal > Microsoft Entra ID > Overview. |
| **Entra ID App Registration** | `COTXCopilotEntraIdAppRegId` | Yes | The Application (client) ID of the SPA app registration. Found in Azure Portal > App registrations. |

#### Dataverse Settings

| Field | EDT | Required | Description |
|-------|-----|----------|-------------|
| **Agent Schema Name** | `COTXCopilotAgentSchemaName` | Yes | The schema name of your Copilot Studio agent. Found in Copilot Studio > Agent settings > Advanced. Format: `cr123_agentName` |
| **Dataverse Environment** | `COTXCopilotAgentDataverseEnvironment` | Yes | The Dataverse environment GUID. Found in Power Platform Admin Center > Environments > [Environment Name] > Environment details, or in Copilot Studio agent settings. Format: `a1b2c3d4-1234-5678-90ab-cdef12345678` |

#### Context Settings

| Field | EDT | Description |
|-------|-----|-------------|
| **Send Global FSCM Context** | `COTXCopilotSendGlobalContext` | When enabled (`Yes`), the global side panel sends ERP context (legal entity, current form, record info) with each message to the agent. |
| **Show tool usage** | `COTXCopilotHostShowToolUsage` | When enabled (`Yes`), tool call details are displayed as Adaptive Cards in the chat, showing which tools the agent invoked and its reasoning. |

### Available In Tab

This tab maps application areas to the current agent configuration. Each row associates a `COTXCopilotAgentApplicationArea` enum value with this parameter record.

| Application Area | Description |
|-------------------|-------------|
| **Fallback** | Default fallback agent. Used when no specific mapping exists for a requested application area. |
| **Side Panel** | The global Copilot side panel accessible from the Settings menu. |
| *(Custom areas)* | Additional areas defined via enum extensions in other models. |

> **Important:** If a control requests an application area that has no mapping, the system falls back to the `Fallback` area. Always ensure a Fallback mapping exists.

## Entra ID App Registration Setup

### Step 1: Create the App Registration

1. Go to [Azure Portal](https://portal.azure.com) > **Microsoft Entra ID** > **App registrations**
2. Click **New registration**
3. Configure:
   - **Name:** `D365 Copilot Toolbox` (or your preferred name)
   - **Supported account types:** Accounts in this organizational directory only (Single tenant)
   - **Redirect URI:** Select **Single-page application (SPA)** and enter your D365 environment origin URL (e.g., `https://yourenv.operations.dynamics.com`)
4. Click **Register**

### Step 2: Configure API Permissions

1. In the app registration, go to **API permissions**
2. Click **Add a permission** > **APIs my organization uses**
3. Search for **Power Platform API**
4. Select **Application permissions**
5. Check **CopilotStudio.Copilots.Invoke** (Allows Invoking Copilots)
6. Click **Add permissions**
7. Click **Grant admin consent**

### Step 3: Note the Values

| Value | Where to Find | Maps To |
|-------|---------------|---------|
| Application (client) ID | App registration > Overview | **Entra ID App Registration** field |
| Directory (tenant) ID | App registration > Overview | **Entra ID Tenant** field |

## Copilot Studio Agent Setup

### Agent Requirements

The Copilot Studio agent should be:
- **Published** to a Dataverse environment
- Configured with appropriate **topics** and **actions** (tools)
- Optionally configured to read the `channelData.context` from incoming messages

### Reading ERP Context in the Agent

When `Send Global FSCM Context` is enabled, every user message includes a `context` object in `channelData`. In Copilot Studio, you can access this via:

1. **Power Automate actions** that read the conversation context
2. **Custom topics** that extract context variables
3. **Plugin actions** that receive the full turn context

The context structure is:

```json
{
  "userLanguage": "en-us",
  "userTimeZone": "GMT Standard Time",
  "legalEntity": "USMF",
  "currentUser": "Admin",
  "currentForm": "All Sales Orders",
  "currentMenuItem": "Sales order",
  "currentRecord": {
    "tableName": "Sales order",
    "naturalKey": "Sales order",
    "naturalValue": "SO-000123"
  }
}
```

## Multi-Agent Configuration

You can configure multiple agent parameter records, each connected to a different Copilot Studio agent and mapped to different application areas.

**Example setup:**

| Parameter Name | Agent | Application Areas |
|---------------|-------|-------------------|
| General Assistant | `cr123_generalAgent` | Fallback, Side Panel |
| Sales Agent | `cr456_salesAgent` | SalesTable |
| Inventory Agent | `cr789_invAgent` | InventOnhand |

When a form control requests the `SalesTable` area, it gets the Sales Agent. When it requests an unmapped area, it falls back to the General Assistant.

