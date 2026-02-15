# Architecture

This document describes the technical architecture of the D365 Copilot Toolbox core solution for integrating Microsoft Copilot Studio agents into D365 Finance & Operations.

The Copilot Toolbox is designed to enable **multi-agent workflows** in D365 F&O. This architecture focuses on the foundational Copilot Studio integration, which provides the framework for embedding agents, routing to different agents based on application areas, and managing context flow between D365 and AI agents.

## High-Level Architecture

```mermaid
graph TB
    subgraph D365["D365 Finance & Operations (Browser)"]
        subgraph Server["X++ Server-Side"]
            Control["COTXCopilotHostControl<br/>(FormTemplateControl)"]
            FormCtx["COTXCopilotHostFormContext"]
            GlobalCtx["COTXCopilotHostGlobalContext"]
        end
        subgraph Browser["Browser-Side (HTML/JS/CSS)"]
            JS["COTXCopilotHostControl.js"]
            MSAL["MSAL.js<br/>(Auth)"]
            WebChat["WebChat<br/>(UI)"]
            SDK["Copilot Studio SDK<br/>(DirectLine connection)"]
        end
    end
    subgraph Copilot["Microsoft Copilot Studio<br/>(Dataverse / Power Platform)"]
        Agent["Agent processes messages,<br/>executes tools, returns responses"]
    end

    Control --> JS
    MSAL --> SDK
    WebChat --> SDK
    SDK --> Agent
```

## Component Overview

### X++ Server-Side Components

| Class | Responsibility |
|-------|---------------|
| `COTXCopilotHostControl` | Main extensible form control. Reads agent configuration from the database, initializes form properties, and passes them to the browser-side JS. Handles incoming agent responses via `RaiseAgentResponse` command. |
| `COTXCopilotHostControlBuild` | Design-time companion class. Exposes `Application Area` and `Context Scope` properties in the Visual Studio form designer. |
| `COTXCopilotHostFormContext` | Tracks a single form's context: data area, form caption, menu item name, root data source table/record, and natural key/value. Fires `onFormContextChange` when the active record changes. |
| `COTXCopilotHostGlobalContext` | Singleton that subscribes to `Info.onActivate`; when the user navigates between root-navigable forms, it constructs a new `COTXCopilotHostFormContext` and propagates changes to the side panel control. |

### Browser-Side Components

| File | Responsibility |
|------|---------------|
| `COTXCopilotHostControl.html` | Loads MSAL.js 4.13.1, WebChat 4.18.0, and the main JS file. Contains the root `<div>` for the control. |
| `COTXCopilotHostControl.js` | Orchestrates the entire browser-side flow: MSAL token acquisition, Copilot Studio SDK connection, WebChat rendering, context injection middleware, tool call card rendering, and D365 extensible control registration. |
| `COTXCopilotHostControl.css` | Styles the chat interface (bubble appearance, tables, lists, scrollbars, headings) to match a modern Copilot aesthetic. |

### Data Model

| Table | Purpose |
|-------|---------|
| `COTXCopilotAgentParameters` | Stores per-agent configuration: Entra ID credentials, Dataverse connection details, context and display preferences. Cross-company (shared). |
| `COTXCopilotAgentApplicationAreas` | Maps `COTXCopilotAgentApplicationArea` enum values to `COTXCopilotAgentParameters` records. Enables multi-agent routing by application area. |

## Control Lifecycle

### 1. Form Initialization

```mermaid
flowchart TD
    A["Form.init()"] --> B["FormRun creates COTXCopilotHostControl"]
    B --> C["new(): Registers all FormProperty bindings"]
    C --> D["applyBuild(): Reads design-time properties"]
    D --> E["initializeControl(applicationArea)"]
    E --> F["Reads COTXCopilotAgentParameters for the area"]
    E --> G["Sets connection properties (AppClientId, TenantId, etc.)"]
    E --> H["Sets user properties (UserId, UserName)"]
    E --> I["Subscribes to context changes"]
    I --> J["Local scope → COTXCopilotHostFormContext"]
    I --> K["Global scope → COTXCopilotHostGlobalContext"]
```

### 2. Browser-Side Rendering

```mermaid
flowchart TD
    A["JS init()"] --> B["tryRender() — retries up to 60 frames"]
    B --> C["acquireToken(appClientId, tenantId)"]
    B --> D["createCopilotConnection(token, envId, agentId)"]
    B --> E["WebChat.renderWebChat(directLine, store, element)"]
    C --> C1["Try silent — session cache → access token"]
    C --> C2["Fallback: popup → access token"]
    D --> D1["CopilotStudioWebChat.createConnection() → DirectLine"]
    E --> E1["Store middleware intercepts"]
    E1 --> E2["Outgoing messages: injects ERP context"]
    E1 --> E3["Incoming messages: captures agent responses"]
    E1 --> E4["Events: renders tool call Adaptive Cards"]
```

### 3. Context Flow

#### Global Context (Side Panel)

```mermaid
sequenceDiagram
    participant User
    participant Info
    participant GlobalCtx as COTXCopilotHostGlobalContext
    participant FormCtx as COTXCopilotHostFormContext
    participant Control as COTXCopilotHostControl
    participant JS as Browser JS

    User->>Info: Navigates to a new form
    Info->>GlobalCtx: onActivate fires
    GlobalCtx->>GlobalCtx: handleFormActivation(formRun)
    GlobalCtx->>FormCtx: Creates new COTXCopilotHostFormContext
    FormCtx->>FormCtx: Subscribes to root data source OnActivated
    FormCtx->>GlobalCtx: Fires onFormContextChange
    GlobalCtx->>Control: formContextChange()
    Control->>Control: Updates FormProperty bindings
    Control->>JS: JS reads updated properties
    JS->>JS: Next chat message includes new context
```

#### Local Context (Embedded Control)

```mermaid
sequenceDiagram
    participant Form
    participant Control as COTXCopilotHostControl
    participant FormCtx as COTXCopilotHostFormContext
    participant DS as Root DataSource

    Form->>Control: initializeControl() with Local scope
    Control->>FormCtx: Creates COTXCopilotHostFormContext for this form
    FormCtx->>DS: Subscribes to OnActivated
    DS->>FormCtx: User changes record
    FormCtx->>Control: formContextChange fires
    Note over Control: Same flow as global, but scoped to one form
```

### 4. Agent Response Handling

```mermaid
sequenceDiagram
    participant Agent as Copilot Studio Agent
    participant WebChat as WebChat Store
    participant JS as Browser JS
    participant Control as COTXCopilotHostControl (X++)
    participant Form as Form Event Handlers

    Agent->>WebChat: Sends reply
    WebChat->>JS: Captures incoming bot message
    alt waitingForBotReply (X++ initiated)
        JS->>Control: Calls RaiseAgentResponse command
        Control->>Control: Sets parmAgentResponse property
        Control->>Form: Fires onAgentResponse delegate
        Form->>Form: Event handlers react
    end
```

## Context Data Structure

The ERP context is injected into the `channelData.context` of every outgoing WebChat message:

```json
{
  "channelData": {
    "context": {
      "userLanguage": "en-us",
      "userTimeZone": "GMT Standard Time",
      "callingMethod": "",
      "legalEntity": "USMF",
      "currentUser": "Admin",
      "currentForm": "All Sales Orders",
      "currentMenuItem": "Sales order",
      "formMode": "",
      "currentRecord": {
        "tableName": "Sales order",
        "naturalKey": "Sales order",
        "naturalValue": "SO-000123"
      }
    }
  }
}
```

## Design Decisions

| Decision | Rationale |
|----------|-----------|
| **Browser-side MSAL** | No server-side secrets needed; leverages the user's existing Entra ID session. Popup fallback ensures first-time auth works. |
| **FormTemplateControl** | D365's extensible control pattern provides property binding, build-time designer support, and lifecycle hooks. |
| **Global singleton for context** | A single `COTXCopilotHostGlobalContext` instance subscribes once to `Info.onActivate`, avoiding redundant subscriptions. |
| **Application area routing** | Lookup table pattern allows multiple agents, with `Fallback` as a catch-all, extensible via enum extensions. |
| **Custom form pattern for side panel** | Aside pane forms require the `Custom` pattern and `setDisplayTarget(AsidePane)` before `super()` — this is per Microsoft guidance. |

## External Dependencies

| Library | Version | CDN | Purpose |
|---------|---------|-----|---------|
| MSAL.js | 4.13.1 | unpkg | Browser-side OAuth2 token acquisition |
| Bot Framework WebChat | 4.18.0 | unpkg | Chat UI rendering |
| Copilot Studio Client SDK | 1.2.3 | unpkg (ESM) | DirectLine connection to Copilot Studio agents |
