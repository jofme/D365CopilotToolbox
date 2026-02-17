# Changelog

All notable changes to the D365 Copilot Toolbox are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [Unreleased]

### Added
- **New chat button** — a restart button ("↻ New chat") in the chat header bar lets users start a fresh conversation without reloading the form
- **MSAL instance caching** — `PublicClientApplication` instances are now cached per `clientId|tenantId` pair and reused across chat restarts, avoiding the memory overhead of redundant MSAL instances
- **Chat header bar** — new `ensureChatLayout()` function creates a persistent header bar and inner container, separating chrome from the WebChat React tree
- **Stale init detection** — a monotonically increasing `_initId` counter prevents race conditions when a restart is triggered while a previous initialisation is still in-flight

### Fixed
- **Enter key now sends messages** — the chat container intercepts the Enter keydown event (`stopPropagation`) so the D365 form framework no longer captures it; Shift+Enter still inserts a newline
- **Memory leaks on dispose / restart** — the React component tree is now explicitly unmounted via `ReactDOM.unmountComponentAtNode`, `$dyn.observe` subscriptions are properly disposed, and in-flight initialisations are invalidated
- **Observable side-effects in readControlParameters** — switched from `$dyn.value` to `$dyn.peek` to avoid creating unintended reactive subscriptions when reading control parameters

## [1.26.2.0] - 2026-02-15

Initial release of the D365 Copilot Toolbox, establishing the foundation for multi-agent workflows in D365 Finance & Operations. This release focuses on Microsoft Copilot Studio agent integration.

### Added
- **COTXCopilotHostControl** — extensible form control that embeds a Copilot Studio agent via the M365 Agent SDK
- **COTXCopilotHostSidePanel** — global aside-pane form for Copilot chat, accessible from the Settings menu
- **COTXCopilotAgentParameters** — configuration table and form for managing agent connections (Entra ID, Dataverse, context settings)
- **COTXCopilotAgentApplicationAreas** — application area mapping table enabling multi-agent routing
- **COTXCopilotHostFormContext** — local form context tracker (data source, record, navigation fields)
- **COTXCopilotHostGlobalContext** — global singleton that tracks form navigation across D365 and propagates context to the side panel
- **Browser-side MSAL.js authentication** — delegated token acquisition with silent + popup fallback
- **Copilot Studio Agent SDK integration** — DirectLine connection via `@microsoft/agents-copilotstudio-client`
- **ERP context injection** — automatic injection of legal entity, form, record, and user info into agent messages
- **Tool call visualization** — optional Adaptive Card display of agent tool execution and reasoning
- **Agent response delegate** — `onAgentResponse` event for X++ form code to react to agent replies
- **Programmatic messaging** — `sendMessage()` API for X++ code to send messages to the agent
- **Extensible application areas** — `COTXCopilotAgentApplicationArea` enum with `IsExtensible = true`
- **Context scope types** — Global, Local, and None scope options for embedded controls
- **Security model** — Admin and User roles, duties, and privileges
- **Settings menu integration** — Copilot Agent item in the Settings gear menu
- **System Administration integration** — Agent Parameters form under Setup > Copilot Toolbox
- **Example: SalesTable integration** — `CopilotToolboxExamples` model demonstrating embedded agent on the Sales Order form
- **Project documentation** — Getting Started, Architecture, Configuration, Extending, Security, and Contributing guides
- **CI/CD** — GitHub Actions build workflow using FSC-PS
