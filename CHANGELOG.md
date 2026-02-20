# Changelog

All notable changes to the D365 Copilot Toolbox are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [1.26.2.1] - 2026-02-20

### Added
- **Multiple conversation tabs** — users can open up to 8 parallel chat sessions inside a single control instance. Tabs support add, close, rename (double-click label), and switch. Each tab has its own Direct Line connection, WebChat store, and mutable state
- **Tab bar UI** — a persistent tab strip with an add-tab (`+`) button, per-tab close (`×`) button, and a restart (`↻`) button is rendered above the chat area
- **New chat / restart button** — the restart button (`↻`) in the tab bar tears down the active tab's WebChat session (React tree, Direct Line, subscriptions) and bootstraps a fresh one without reloading the form
- **MSAL v4 → v5 migration** — upgraded from `@azure/msal-browser` v4 to v5. The new version uses a redirect bridge (`COTXMsalRedirectBridge.html` / `COTXMsalRedirectBridge.js`) for COOP-compatible authentication in popups and iframes
- **Locally bundled vendor libraries** — MSAL Browser, MSAL Redirect Bridge, WebChat, and Copilot Studio Client are now shipped as D365 AxResource items (`COTXMsalBrowser_JS`, `COTXMsalRedirectBridge_JS`, `COTXMsalRedirectBridge_HTML`, `COTXWebChat_JS`, `COTXCopilotStudioClient_MJS`) instead of being loaded from external CDNs at runtime — eliminating supply-chain risk
- **Vendor library management** — new `Scripts/Update-VendorLibs.ps1` PowerShell script and `Scripts/vendor-libs.json` manifest for downloading and managing third-party JavaScript libraries from npm
- **Automated vendor update workflow** — GitHub Actions workflow (`.github/workflows/update-vendor-libs.yml`) runs weekly to check for npm updates and opens a pull request when newer versions are available
- **MSAL instance caching** — `PublicClientApplication` instances are cached per `clientId|tenantId` pair and reused across chat restarts and tab creation, avoiding duplicate MSAL instances and redundant `initialize()` calls
- **Chat header layout** — new `ensureChatLayout()` function creates a persistent chrome (tab bar + container area) that is separate from the per-tab WebChat React trees
- **Stale init detection** — a monotonically increasing `initId` counter per tab prevents race conditions when a restart is triggered while a previous initialisation is still in-flight
- **Keep connection alive parameter** — new `KeepConnectionAlive` toggle on Agent Parameters; when enabled, `dispose()` skips terminating the Direct Line connection so long-running agents survive form re-opens

### Changed
- **Browser-side rendering lifecycle** — `tryRender()` replaced by promise-based `waitForDependencies()` → `readControlParameters()` → `ensureChatLayout()` → `createTab()` flow
- **Per-tab state isolation** — `waitingForBotReply` and `toolCalls` are now scoped to each tab object instead of module-level variables
- **`PendingMessage` observer** — only the active tab processes X++ pending messages; inactive tabs ignore them
- **Control parameter reads** — switched from `$dyn.value` to `$dyn.peek` to avoid creating unintended reactive subscriptions

### Fixed
- **Multi-tenant MSAL account selection** — `acquireToken` now picks the cached MSAL account whose `tenantId` matches the agent's configured tenant instead of blindly taking `accounts[0]`. Prevents the wrong user identity from being used when both the home-tenant and a cross-tenant account are present in the session-storage cache
- **Enter key now sends messages** — the chat container intercepts the Enter keydown event (`stopPropagation`) so the D365 form framework no longer captures it; Shift+Enter still inserts a newline
- **Memory leaks on dispose / restart** — the React component tree is now explicitly unmounted via `ReactDOM.unmountComponentAtNode`, `$dyn.observe` subscriptions are properly disposed, and in-flight initialisations are invalidated
- **Proper dispose lifecycle** — `dispose()` now iterates all tabs, disposes subscriptions, unmounts React trees, and ends Direct Line connections (respecting `keepConnectionAlive`)
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
