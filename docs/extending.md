# Extending the Copilot Toolbox

This guide explains how to add Copilot agents to your own custom forms, define new application areas, and handle agent responses in X++.

## Adding a Copilot Agent to a Form

There are two approaches: **embedded (local context)** and **global (side panel)**. Most customizations use the embedded approach.

### Embedded Agent on a Form (Local Context)

This places the Copilot chat directly on a form tab. The agent receives context specific to that form's active record.

### Complete Example: SalesTable

The `CopilotToolboxExamples` model demonstrates this pattern:

1. **Enum extension** (`COTXCopilotAgentApplicationArea.CopilotToolboxExamples`):
   - Adds `CTXESalesTable` value with label "Sales Table"

2. **Form extension** (`SalesTable.CopilotToolboxExamples`):
   - Adds a `Copilot Agent` tab page to the `LineViewTab`
   - Places `COTXCopilotHostControl` with `parmCopilotContext = CTXESalesTable` and `parmContextScopeType = Local`

3. **Runtime:** When the SalesTable form opens, the control reads the agent configuration mapped to `CTXESalesTable`, connects to Copilot Studio, and tracks the selected sales order as context.

## Context Scope Types

| Scope | Enum Value | Behavior |
|-------|------------|----------|
| **Global** | `COTXCopilotAgentContextScopeType::Global` | Subscribes to `COTXCopilotHostGlobalContext`; receives context updates as the user navigates between forms. Used by the side panel. |
| **Local** | `COTXCopilotAgentContextScopeType::Local` | Creates a `COTXCopilotHostFormContext` scoped to the hosting form. Tracks only that form's root data source. |
| **None** | `COTXCopilotAgentContextScopeType::None` | No context is tracked or sent. The agent receives messages without ERP context. |

## Sending Messages from X++

You can programmatically send messages to the Copilot agent from X++ code:

```xpp
// Get a reference to the control (e.g., from form control declaration)
COTXCopilotHostControl copilotControl = this.control(
    this.controlId(formControlStr(YourForm, CopilotAgentHost)));

// Send a message
copilotControl.sendMessage("What is the status of this order?");
```

The `sendMessage` method sets the `PendingMessage` form property, which the JavaScript layer picks up and dispatches through WebChat.

## Handling Agent Responses in X++

The `COTXCopilotHostControl` raises an `onAgentResponse` delegate when the agent replies to a programmatically sent message.

### Subscribe to the Delegate

```xpp
// In your form's init or after creating the control reference:
copilotControl.onAgentResponse += eventhandler(this.handleAgentResponse);
```

### Implement the Handler

```xpp
private void handleAgentResponse(COTXCopilotHostControl _sender, str _responseText)
{
    // _responseText contains the agent's reply
    info(strFmt("Agent says: %1", _responseText));

    // Parse the response, update form fields, trigger actions, etc.
}
```

> **Note:** The `onAgentResponse` delegate only fires for responses to messages initiated via `sendMessage()`. It does not fire for user-typed messages in the chat UI.

## Accessing the Latest Response

You can also poll the latest agent response without using the delegate:

```xpp
str lastResponse = copilotControl.parmAgentResponse();
```

## Design-Time Properties

When adding the `COTXCopilotHostControl` in the Visual Studio form designer, two properties are available under the **COTX** category:

| Property | Type | Description |
|----------|------|-------------|
| **Application Area** | `COTXCopilotAgentApplicationArea` | Selects which agent configuration to use |
| **Copilot Context Scope** | `COTXCopilotAgentContextScopeType` | Controls how ERP context is tracked |

## Customizing the Chat Appearance

The chat UI is styled via `COTXCopilotHostControl.css`. The style options in `COTXCopilotHostControl.js` control WebChat's built-in theming:

| Option | Default | Description |
|--------|---------|-------------|
| `accent` | `#7B68EE` | Primary accent color |
| `bubbleBackground` | `#F5F5F5` | Bot message bubble background |
| `bubbleFromUserBackground` | `#E8E0FF` | User message bubble background |
| `sendBoxBackground` | `#F5F5F5` | Input box background |
| `sendBoxBorderRadius` | `24` | Input box corner rounding |
| `bubbleMaxWidth` | `980` | Maximum message width in pixels |