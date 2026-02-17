(function () {
    'use strict';

    // -----------------------------------------------------------------------
    // Constants
    // -----------------------------------------------------------------------

    /** @type {string} CDN URL for the Copilot Studio Agent SDK browser bundle. */
    var COPILOT_SDK_URL = 'https://unpkg.com/@microsoft/agents-copilotstudio-client@1.2.3/dist/src/browser.mjs';

    /** @type {number} Maximum number of animation-frame ticks to wait for dependencies. */
    var MAX_RENDER_ATTEMPTS = 60;

    /** @type {string} Log prefix */
    var LOG_PREFIX = 'COTXCopilotHostControl';

    /** Direct Line / WebChat action type constants. */
    var ACTION_TYPES = {
        POST_ACTIVITY: 'DIRECT_LINE/POST_ACTIVITY',
        INCOMING_ACTIVITY: 'DIRECT_LINE/INCOMING_ACTIVITY',
        SEND_MESSAGE: 'WEB_CHAT/SEND_MESSAGE'
    };

    /** Copilot Studio Dynamic Plan event names. */
    var EVENT_NAMES = {
        PLAN_RECEIVED: 'DynamicPlanReceived',
        PLAN_STEP_TRIGGERED: 'DynamicPlanStepTriggered'
    };

    /** Activity type constants. */
    var ACTIVITY_TYPES = {
        MESSAGE: 'message',
        TYPING: 'typing',
        EVENT: 'event'
    };

    /** Entity type constants found in activity.entities[]. */
    var ENTITY_TYPES = {
        THOUGHT: 'thought'
    };

    /** Thought status constants. */
    var THOUGHT_STATUS = {
        COMPLETE: 'complete'
    };

    /** WebChat visual style overrides. */
    var STYLE_OPTIONS = {
        accent: '#7B68EE',
        autoScrollSnapOnPage: true,
        autoScrollSnapOnPageOffset: 0,
        backgroundColor: '#FFFFFF',
        bubbleBackground: '#F5F5F5',
        bubbleBorderColor: '#E0E0E0',
        bubbleBorderRadius: 12,
        bubbleBorderStyle: 'solid',
        bubbleBorderWidth: 0,
        bubbleTextColor: '#1A1A1A',
        bubbleNubSize: 0,
        bubbleNubOffset: 0,
        bubbleFromUserBackground: '#E8E0FF',
        bubbleFromUserBorderColor: '#D4C8FF',
        bubbleFromUserBorderRadius: 12,
        bubbleFromUserBorderStyle: 'solid',
        bubbleFromUserBorderWidth: 0,
        bubbleFromUserTextColor: '#1A1A1A',
        bubbleFromUserNubSize: 0,
        bubbleFromUserNubOffset: 0,
        bubbleMaxWidth: 980,
        bubbleMinWidth: 120,
        bubbleMinHeight: 40,
        sendBoxBackground: '#F5F5F5',
        sendBoxBorderTop: 'none',
        sendBoxBorderRadius: 24,
        sendBoxButtonColor: '#7B68EE',
        sendBoxButtonColorOnDisabled: '#CCCCCC',
        sendBoxButtonColorOnFocus: '#6A5ACD',
        sendBoxButtonColorOnHover: '#6A5ACD',
        sendBoxHeight: 50,
        sendBoxPlaceholderColor: '#767676',
        sendBoxTextColor: '#1A1A1A',
        hideUploadButton: false,
        suggestedActionBackgroundColor: '#FFFFFF',
        suggestedActionBackgroundColorOnHover: '#F0EBFF',
        suggestedActionBorderColor: '#7B68EE',
        suggestedActionBorderRadius: 20,
        suggestedActionBorderWidth: 1,
        suggestedActionTextColor: '#7B68EE',
        suggestedActionTextColorOnHover: '#6A5ACD',
        suggestedActionLayout: 'flow',
        suggestedActionHeight: 36,
        paddingRegular: 12,
        paddingWide: 16,
        groupTimestamp: false,
        subtleColor: '#767676',
        timestampColor: '#767676',
        typingAnimationHeight: 20,
        typingAnimationWidth: 64,
        rootHeight: '100%',
        rootWidth: '100%',
        hideScrollToEndButton: false,
        markdownRenderHTML: true,
        emojiSet: true
    };

    // -----------------------------------------------------------------------
    // Module-scoped mutable state
    // -----------------------------------------------------------------------

    /**
     * Shared runtime state for the control instance.
     *
     * @property {boolean}  waitingForBotReply - True while we expect the next
     *           bot message to be the response to a programmatic X++ send.
     * @property {Object.<string, {tools: Array, cardSent: boolean}>} toolCalls
     *           - Tracks Dynamic Plan tool definitions keyed by plan identifier.
     */
    var _state = {
        waitingForBotReply: false,
        toolCalls: {}
    };

    /**
     * Module-scoped cache for MSAL PublicClientApplication instances, keyed by
     * "clientId|tenantId". Prevents creating duplicate MSAL instances across
     * restarts or multiple control instances on the same page.
     *
     * @type {Object.<string, {instance: Object, initPromise: Promise}>}
     */
    var _msalCache = {};

    // -----------------------------------------------------------------------
    // Token acquisition via MSAL.js
    // -----------------------------------------------------------------------

    /**
     * Acquires a delegated user token for the Power Platform API using MSAL.js.
     *
     * Tries silent acquisition first (leveraging the D365 Entra ID session),
     * then falls back to a popup if interaction is required.
     *
     * The underlying MSAL PublicClientApplication is created once per
     * clientId + tenantId pair and reused across subsequent calls, avoiding the
     * memory overhead of redundant instances.
     *
     * @param {string} appClientId - App registration client ID (SPA / public client).
     * @param {string} tenantId    - Entra ID tenant ID.
     * @returns {Promise<string>}    The access token string.
     */
    function acquireToken(appClientId, tenantId) {
        var cacheKey = appClientId + '|' + tenantId;

        if (!_msalCache[cacheKey]) {
            var msalConfig = {
                auth: {
                    clientId: appClientId,
                    authority: 'https://login.microsoftonline.com/' + tenantId
                },
                cache: {
                    cacheLocation: 'sessionStorage'
                }
            };

            var instance = new msal.PublicClientApplication(msalConfig);

            _msalCache[cacheKey] = {
                instance: instance,
                initPromise: instance.initialize()
            };
        }

        var cached = _msalCache[cacheKey];
        var msalInstance = cached.instance;

        var loginRequest = {
            scopes: ['https://api.powerplatform.com/.default'],
            redirectUri: window.location.origin
        };

        return cached.initPromise.then(function () {
            var accounts = msalInstance.getAllAccounts();

            if (accounts.length > 0) {
                return msalInstance.acquireTokenSilent({
                    scopes: loginRequest.scopes,
                    account: accounts[0],
                    redirectUri: loginRequest.redirectUri
                }).then(function (response) {
                    return response.accessToken;
                }).catch(function (silentError) {
                    console.warn(LOG_PREFIX + '.acquireToken: Silent token failed, trying popup', silentError);
                    return msalInstance.acquireTokenPopup(loginRequest).then(function (response) {
                        return response.accessToken;
                    });
                });
            }

            return msalInstance.acquireTokenPopup(loginRequest).then(function (response) {
                return response.accessToken;
            });
        });
    }

    // -----------------------------------------------------------------------
    // Store middleware — individual handlers
    // -----------------------------------------------------------------------

    /**
     * Determines whether the given action is a typing indicator that should be
     * suppressed (dropped before reaching the WebChat renderer).
     *
     * @param {Object} action - Redux action dispatched through the WebChat store.
     * @returns {boolean} True if the action should be discarded.
     */
    function isTypingIndicator(action) {
        if (action.type !== ACTION_TYPES.INCOMING_ACTIVITY) {
            return false;
        }

        var activity = action.payload && action.payload.activity;

        return !!(activity && activity.type === ACTIVITY_TYPES.TYPING);
    }

    /**
     * If the action is an outgoing user message and context sending is enabled,
     * returns a shallow clone of the action with ERP context attached to
     * `channelData.context`. Otherwise returns the original action unchanged.
     *
     * @param {Object} action - Redux action.
     * @param {Object} data   - D365 control data bindings.
     * @returns {Object} The (possibly enriched) action.
     */
    function enrichOutgoingActivity(action, data) {
        if (action.type !== ACTION_TYPES.POST_ACTIVITY) {
            return action;
        }

        var activity = action.payload.activity;

        if (activity.type !== ACTIVITY_TYPES.MESSAGE) {
            return action;
        }

        if (!$dyn.peek(data.SendContext)) {
            return action;
        }

        var erpContext = {
            userLanguage: $dyn.peek(data.UserLanguage),
            userTimeZone: $dyn.peek(data.UserTimeZone),
            callingMethod: $dyn.peek(data.CallingMethod),
            legalEntity: $dyn.peek(data.LegalEntity),
            currentUser: $dyn.peek(data.UserId),
            currentForm: $dyn.peek(data.CurrentFormName),
            currentMenuItem: $dyn.peek(data.CurrentMenuItem),
            formMode: $dyn.peek(data.FormMode),
            currentRecord: {
                tableName: $dyn.peek(data.TableName),
                naturalKey: $dyn.peek(data.NaturalKey),
                naturalValue: $dyn.peek(data.NaturalValue)
            }
        };

        return Object.assign({}, action, {
            payload: Object.assign({}, action.payload, {
                activity: Object.assign({}, activity, {
                    channelData: Object.assign({}, activity.channelData, {
                        context: erpContext
                    })
                })
            })
        });
    }

    /**
     * Inspects an incoming bot message. If we are waiting for a programmatic
     * reply (X++ sent a message via PendingMessage), captures the first
     * non-tool-thought bot text and forwards it to X++ through
     * `data.RaiseAgentResponse`.
     *
     * @param {Object} action - Redux action.
     * @param {Object} data   - D365 control data bindings.
     * @returns {void}
     */
    function captureAgentResponse(action, data) {
        if (action.type !== ACTION_TYPES.INCOMING_ACTIVITY) {
            return;
        }

        var activity = action.payload && action.payload.activity;

        if (!activity || activity.type !== ACTIVITY_TYPES.MESSAGE) {
            return;
        }

        if (activity.from.role !== 'bot' || !activity.text) {
            return;
        }

        if (!_state.waitingForBotReply) {
            return;
        }

        var isToolThought = activity.channelData && activity.channelData.isToolThought;

        if (isToolThought) {
            return;
        }

        _state.waitingForBotReply = false;
        $dyn.callFunction(data.RaiseAgentResponse, null, [activity.text]);
    }

    // -----------------------------------------------------------------------
    // Dynamic Plan / tool-call card building
    // -----------------------------------------------------------------------

    /**
     * Builds an Adaptive Card JSON object that visualises one or more tool
     * calls, each with an expandable "Show Thoughts" section.
     *
     * @param {Array<{llmIdentifierPrefix: string, description: string, thought: string}>} tools
     *        - Array of tool metadata objects.
     * @returns {Object} Adaptive Card payload.
     */
    function buildToolCallCard(tools) {
        var card = {
            "$schema": "https://adaptivecards.io/schemas/adaptive-card.json",
            "type": "AdaptiveCard",
            "version": "1.6",
            "body": [
                {
                    "type": "TextBlock",
                    "text": "🔧 Tool Call",
                    "weight": "Bolder",
                    "size": "Small"
                }
            ]
        };

        tools.forEach(function (tool) {
            card.body.push({
                "type": "Container",
                "spacing": "Small",
                "separator": true,
                "items": [
                    {
                        "type": "ColumnSet",
                        "columns": [
                            {
                                "type": "Column",
                                "width": "auto",
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": "⚡",
                                        "size": "Small"
                                    }
                                ]
                            },
                            {
                                "type": "Column",
                                "width": "stretch",
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": tool.llmIdentifierPrefix,
                                        "weight": "Bolder",
                                        "wrap": true,
                                        "size": "Small"
                                    },
                                    {
                                        "type": "TextBlock",
                                        "text": tool.description,
                                        "isSubtle": true,
                                        "spacing": "None",
                                        "wrap": true,
                                        "size": "Small"
                                    }
                                ]
                            }
                        ]
                    },
                    {
                        "type": "ActionSet",
                        "spacing": "Small",
                        "actions": [
                            {
                                "type": "Action.ShowCard",
                                "title": "💭 Show Thoughts",
                                "card": {
                                    "type": "AdaptiveCard",
                                    "body": [
                                        {
                                            "type": "TextBlock",
                                            "text": tool.thought,
                                            "wrap": true,
                                            "size": "Small",
                                            "isSubtle": true
                                        }
                                    ]
                                }
                            }
                        ]
                    }
                ]
            });
        });

        return card;
    }

    /**
     * Builds a plain-text summary of tool calls for use as a chat message
     * instead of an Adaptive Card.
     *
     * @param {Array<{llmIdentifierPrefix: string, description: string}>} tools
     *        - Array of tool metadata objects.
     * @returns {string} Markdown-formatted tool call summary.
     */
    function buildToolCallText(tools) {
        var lines = ['🔧 **Tool Call**'];

        tools.forEach(function (tool) {
            lines.push('⚡ **' + tool.llmIdentifierPrefix + '** — ' + tool.description);
        });

        return lines.join('\n\n');
    }

    /**
     * Stores tool definitions when a `DynamicPlanReceived` event arrives.
     *
     * @param {Object} planValue - The `activity.value` from the plan event.
     * @returns {void}
     */
    function handlePlanReceived(planValue) {
        var planId = planValue.planIdentifier;

        _state.toolCalls[planId] = {
            tools: (planValue.toolDefinitions || []).map(function (tool, index) {
                return {
                    id: index.toString(),
                    llmIdentifierPrefix: tool.llmIdentifierPrefix || tool.displayName,
                    description: tool.description || '',
                    identifier: tool.identifier,
                    thought: 'Waiting for agent reasoning...'
                };
            }),
            cardSent: false
        };

        console.log(LOG_PREFIX + '.handlePlanReceived: Plan registered', planId);
    }

    /**
     * Updates the matching tool's thought text when a
     * `DynamicPlanStepTriggered` event arrives. If this is the first step for
     * the plan and tool-call display is enabled, injects an Adaptive Card into
     * the WebChat conversation.
     *
     * @param {Object}   stepValue    - The `activity.value` from the step event.
     * @param {Object}   data         - D365 control data bindings.
     * @param {Function} dispatchNext - The store's `next` function for injecting activities.
     * @returns {void}
     */
    function handlePlanStepTriggered(stepValue, data, dispatchNext) {
        var planId = stepValue.planIdentifier;
        var thought = stepValue.thought || '';
        var planData = _state.toolCalls[planId];

        if (!planData) {
            return;
        }

        // Update the matching tool with its thought
        planData.tools.forEach(function (tool) {
            if (stepValue.taskDialogId && stepValue.taskDialogId.indexOf(tool.identifier) !== -1) {
                tool.thought = thought;
            }
        });

        var showToolCalls = !!$dyn.peek(data.ShowToolCalls);
        var showThoughts = !!$dyn.peek(data.ShowThoughts);

        if (planData.cardSent || planData.tools.length === 0 || !showToolCalls) {
            console.log(LOG_PREFIX + '.handlePlanStepTriggered:', thought);
            return;
        }

        planData.cardSent = true;

        var cardActivity;

        if (showThoughts) {
            // Plain text message — avoids Adaptive Card scroll issues
            cardActivity = {
                type: ACTIVITY_TYPES.MESSAGE,
                from: { role: 'bot' },
                text: buildToolCallText(planData.tools),
                channelData: { isToolThought: true }
            };
        } else {
            // Adaptive Card with expandable thought sections
            cardActivity = {
                type: ACTIVITY_TYPES.MESSAGE,
                from: { role: 'bot' },
                attachments: [{
                    contentType: 'application/vnd.microsoft.card.adaptive',
                    content: buildToolCallCard(planData.tools)
                }],
                channelData: { isToolThought: true }
            };
        }

        dispatchNext({
            type: ACTION_TYPES.INCOMING_ACTIVITY,
            payload: { activity: cardActivity }
        });

        console.log(LOG_PREFIX + '.handlePlanStepTriggered:', thought);
    }

    /**
     * Processes Dynamic Plan event activities (`DynamicPlanReceived` and
     * `DynamicPlanStepTriggered`). These events visualise which tools/actions
     * the Copilot Studio orchestrator is invoking.
     *
     * @param {Object}   action       - Redux action.
     * @param {Object}   data         - D365 control data bindings.
     * @param {Function} dispatchNext - The store's `next` function.
     * @returns {boolean} True if the action was a plan event and was consumed
     *          (caller should **not** forward it to `next`).
     */
    function handleDynamicPlanEvents(action, data, dispatchNext) {
        if (action.type !== ACTION_TYPES.INCOMING_ACTIVITY) {
            return false;
        }

        var activity = action.payload && action.payload.activity;

        if (!activity || activity.type !== ACTIVITY_TYPES.EVENT) {
            return false;
        }

        if (activity.name === EVENT_NAMES.PLAN_RECEIVED && activity.value) {
            handlePlanReceived(activity.value);
        }

        if (activity.name === EVENT_NAMES.PLAN_STEP_TRIGGERED && activity.value) {
            handlePlanStepTriggered(activity.value, data, dispatchNext);
        }

        // All event-type activities are consumed — don't pass to WebChat
        return true;
    }

    // -----------------------------------------------------------------------
    // Thought / chain-of-thought bubble injection
    // -----------------------------------------------------------------------

    /**
     * Extracts a completed "thought" entity from a typing activity and injects
     * it as a visible bot message in the WebChat conversation.
     *
     * The thought is rendered as an italicised, subtle chat bubble so the user
     * can see the agent's reasoning without confusing it with actual replies.
     *
     * @param {Object}   action       - Redux action.
     * @param {Object}   data         - D365 control data bindings.
     * @param {Function} dispatchNext - The store's `next` function for injecting activities.
     * @returns {boolean} True if a thought was injected (caller may still
     *          forward or suppress the original action as desired).
     */
    function handleCompletedThoughts(action, data, dispatchNext) {
        if (action.type !== ACTION_TYPES.INCOMING_ACTIVITY) {
            return false;
        }

        var activity = action.payload && action.payload.activity;

        if (!activity || activity.type !== ACTIVITY_TYPES.TYPING) {
            return false;
        }

        if (!$dyn.peek(data.ShowThoughts)) {
            return false;
        }

        var entities = activity.entities;

        if (!entities || !entities.length) {
            return false;
        }

        var injected = false;

        entities.forEach(function (entity) {
            if (entity.type !== ENTITY_TYPES.THOUGHT) {
                return;
            }

            if (entity.status !== THOUGHT_STATUS.COMPLETE) {
                return;
            }

            if (!entity.text) {
                return;
            }

            var thoughtText = '💭 ' + entity.text;

            var thoughtActivity = {
                type: ACTIVITY_TYPES.MESSAGE,
                from: { role: 'bot' },
                text: thoughtText,
                channelData: { isToolThought: true }
            };

            dispatchNext({
                type: ACTION_TYPES.INCOMING_ACTIVITY,
                payload: { activity: thoughtActivity }
            });

            console.log(LOG_PREFIX + '.handleCompletedThoughts: Injected thought —', entity.title);
            injected = true;
        });

        return injected;
    }

    // -----------------------------------------------------------------------
    // WebChat store creation
    // -----------------------------------------------------------------------

    /**
     * Creates a WebChat Redux store with middleware that:
     *  1. Suppresses typing indicators.
     *  2. Enriches outgoing messages with ERP context.
     *  3. Captures bot replies destined for X++.
     *  4. Handles Dynamic Plan tool-call events.
     *
     * @param {Object} data - D365 control data bindings.
     * @returns {Object} A WebChat Redux store instance.
     */
    function createStore(data) {
        return window.WebChat.createStore({}, function () {
            return function (next) {
                return function (action) {
                    // Inject completed thoughts before suppressing the typing indicator
                    handleCompletedThoughts(action, data, next);

                    if (isTypingIndicator(action)) {
                        return;
                    }

                    action = enrichOutgoingActivity(action, data);

                    captureAgentResponse(action, data);

                    if (handleDynamicPlanEvents(action, data, next)) {
                        return;
                    }

                    return next(action);
                };
            };
        });
    }

    // -----------------------------------------------------------------------
    // Copilot Studio connection via Agent SDK
    // -----------------------------------------------------------------------

    /**
     * Establishes a Direct Line connection to a Copilot Studio agent using the
     * Agent SDK.
     *
     * @param {string} token         - Power Platform API access token.
     * @param {string} environmentId - Power Platform environment ID.
     * @param {string} agentId       - Copilot Studio agent identifier.
     * @returns {Promise<Object>} A Direct Line connection object compatible
     *          with WebChat's `directLine` prop.
     */
    function createCopilotConnection(token, environmentId, agentId) {
        return import(COPILOT_SDK_URL).then(function (sdk) {
            var settings = {
                environmentId: environmentId,
                agentIdentifier: agentId
            };

            var client = new sdk.CopilotStudioClient(settings, token);

            return sdk.CopilotStudioWebChat.createConnection(client, {
                showTyping: false
            });
        });
    }

    // -----------------------------------------------------------------------
    // Subscription & layout helpers
    // -----------------------------------------------------------------------

    /**
     * Safely disposes a `$dyn.observe` subscription if it exists.
     *
     * @param {Object|null} subscription - The subscription returned by `$dyn.observe`.
     * @returns {null} Always returns null for assignment convenience.
     */
    function disposeSubscription(subscription) {
        if (subscription && typeof subscription.dispose === 'function') {
            subscription.dispose();
        }

        return null;
    }

    /**
     * Creates the chat header bar (with a restart button) and an inner
     * container for WebChat, inserting both into the control's root element.
     *
     * Called once during `init`; subsequent calls return the existing inner
     * container without modification.
     *
     * @param {HTMLElement} element - The control's root DOM element.
     * @param {Object}      self   - The control instance.
     * @returns {HTMLElement} The inner container element where WebChat should render.
     */
    function ensureChatLayout(element, self) {
        if (self._chatContainer) {
            return self._chatContainer;
        }

        // Header bar
        var header = document.createElement('div');
        header.className = 'cotx-chat-header';

        var restartBtn = document.createElement('button');
        restartBtn.className = 'cotx-chat-restart-btn';
        restartBtn.type = 'button';
        restartBtn.title = 'Start a new conversation';
        restartBtn.textContent = '\u21BB New chat';

        header.appendChild(restartBtn);

        // Inner container for WebChat
        var container = document.createElement('div');
        container.className = 'cotx-chat-container';

        element.appendChild(header);
        element.appendChild(container);

        self._chatContainer = container;
        self._restartButton = restartBtn;

        return container;
    }

    // -----------------------------------------------------------------------
    // Chat restart
    // -----------------------------------------------------------------------

    /**
     * Tears down the current WebChat session and bootstraps a fresh one.
     *
     * Cleanup steps:
     *  1. Disposes the `$dyn.observe` subscription on PendingMessage.
     *  2. Ends the Direct Line connection (frees server-side resources).
     *  3. Unmounts the React component tree via `ReactDOM.unmountComponentAtNode`
     *     so that all internal subscriptions, timers, and closures held by
     *     WebChat and its Redux store are properly released.
     *  4. Resets module-scoped mutable state.
     *
     * A new token, connection, store, and React tree are then created via
     * {@link initializeWebChat}.
     *
     * @param {HTMLElement} chatContainer - The inner element hosting WebChat.
     * @param {Object}      data          - D365 control data bindings.
     * @param {Object}      self          - The control instance.
     * @returns {Promise<void>}
     */
    function restartChat(chatContainer, data, self) {
        // Double-click protection
        if (self._restartButton) {
            self._restartButton.disabled = true;
        }

        // 1. Dispose $dyn.observe subscription
        self._pendingMessageSubscription = disposeSubscription(self._pendingMessageSubscription);

        // 2. End the Direct Line connection
        if (self._directLine && typeof self._directLine.end === 'function') {
            self._directLine.end();
            self._directLine = null;
        }

        // 3. Properly unmount the React component tree so WebChat releases
        //    all internal subscriptions, timers, and Redux store closures.
        //    ReactDOM is exposed globally by the WebChat CDN full bundle.
        if (window.ReactDOM && window.ReactDOM.unmountComponentAtNode) {
            window.ReactDOM.unmountComponentAtNode(chatContainer);
        }

        // 4. Reset module-scoped state
        _state.waitingForBotReply = false;
        _state.toolCalls = {};

        // 5. Re-initialise
        var params = readControlParameters(data);

        if (!params) {
            if (self._restartButton) {
                self._restartButton.disabled = false;
            }

            return Promise.reject(new Error(LOG_PREFIX + '.restartChat: Missing control parameters'));
        }

        return initializeWebChat(chatContainer, data, params, self)
            .then(function () {
                if (self._restartButton) {
                    self._restartButton.disabled = false;
                }

                console.log(LOG_PREFIX + '.restartChat: Chat restarted successfully');
            })
            .catch(function (error) {
                if (self._restartButton) {
                    self._restartButton.disabled = false;
                }

                console.error(LOG_PREFIX + '.restartChat: Error restarting chat', error);
            });
    }

    // -----------------------------------------------------------------------
    // Initialisation helpers
    // -----------------------------------------------------------------------

    /**
     * Returns a Promise that resolves once the required browser globals
     * (`WebChat`, `msal`) are available, polling via `requestAnimationFrame`.
     *
     * Rejects if the globals have not appeared after {@link MAX_RENDER_ATTEMPTS}
     * animation frames.
     *
     * @param {HTMLElement} element - The DOM element that will host WebChat.
     * @returns {Promise<void>}
     */
    function waitForDependencies(element) {
        return new Promise(function (resolve, reject) {
            function check(attempt) {
                var ready = element
                    && window.WebChat
                    && window.WebChat.renderWebChat
                    && window.WebChat.createStore
                    && window.msal;

                if (ready) {
                    resolve();
                    return;
                }

                if (attempt >= MAX_RENDER_ATTEMPTS) {
                    reject(new Error(LOG_PREFIX + '.waitForDependencies: Dependencies did not load within '
                        + MAX_RENDER_ATTEMPTS + ' frames.'));
                    return;
                }

                requestAnimationFrame(function () { check(attempt + 1); });
            }

            check(0);
        });
    }

    /**
     * Reads the control parameters from D365 data bindings and validates that
     * all required values are present.
     *
     * @param {Object} data - D365 control data bindings.
     * @returns {{ appClientId: string, tenantId: string, environmentId: string, agentId: string } | null}
     *          The parameter set, or `null` if any required value is missing.
     */
    function readControlParameters(data) {
        var params = {
            appClientId: $dyn.peek(data.AppClientId),
            tenantId: $dyn.peek(data.TenantId),
            environmentId: $dyn.peek(data.EnvironmentId),
            agentId: $dyn.peek(data.AgentIdentifier)
        };

        if (!params.appClientId || !params.tenantId || !params.environmentId || !params.agentId) {
            console.warn(LOG_PREFIX + '.readControlParameters: One or more required parameters are missing.', params);
            return null;
        }

        return params;
    }

    /**
     * Wires up the `PendingMessage` observer so that when X++ sets a value,
     * the message is dispatched through WebChat and the observable is cleared.
     *
     * @param {Object} data  - D365 control data bindings.
     * @param {Object} store - WebChat Redux store.
     * @returns {Object} The `$dyn.observe` subscription (call `.dispose()` to
     *          unsubscribe).
     */
    function observePendingMessages(data, store) {
        return $dyn.observe(data.PendingMessage, function (message) {
            if (!message) {
                return;
            }

            _state.waitingForBotReply = true;

            store.dispatch({
                type: ACTION_TYPES.SEND_MESSAGE,
                payload: { text: message }
            });

            // Clear the observable so X++ can set the next message
            $dyn.callFunction(
                $dyn.observable(data.PendingMessage),
                null,
                ['']
            );
        });
    }

    /**
     * Orchestrates the full initialisation sequence: acquires a token, creates
     * the Copilot Studio Direct Line connection, renders WebChat into the host
     * element, and starts observing X++ pending messages.
     *
     * Uses an incrementing `_initId` on the control instance to detect stale
     * initialisations — if a restart is triggered while a previous init is
     * in-flight, the earlier init's `.then` callbacks silently bail out.
     *
     * @param {HTMLElement} element - The DOM element that will host WebChat.
     * @param {Object}      data   - D365 control data bindings.
     * @param {{ appClientId: string, tenantId: string, environmentId: string, agentId: string }} params
     *        - Validated control parameters.
     * @param {Object}      self   - The control instance (for storing the Direct Line reference).
     * @returns {Promise<void>}
     */
    function initializeWebChat(element, data, params, self) {
        var myInitId = ++self._initId;

        return acquireToken(params.appClientId, params.tenantId)
            .then(function (token) {
                if (self._initId !== myInitId) { return; }

                return createCopilotConnection(token, params.environmentId, params.agentId);
            })
            .then(function (directLine) {
                if (!directLine || self._initId !== myInitId) { return; }

                self._directLine = directLine;
                self._keepConnectionAlive = !!$dyn.peek(data.KeepConnectionAlive);

                var store = createStore(data);

                window.WebChat.renderWebChat({
                    directLine: directLine,
                    styleOptions: STYLE_OPTIONS,
                    store: store
                }, element);

                self._pendingMessageSubscription = observePendingMessages(data, store);
            });
    }

    // -----------------------------------------------------------------------
    // D365 F&O Extensible Control Registration
    // -----------------------------------------------------------------------

    /**
     * COTXCopilotHostControl — embeds a Copilot Studio WebChat instance inside
     * a D365 Finance & Operations form control.
     *
     * @constructor
     * @param {Object}      data    - D365 control data bindings.
     * @param {HTMLElement}  element - The DOM element assigned to this control.
     */
    $dyn.controls.COTXCopilotHostControl = function (data, element) {
        $dyn.ui.Control.apply(this, arguments);
        $dyn.ui.applyDefaults(this, data, $dyn.ui.defaults.COTXCopilotHostControl);

        /** @type {Object|null} Direct Line connection reference (for cleanup). */
        this._directLine = null;

        /** @type {boolean} When true, the Direct Line connection is kept alive on dispose. */
        this._keepConnectionAlive = false;

        /** @type {Object|null} Subscription returned by `$dyn.observe` on PendingMessage. */
        this._pendingMessageSubscription = null;

        /** @type {HTMLElement|null} Inner container element where WebChat renders. */
        this._chatContainer = null;

        /** @type {HTMLElement|null} The restart button element. */
        this._restartButton = null;

        /** @type {number} Monotonically increasing initialisation counter for stale-init detection. */
        this._initId = 0;
    };

    $dyn.controls.COTXCopilotHostControl.prototype = $dyn.ui.extendPrototype($dyn.ui.Control.prototype, {

        /**
         * Lifecycle hook called by the D365 control framework after the
         * control's DOM element is available. Waits for external scripts to
         * load, validates parameters, and bootstraps the WebChat session.
         *
         * @param {Object}      data    - D365 control data bindings.
         * @param {HTMLElement}  element - The DOM element assigned to this control.
         * @returns {void}
         */
        init: function (data, element) {
            $dyn.ui.Control.prototype.init.apply(this, arguments);

            var self = this;

            waitForDependencies(element)
                .then(function () {
                    var params = readControlParameters(data);

                    if (!params) {
                        return;
                    }

                    var chatContainer = ensureChatLayout(element, self);

                    // Prevent D365 form framework from capturing Enter key inside WebChat
                    chatContainer.addEventListener('keydown', function (event) {
                        if (event.key === 'Enter' && !event.shiftKey) {
                            event.stopPropagation();
                        }
                    });

                    // Wire up the restart button
                    self._restartButton.addEventListener('click', function () {
                        restartChat(chatContainer, data, self);
                    });

                    return initializeWebChat(chatContainer, data, params, self);
                })
                .catch(function (error) {
                    console.error(LOG_PREFIX + '.init: Error initializing', error);
                });
        },

        /**
         * Lifecycle hook called by the D365 control framework when the control
         * is being destroyed (e.g. form close, navigation). Properly tears
         * down all resources: observer subscriptions, the React component tree,
         * the Direct Line connection, and module-scoped state.
         *
         * @returns {void}
         */
        dispose: function () {
            // Dispose the $dyn.observe subscription
            this._pendingMessageSubscription = disposeSubscription(this._pendingMessageSubscription);

            // Invalidate any in-flight initialisation so stale callbacks bail out
            this._initId++;

            // Unmount the React tree so WebChat releases internal subscriptions,
            // timers, and Redux store closures
            if (this._chatContainer) {
                if (window.ReactDOM && window.ReactDOM.unmountComponentAtNode) {
                    window.ReactDOM.unmountComponentAtNode(this._chatContainer);
                }

                this._chatContainer = null;
            }

            // End the Direct Line connection
            if (this._directLine) {
                if (this._keepConnectionAlive) {
                    console.log(LOG_PREFIX + '.dispose: KeepConnectionAlive is set — Direct Line connection preserved');
                } else if (typeof this._directLine.end === 'function') {
                    this._directLine.end();
                    this._directLine = null;
                    console.log(LOG_PREFIX + '.dispose: Direct Line connection ended');
                }
            }

            // Reset module state so a fresh control instance starts clean
            _state.waitingForBotReply = false;
            _state.toolCalls = {};

            $dyn.ui.Control.prototype.dispose.apply(this, arguments);
        }
    });
})();