(function () {
    'use strict';

    // -----------------------------------------------------------------------
    // Constants
    // -----------------------------------------------------------------------

    /** @type {string} Relative path to the Copilot Studio Agent SDK browser bundle (vendored). */
    var COPILOT_SDK_URL = '../Resources/Scripts/COTXCopilotStudioClient.js';

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
            var redirectBridgeUrl = new URL(
                'resources/html/COTXMsalRedirectBridge.html',
                window.location.href
            ).href;

            var msalConfig = {
                auth: {
                    clientId: appClientId,
                    authority: 'https://login.microsoftonline.com/' + tenantId,
                    redirectUri: redirectBridgeUrl
                },
                cache: {
                    cacheLocation: 'sessionStorage'
                }
            };

            var instance = new msal.PublicClientApplication(msalConfig);

            _msalCache[cacheKey] = {
                instance: instance,
                initPromise: instance.initialize().then(function () {
                    // Drain any stale redirect state without navigating away
                    return instance.handleRedirectPromise({
                        navigateToLoginRequestUrl: false
                    });
                })
            };
        }

        var cached = _msalCache[cacheKey];
        var msalInstance = cached.instance;

        var loginRequest = {
            scopes: ['https://api.powerplatform.com/.default']
        };

        return cached.initPromise.then(function () {
            var accounts = msalInstance.getAllAccounts();
            var matchedAccount = accounts.find(function (a) {
                return a.tenantId.toLowerCase() === tenantId.toLowerCase();
            }) || accounts[0];

            if (matchedAccount) {
                return msalInstance.acquireTokenSilent({
                    scopes: loginRequest.scopes,
                    account: matchedAccount
                }).then(function (response) {
                    return response.accessToken;
                }).catch(function (silentError) {
                    console.warn(LOG_PREFIX + '.acquireToken: Silent token failed, trying popup', silentError);
                    return msalInstance.acquireTokenPopup({
                        scopes: loginRequest.scopes,
                        loginHint: matchedAccount.username
                    }).then(function (response) {
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
     * @param {Object} action   - Redux action.
     * @param {Object} data     - D365 control data bindings.
     * @param {Object} tabState - Per-tab mutable state.
     * @returns {void}
     */
    function captureAgentResponse(action, data, tabState) {
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

        if (!tabState.waitingForBotReply) {
            return;
        }

        var isToolThought = activity.channelData && activity.channelData.isToolThought;

        if (isToolThought) {
            return;
        }

        tabState.waitingForBotReply = false;
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
     * @param {Object} tabState  - Per-tab mutable state.
     * @returns {void}
     */
    function handlePlanReceived(planValue, tabState) {
        var planId = planValue.planIdentifier;

        tabState.toolCalls[planId] = {
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
     * @param {Object}   tabState     - Per-tab mutable state.
     * @returns {void}
     */
    function handlePlanStepTriggered(stepValue, data, dispatchNext, tabState) {
        var planId = stepValue.planIdentifier;
        var thought = stepValue.thought || '';
        var planData = tabState.toolCalls[planId];

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
     * @param {Object}   tabState     - Per-tab mutable state.
     * @returns {boolean} True if the action was a plan event and was consumed
     *          (caller should **not** forward it to `next`).
     */
    function handleDynamicPlanEvents(action, data, dispatchNext, tabState) {
        if (action.type !== ACTION_TYPES.INCOMING_ACTIVITY) {
            return false;
        }

        var activity = action.payload && action.payload.activity;

        if (!activity || activity.type !== ACTIVITY_TYPES.EVENT) {
            return false;
        }

        if (activity.name === EVENT_NAMES.PLAN_RECEIVED && activity.value) {
            handlePlanReceived(activity.value, tabState);
        }

        if (activity.name === EVENT_NAMES.PLAN_STEP_TRIGGERED && activity.value) {
            handlePlanStepTriggered(activity.value, data, dispatchNext, tabState);
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
     * @param {Object} data     - D365 control data bindings.
     * @param {Object} tabState - Per-tab mutable state.
     * @returns {Object} A WebChat Redux store instance.
     */
    function createStore(data, tabState) {
        return window.WebChat.createStore({}, function () {
            return function (next) {
                return function (action) {
                    // Inject completed thoughts before suppressing the typing indicator
                    handleCompletedThoughts(action, data, next);

                    if (isTypingIndicator(action)) {
                        return;
                    }

                    action = enrichOutgoingActivity(action, data);

                    captureAgentResponse(action, data, tabState);

                    if (handleDynamicPlanEvents(action, data, next, tabState)) {
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
        var sdkUrl = new URL(COPILOT_SDK_URL, window.location.href).href;

        return import(sdkUrl).then(function (sdk) {
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
     * Creates the chat header bar (with a restart button), a tab bar strip,
     * and an inner container area for per-tab WebChat instances.
     *
     * Called once during `init`; subsequent calls are no-ops.
     *
     * @param {HTMLElement} element - The control's root DOM element.
     * @param {Object}      self   - The control instance.
     * @returns {void}
     */
       function ensureChatLayout(element, self) {
        if (self._layoutReady) {
            return;
        }

        // Tab bar
        var tabBar = document.createElement('div');
        tabBar.className = 'cotx-tab-bar';

        var tabList = document.createElement('div');
        tabList.className = 'cotx-tab-list';

        var addTabBtn = document.createElement('button');
        addTabBtn.className = 'cotx-tab-add-btn';
        addTabBtn.type = 'button';
        addTabBtn.title = 'New conversation tab';
        addTabBtn.textContent = '+';

        var restartBtn = document.createElement('button');
        restartBtn.className = 'cotx-tab-restart-btn';
        restartBtn.type = 'button';
        restartBtn.title = 'Restart current conversation';
        restartBtn.textContent = '\u21BB';

        tabBar.appendChild(tabList);
        tabBar.appendChild(addTabBtn);
        tabBar.appendChild(restartBtn);

        // Container area where per-tab chat containers live
        var containerArea = document.createElement('div');
        containerArea.className = 'cotx-tab-container-area';

        element.appendChild(tabBar);
        element.appendChild(containerArea);

        self._restartButton = restartBtn;
        self._tabList = tabList;
        self._addTabBtn = addTabBtn;
        self._containerArea = containerArea;
        self._layoutReady = true;
    }

    // -----------------------------------------------------------------------
    // Tab management
    // -----------------------------------------------------------------------

    /**
     * Creates a new conversation tab, initialises its WebChat connection, and
     * switches to it.
     *
     * @param {Object} data   - D365 control data bindings.
     * @param {Object} params - Validated control parameters.
     * @param {Object} self   - The control instance.
     * @returns {string|null} The new tab's ID, or null if the limit was reached.
     */
    function createTab(data, params, self) {
        var tm = self._tabManager;

        if (tm.tabOrder.length >= tm.maxTabs) {
            console.warn(LOG_PREFIX + '.createTab: Maximum tabs reached (' + tm.maxTabs + ')');
            return null;
        }

        var tabId = 'tab-' + Date.now() + '-' + Math.random().toString(36).substr(2, 5);
        var tabName = 'Chat ' + tm.nextTabNum++;

        var tab = {
            id: tabId,
            name: tabName,
            container: null,
            tabButton: null,
            tabLabel: null,
            directLine: null,
            store: null,
            state: { waitingForBotReply: false, toolCalls: {} },
            pendingMessageSubscription: null,
            keepConnectionAlive: false,
            initId: 0,
            initialized: false
        };

        // Create chat container for this tab
        var container = document.createElement('div');
        container.className = 'cotx-chat-container';
        container.style.display = 'none';
        container.setAttribute('data-tab-id', tabId);
        self._containerArea.appendChild(container);
        tab.container = container;

        // Prevent D365 form framework from capturing Enter key inside WebChat
        container.addEventListener('keydown', function (event) {
            if (event.key === 'Enter' && !event.shiftKey) {
                event.stopPropagation();
            }
        });

        // Create tab button
        var tabBtn = document.createElement('button');
        tabBtn.className = 'cotx-tab-btn';
        tabBtn.type = 'button';
        tabBtn.setAttribute('data-tab-id', tabId);

        var tabLabel = document.createElement('span');
        tabLabel.className = 'cotx-tab-label';
        tabLabel.textContent = tabName;
        tabBtn.appendChild(tabLabel);

        var closeBtn = document.createElement('span');
        closeBtn.className = 'cotx-tab-close';
        closeBtn.textContent = '\u00D7';
        closeBtn.title = 'Close tab';
        tabBtn.appendChild(closeBtn);

        closeBtn.addEventListener('click', function (e) {
            e.stopPropagation();
            if (tm.tabOrder.length > 1) {
                closeTabById(tabId, data, self);
            }
        });

        tabBtn.addEventListener('click', function () {
            switchToTab(tabId, self);
        });

        // Double-click tab label to rename
        tabLabel.addEventListener('dblclick', function (e) {
            e.stopPropagation();
            startTabRename(tabId, self);
        });

        self._tabList.appendChild(tabBtn);
        tab.tabButton = tabBtn;
        tab.tabLabel = tabLabel;

        // Register in tab manager
        tm.tabs[tabId] = tab;
        tm.tabOrder.push(tabId);

        // Switch to new tab
        switchToTab(tabId, self);

        // Initialise WebChat in the tab
        initializeWebChat(container, data, params, self, tab)
            .then(function () {
                tab.initialized = true;
                console.log(LOG_PREFIX + '.createTab: Tab initialized \u2014 ' + tabName);
            })
            .catch(function (error) {
                console.error(LOG_PREFIX + '.createTab: Error initializing tab', error);

                // Show error message inside the tab container
                if (tab.container) {
                    tab.container.innerHTML = '<div style="padding:16px;color:#C00;font-size:12px;">' +
                        '\u26A0 Failed to initialise conversation. Please close this tab and try again.' +
                        '</div>';
                }
            });

        updateTabCloseButtons(self);
        return tabId;
    }

    /**
     * Switches the visible conversation to the specified tab.
     *
     * @param {string} tabId - The tab ID to activate.
     * @param {Object} self  - The control instance.
     * @returns {void}
     */
    function switchToTab(tabId, self) {
        var tm = self._tabManager;
        if (!tm.tabs[tabId]) { return; }

        // Hide current tab
        if (tm.activeTabId && tm.tabs[tm.activeTabId]) {
            var currentTab = tm.tabs[tm.activeTabId];
            currentTab.container.style.display = 'none';
            if (currentTab.tabButton) {
                currentTab.tabButton.classList.remove('cotx-tab-active');
            }
        }

        // Show target tab
        var newTab = tm.tabs[tabId];
        newTab.container.style.display = '';
        if (newTab.tabButton) {
            newTab.tabButton.classList.add('cotx-tab-active');
        }

        tm.activeTabId = tabId;
    }

    /**
     * Closes and tears down a conversation tab.
     *
     * @param {string} tabId - The tab ID to close.
     * @param {Object} data  - D365 control data bindings.
     * @param {Object} self  - The control instance.
     * @returns {void}
     */
    function closeTabById(tabId, data, self) {
        var tm = self._tabManager;
        var tab = tm.tabs[tabId];
        if (!tab) { return; }
        if (tm.tabOrder.length <= 1) { return; }

        // Dispose resources
        tab.pendingMessageSubscription = disposeSubscription(tab.pendingMessageSubscription);
        tab.initId++;

        if (tab.container && window.ReactDOM && window.ReactDOM.unmountComponentAtNode) {
            window.ReactDOM.unmountComponentAtNode(tab.container);
        }

        if (tab.directLine && typeof tab.directLine.end === 'function') {
            if (!tab.keepConnectionAlive) {
                tab.directLine.end();
            }
            tab.directLine = null;
        }

        // Remove DOM elements
        if (tab.container && tab.container.parentNode) {
            tab.container.parentNode.removeChild(tab.container);
        }
        if (tab.tabButton && tab.tabButton.parentNode) {
            tab.tabButton.parentNode.removeChild(tab.tabButton);
        }

        // Remove from tab manager
        var idx = tm.tabOrder.indexOf(tabId);
        tm.tabOrder.splice(idx, 1);
        delete tm.tabs[tabId];

        // Switch to nearest tab if this was the active one
        if (tm.activeTabId === tabId) {
            var newIdx = Math.min(idx, tm.tabOrder.length - 1);
            switchToTab(tm.tabOrder[newIdx], self);
        }

        updateTabCloseButtons(self);
    }

    /**
     * Displays an inline rename input on the tab label.
     *
     * @param {string} tabId - The tab ID to rename.
     * @param {Object} self  - The control instance.
     * @returns {void}
     */
    function startTabRename(tabId, self) {
        var tab = self._tabManager.tabs[tabId];
        if (!tab || !tab.tabLabel) { return; }

        var label = tab.tabLabel;
        var currentName = label.textContent;

        var input = document.createElement('input');
        input.type = 'text';
        input.className = 'cotx-tab-rename-input';
        input.value = currentName;

        label.textContent = '';
        label.appendChild(input);
        input.focus();
        input.select();

        function finishRename() {
            var newName = input.value.trim() || currentName;
            tab.name = newName;
            if (label.contains(input)) {
                label.removeChild(input);
            }
            label.textContent = newName;
        }

        input.addEventListener('blur', finishRename);
        input.addEventListener('keydown', function (e) {
            if (e.key === 'Enter') { input.blur(); }
            if (e.key === 'Escape') { input.value = currentName; input.blur(); }
            e.stopPropagation();
        });
    }

    /**
     * Shows or hides the close button on each tab depending on whether more
     * than one tab is open (the last tab cannot be closed). Also enables or
     * disables the add-tab button when the tab limit is reached.
     *
     * @param {Object} self - The control instance.
     * @returns {void}
     */
    function updateTabCloseButtons(self) {
        var tm = self._tabManager;
        var hideClose = tm.tabOrder.length <= 1;

        tm.tabOrder.forEach(function (tabId) {
            var tab = tm.tabs[tabId];
            if (tab && tab.tabButton) {
                var closeEl = tab.tabButton.querySelector('.cotx-tab-close');
                if (closeEl) {
                    closeEl.style.display = hideClose ? 'none' : '';
                }
            }
        });

        // Disable the add-tab button when the limit is reached
        if (self._addTabBtn) {
            var atLimit = tm.tabOrder.length >= tm.maxTabs;
            self._addTabBtn.disabled = atLimit;
            self._addTabBtn.title = atLimit
                ? 'Maximum of ' + tm.maxTabs + ' tabs reached'
                : 'New conversation tab';
        }
    }

    // -----------------------------------------------------------------------
    // Chat restart
    // -----------------------------------------------------------------------

    /**
     * Tears down the active tab's WebChat session and bootstraps a fresh one.
     *
     * Cleanup steps:
     *  1. Disposes the `$dyn.observe` subscription on PendingMessage.
     *  2. Ends the Direct Line connection (frees server-side resources).
     *  3. Unmounts the React component tree via `ReactDOM.unmountComponentAtNode`
     *     so that all internal subscriptions, timers, and closures held by
     *     WebChat and its Redux store are properly released.
     *  4. Resets per-tab mutable state.
     *
     * A new token, connection, store, and React tree are then created via
     * {@link initializeWebChat}.
     *
     * @param {Object}      data          - D365 control data bindings.
     * @param {Object}      self          - The control instance.
     * @returns {Promise<void>}
     */
    function restartChat(data, self) {
        var tm = self._tabManager;
        var tab = tm.tabs[tm.activeTabId];
        if (!tab) { return Promise.resolve(); }

        // Double-click protection
        if (self._restartButton) {
            self._restartButton.disabled = true;
        }

        // 1. Dispose $dyn.observe subscription
        tab.pendingMessageSubscription = disposeSubscription(tab.pendingMessageSubscription);

        // 2. End the Direct Line connection
        if (tab.directLine && typeof tab.directLine.end === 'function') {
            tab.directLine.end();
            tab.directLine = null;
        }

        // 3. Properly unmount the React component tree so WebChat releases
        //    all internal subscriptions, timers, and Redux store closures.
        //    ReactDOM is exposed globally by the WebChat CDN full bundle.
        if (tab.container && window.ReactDOM && window.ReactDOM.unmountComponentAtNode) {
            window.ReactDOM.unmountComponentAtNode(tab.container);
        }

        // 4. Reset per-tab state
        tab.store = null;
        tab.state.waitingForBotReply = false;
        tab.state.toolCalls = {};

        // 5. Re-initialise
        var params = readControlParameters(data);

        if (!params) {
            if (self._restartButton) {
                self._restartButton.disabled = false;
            }

            return Promise.reject(new Error(LOG_PREFIX + '.restartChat: Missing control parameters'));
        }

        return initializeWebChat(tab.container, data, params, self, tab)
            .then(function () {
                if (self._restartButton) {
                    self._restartButton.disabled = false;
                }

                console.log(LOG_PREFIX + '.restartChat: Tab restarted \u2014 ' + tab.name);
            })
            .catch(function (error) {
                if (self._restartButton) {
                    self._restartButton.disabled = false;
                }

                console.error(LOG_PREFIX + '.restartChat: Error restarting tab', error);
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
            agentId: $dyn.peek(data.AgentIdentifier),
        };

        if (!params.appClientId || !params.tenantId || !params.environmentId || !params.agentId) {
            console.warn(LOG_PREFIX + '.readControlParameters: One or more required parameters are missing.', params);
            return null;
        }

        return params;
    }

    /**
     * Wires up the `PendingMessage` observer so that when X++ sets a value,
     * the message is dispatched through the active tab's WebChat store and
     * the observable is cleared. Only the active tab processes the message;
     * inactive tabs ignore it.
     *
     * @param {Object} data      - D365 control data bindings.
     * @param {Object} store     - WebChat Redux store.
     * @param {Object} tabState  - Per-tab mutable state.
     * @param {string} tabId     - The tab this subscription belongs to.
     * @param {Object} self      - The control instance.
     * @returns {Object} The `$dyn.observe` subscription (call `.dispose()` to
     *          unsubscribe).
     */
    function observePendingMessages(data, store, tabState, tabId, self) {
        return $dyn.observe(data.PendingMessage, function (message) {
            if (!message) {
                return;
            }

            // Only the active tab should process PendingMessage
            if (self._tabManager.activeTabId !== tabId) {
                return;
            }

            tabState.waitingForBotReply = true;

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
     * Uses an incrementing `initId` on the tab object to detect stale
     * initialisations — if a restart is triggered while a previous init is
     * in-flight, the earlier init's `.then` callbacks silently bail out.
     *
     * @param {HTMLElement} element - The DOM element that will host WebChat.
     * @param {Object}      data   - D365 control data bindings.
     * @param {{ appClientId: string, tenantId: string, environmentId: string, agentId: string }} params
     *        - Validated control parameters.
     * @param {Object}      self   - The control instance.
     * @param {Object}      tab    - The tab object to initialise.
     * @returns {Promise<void>}
     */
    function initializeWebChat(element, data, params, self, tab) {
        var myInitId = ++tab.initId;

        return acquireToken(params.appClientId, params.tenantId)
            .then(function (token) {
                if (tab.initId !== myInitId) { return; }

                return createCopilotConnection(token, params.environmentId, params.agentId);
            })
            .then(function (directLine) {
                if (!directLine || tab.initId !== myInitId) { return; }

                tab.directLine = directLine;
                tab.keepConnectionAlive = !!$dyn.peek(data.KeepConnectionAlive);

                var store = createStore(data, tab.state);
                tab.store = store;

                window.WebChat.renderWebChat({
                    directLine: directLine,
                    styleOptions: STYLE_OPTIONS,
                    store: store
                }, element);

                tab.pendingMessageSubscription = observePendingMessages(data, store, tab.state, tab.id, self);
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

        /** @type {boolean} True once the layout shell has been created. */
        this._layoutReady = false;

        /** @type {HTMLElement|null} The restart button element. */
        this._restartButton = null;

        /** @type {HTMLElement|null} The tab list container. */
        this._tabList = null;

        /** @type {HTMLElement|null} The add-tab button. */
        this._addTabBtn = null;

        /** @type {HTMLElement|null} Container area where per-tab chat containers live. */
        this._containerArea = null;

        /**
         * Per-instance tab manager — holds all open conversation tabs and tracks
         * the active one. Instance-scoped to avoid collisions when multiple
         * controls exist on the same page.
         *
         * @type {{ tabs: Object, activeTabId: string|null, tabOrder: string[], nextTabNum: number, maxTabs: number }}
         */
        this._tabManager = {
            tabs: {},
            activeTabId: null,
            tabOrder: [],
            nextTabNum: 1,
            maxTabs: 8
        };
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

                    ensureChatLayout(element, self);

                    // Wire up the restart button (restarts the active tab)
                    self._restartButton.addEventListener('click', function () {
                        restartChat(data, self);
                    });

                    // Wire up the add-tab button
                    self._addTabBtn.addEventListener('click', function () {
                        createTab(data, params, self);
                    });

                    // Create the first conversation tab
                    createTab(data, params, self);
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
            var tm = this._tabManager;

            // Tear down every tab
            tm.tabOrder.slice().forEach(function (tabId) {
                var tab = tm.tabs[tabId];
                if (!tab) { return; }

                tab.pendingMessageSubscription = disposeSubscription(tab.pendingMessageSubscription);
                tab.initId++;

                if (tab.container && window.ReactDOM && window.ReactDOM.unmountComponentAtNode) {
                    window.ReactDOM.unmountComponentAtNode(tab.container);
                }

                if (tab.directLine) {
                    if (tab.keepConnectionAlive) {
                        console.log(LOG_PREFIX + '.dispose: KeepConnectionAlive — tab ' + tab.name + ' preserved');
                    } else if (typeof tab.directLine.end === 'function') {
                        tab.directLine.end();
                        tab.directLine = null;
                        console.log(LOG_PREFIX + '.dispose: DirectLine ended for tab ' + tab.name);
                    }
                }
            });

            // Reset tab manager
            tm.tabs = {};
            tm.activeTabId = null;
            tm.tabOrder = [];
            tm.nextTabNum = 1;

            this._layoutReady = false;
            this._containerArea = null;
            this._tabList = null;
            this._addTabBtn = null;
            this._restartButton = null;

            $dyn.ui.Control.prototype.dispose.apply(this, arguments);
        }
    });
})();