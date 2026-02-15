(function () {
    'use strict';

    var COPILOT_SDK_URL = 'https://unpkg.com/@microsoft/agents-copilotstudio-client@1.2.3/dist/src/browser.mjs';

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
        // Enable emoji support
        emojiSet: true
    };

    var MAX_RENDER_ATTEMPTS = 60;

    var waitingForBotReply = false;

    // -----------------------------------------------------------------------
    // Token acquisition via MSAL.js (delegated permissions — no secret needed)
    // -----------------------------------------------------------------------

    /**
     * Acquires a delegated user token for the Power Platform API using MSAL.js.
     * Tries silent acquisition first (leveraging the D365 Entra ID session),
     * then falls back to a popup if interaction is required.
     *
     * @param {string} appClientId  - App registration client ID (SPA / public client)
     * @param {string} tenantId     - Entra ID tenant ID
     * @returns {Promise<string>}   - The access token
     */
    function acquireToken(appClientId, tenantId) {
        var msalConfig = {
            auth: {
                clientId: appClientId,
                authority: 'https://login.microsoftonline.com/' + tenantId
            },
            cache: {
                cacheLocation: 'sessionStorage'
            }
        };

        var msalInstance = new msal.PublicClientApplication(msalConfig);

        var loginRequest = {
            scopes: ['https://api.powerplatform.com/.default'],
            redirectUri: window.location.origin
        };

        return msalInstance.initialize().then(function () {
            var accounts = msalInstance.getAllAccounts();

            if (accounts.length > 0) {
                // Try silent token acquisition using existing session
                return msalInstance.acquireTokenSilent({
                    scopes: loginRequest.scopes,
                    account: accounts[0],
                    redirectUri: loginRequest.redirectUri
                }).then(function (response) {
                    return response.accessToken;
                }).catch(function (silentError) {
                    // Silent failed — fall back to popup
                    console.warn('COTXCopilotHostControl: Silent token failed, trying popup', silentError);
                    return msalInstance.acquireTokenPopup(loginRequest).then(function (response) {
                        return response.accessToken;
                    });
                });
            }

            // No cached accounts — use popup
            return msalInstance.acquireTokenPopup(loginRequest).then(function (response) {
                return response.accessToken;
            });
        });
    }

    // -----------------------------------------------------------------------
    // WebChat store middleware — injects ERP context into outgoing messages
    // -----------------------------------------------------------------------

    function createStore(data) {
        var shouldSendContext = function () {
            return !!$dyn.peek(data.SendContext);
        };

        return window.WebChat.createStore({}, function () {
            return function (next) {
                return function (action) {
                    // AGGRESSIVE FILTER: Drop typing indicators immediately before any processing
                    if (action.type === 'DIRECT_LINE/INCOMING_ACTIVITY') {
                        var activity = action.payload && action.payload.activity;

                        // Drop typing indicators completely - don't even call next()
                        if (activity && activity.type === 'typing') {
                            return; // Exit early without processing
                        }

                    }
                        

                    if (action.type === 'DIRECT_LINE/POST_ACTIVITY'
                        && action.payload.activity.type === 'message'
                        && shouldSendContext()) {

                        var activity = action.payload.activity;

                        action = Object.assign({}, action, {
                            payload: Object.assign({}, action.payload, {
                                activity: Object.assign({}, activity, {
                                    channelData: Object.assign({}, activity.channelData, {
                                        context: {
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
                                        }
                                    })
                                })
                            })
                        });
                    }

                    // Capture incoming agent responses — only when X++ initiated the message
                    if (action.type === 'DIRECT_LINE/INCOMING_ACTIVITY') {
                        var activity = action.payload.activity;

                        if (activity.type === 'message'
                            && activity.from.role === 'bot'
                            && activity.text
                            && waitingForBotReply
                            && !activity.channelData?.isToolThought) { // Don't capture tool thoughts as final responses

                            waitingForBotReply = false;
                            $dyn.callFunction(data.RaiseAgentResponse, null, [activity.text]);
                        }
                    }

                    // Handle Dynamic Plan events (tool calls from Copilot Studio)
                    if (action.type === 'DIRECT_LINE/INCOMING_ACTIVITY') {
                        var activity = action.payload.activity;

                        if (activity && activity.type === 'event') {
                            // Track tool calls by plan identifier for correlating with thoughts
                            if (!window._copilotToolCalls) {
                                window._copilotToolCalls = {};
                            }

                            if (activity.name === 'DynamicPlanReceived' && activity.value) {
                                var planValue = activity.value;
                                var planId = planValue.planIdentifier;

                                // Store tool definitions for this plan
                                window._copilotToolCalls[planId] = {
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

                                console.log('COTXCopilotHostControl: DynamicPlanReceived', planId);
                            }

                            if (activity.name === 'DynamicPlanStepTriggered' && activity.value) {
                                var stepValue = activity.value;
                                var planId = stepValue.planIdentifier;
                                var thought = stepValue.thought || '';

                                // Check if tool calls should be shown
                                var showToolCalls = !!$dyn.peek(data.ShowToolCalls);

                                // Update the matching tool with its thought
                                if (window._copilotToolCalls[planId]) {
                                    var planData = window._copilotToolCalls[planId];

                                    // Find matching tool by taskDialogId
                                    planData.tools.forEach(function (tool) {
                                        if (stepValue.taskDialogId && stepValue.taskDialogId.indexOf(tool.identifier) !== -1) {
                                            tool.thought = thought;
                                        }
                                    });

                                    // Send the Adaptive Card if not already sent and tool calls are enabled
                                    if (!planData.cardSent && planData.tools.length > 0 && showToolCalls) {
                                        planData.cardSent = true;

                                        var adaptiveCard = {
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

                                        // Build container for each tool
                                        planData.tools.forEach(function (tool) {
                                            adaptiveCard.body.push({
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

                                        // Inject the card as a bot message
                                        var cardActivity = {
                                            type: 'message',
                                            from: { role: 'bot' },
                                            attachments: [{
                                                contentType: 'application/vnd.microsoft.card.adaptive',
                                                content: adaptiveCard
                                            }],
                                            channelData: { isToolThought: true }
                                        };

                                        next({
                                            type: 'DIRECT_LINE/INCOMING_ACTIVITY',
                                            payload: { activity: cardActivity }
                                        });
                                    }
                                }

                                console.log('COTXCopilotHostControl: DynamicPlanStepTriggered', stepValue.thought);
                            }

                            // Don't pass original event activities to WebChat
                            return;
                        }
                    }

                    return next(action);
                };
            };
        });
    }

    // -----------------------------------------------------------------------
    // Copilot Studio connection via Agent SDK
    // -----------------------------------------------------------------------

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
    // D365 F&O Extensible Control Registration
    // -----------------------------------------------------------------------

    $dyn.controls.COTXCopilotHostControl = function (data, element) {
        $dyn.ui.Control.apply(this, arguments);
        $dyn.ui.applyDefaults(this, data, $dyn.ui.defaults.COTXCopilotHostControl);
        
        // Store references for cleanup
        this._directLine = null;
        this._element = element;
    };

    $dyn.controls.COTXCopilotHostControl.prototype = $dyn.ui.extendPrototype($dyn.ui.Control.prototype, {
        init: function (data, element) {
            $dyn.ui.Control.prototype.init.apply(this, arguments);
            
            var self = this;

            function tryRender(attempt) {
                var webChatReady = element
                    && window.WebChat
                    && window.WebChat.renderWebChat
                    && window.WebChat.createStore
                    && window.msal;

                if (!webChatReady) {
                    if (attempt < MAX_RENDER_ATTEMPTS) {
                        requestAnimationFrame(function () { tryRender(attempt + 1); });
                    }
                    return;
                }

                var appClientId = $dyn.value(data.AppClientId);
                var tenantId = $dyn.value(data.TenantId);
                var environmentId = $dyn.value(data.EnvironmentId);
                var agentId = $dyn.value(data.AgentIdentifier);

                if (!appClientId || !tenantId || !environmentId || !agentId) {
                    return;
                }

                acquireToken(appClientId, tenantId)
                    .then(function (token) {
                        return createCopilotConnection(token, environmentId, agentId);
                    })
                    .then(function (directLine) {
                        // Store DirectLine reference for cleanup
                        self._directLine = directLine;
                        
                        var store = createStore(data);

                        window.WebChat.renderWebChat({
                            directLine: directLine,
                            styleOptions: STYLE_OPTIONS,
                            store: store
                        }, element);

                        // Observe PendingMessage — when X++ sets it, send it through WebChat
                        $dyn.observe(data.PendingMessage, function (message) {
                            if (!message) {
                                return;
                            }

                            waitingForBotReply = true;

                            store.dispatch({
                                type: 'WEB_CHAT/SEND_MESSAGE',
                                payload: { text: message }
                            });

                            $dyn.callFunction(
                                $dyn.observable(data.PendingMessage),
                                null,
                                ['']
                            );
                        });

                    })
                    .catch(function (e) {
                        console.error('COTXCopilotHostControl: Error initializing', e);
                    });
            }

            tryRender(0);
        }
    });
})();