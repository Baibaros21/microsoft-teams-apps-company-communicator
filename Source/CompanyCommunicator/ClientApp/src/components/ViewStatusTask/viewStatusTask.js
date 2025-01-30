"use strict";
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.ViewStatusTask = void 0;
var AdaptiveCards = require("adaptivecards");
var React = require("react");
var react_i18next_1 = require("react-i18next");
var react_router_dom_1 = require("react-router-dom");
var react_components_1 = require("@fluentui/react-components");
var react_icons_1 = require("@fluentui/react-icons");
var microsoftTeams = require("@microsoft/teams-js");
var messageListApi_1 = require("../../apis/messageListApi");
var messageListApi_2 = require("../../apis/messageListApi");
var i18n_1 = require("../../i18n");
var adaptiveCard_1 = require("../AdaptiveCard/adaptiveCard");
var store_1 = require("../../store");
var card;
var ViewStatusTask = function () {
    var t = (0, react_i18next_1.useTranslation)().t;
    var id = (0, react_router_dom_1.useParams)().id;
    var _a = React.useState(true), loader = _a[0], setLoader = _a[1];
    var _b = React.useState(false), isCardReady = _b[0], setIsCardReady = _b[1];
    var _c = React.useState(false), exportDisabled = _c[0], setExportDisabled = _c[1];
    var _d = React.useState(''), cardAreaBorderClass = _d[0], setCardAreaBorderClass = _d[1];
    var _e = React.useState({
        logoFileName: "",
        logoLink: "",
        bannerLink: "",
        bannerFileName: ""
    }), defaultsState = _e[0], setDefaultState = _e[1];
    var _f = React.useState({
        id: '',
        title: '',
        isMsgDataUpdated: false,
        template: store_1.TemplateSelection.Default
    }), messageState = _f[0], setMessageState = _f[1];
    var _g = React.useState({
        page: 'ViewStatus',
        teamId: '',
        isTeamDataUpdated: false,
    }), statusState = _g[0], setStatusState = _g[1];
    React.useEffect(function () {
        microsoftTeams.getContext(function (context) {
            setStatusState(__assign(__assign({}, statusState), { teamId: context.teamId, isTeamDataUpdated: true }));
        });
    }, []);
    React.useEffect(function () {
        if (id) {
            getMessage(id);
        }
    }, [id]);
    React.useEffect(function () {
        if (isCardReady && messageState.isMsgDataUpdated) {
            var adaptiveCard = new AdaptiveCards.AdaptiveCard();
            adaptiveCard.parse(card);
            var renderCard = adaptiveCard.render();
            if (renderCard && statusState.page === 'ViewStatus') {
                document.getElementsByClassName('card-area-1')[0].appendChild(renderCard);
                setCardAreaBorderClass('card-area-border');
            }
            adaptiveCard.onExecuteAction = function (action) {
                window.open(action.url, '_blank');
            };
            setLoader(false);
        }
    }, [isCardReady, messageState.isMsgDataUpdated]);
    var getMessage = function (id) { return __awaiter(void 0, void 0, void 0, function () {
        var error_1;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    getDefaultsItem();
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 3, , 4]);
                    return [4 /*yield*/, (0, messageListApi_2.getSentNotification)(id).then(function (response) {
                            updateCardData(response.data);
                            response.data.sendingDuration = (0, i18n_1.formatDuration)(response.data.sendingStartedDate, response.data.sentDate);
                            response.data.sendingStartedDate = (0, i18n_1.formatDate)(response.data.sendingStartedDate);
                            response.data.sentDate = (0, i18n_1.formatDate)(response.data.sentDate);
                            response.data.succeeded = (0, i18n_1.formatNumber)(response.data.succeeded);
                            response.data.failed = (0, i18n_1.formatNumber)(response.data.failed);
                            response.data.seen = (0, i18n_1.formatNumber)(response.data.seen);
                            response.data.unknown = response.data.unknown && (0, i18n_1.formatNumber)(response.data.unknown);
                            response.data.canceled = response.data.canceled && (0, i18n_1.formatNumber)(response.data.canceled);
                            setMessageState(__assign(__assign({}, response.data), { isMsgDataUpdated: true }));
                        })];
                case 2:
                    _a.sent();
                    return [3 /*break*/, 4];
                case 3:
                    error_1 = _a.sent();
                    return [2 /*return*/, error_1];
                case 4: return [2 /*return*/];
            }
        });
    }); };
    var getDefaultsItem = function () { return __awaiter(void 0, void 0, void 0, function () {
        var error_2;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    _a.trys.push([0, 2, , 3]);
                    return [4 /*yield*/, (0, messageListApi_1.getDefaultData)().then(function (response) {
                            var defaultImages = response.data;
                            console.log(defaultImages);
                            setDefaultState({
                                logoFileName: defaultImages.logoFileName,
                                logoLink: defaultImages.logoLink,
                                bannerFileName: defaultImages.bannerFileName,
                                bannerLink: defaultImages.bannerLink
                            });
                        })];
                case 1:
                    _a.sent();
                    return [3 /*break*/, 3];
                case 2:
                    error_2 = _a.sent();
                    return [2 /*return*/, error_2];
                case 3: return [2 /*return*/];
            }
        });
    }); };
    var updateCardData = function (msg) { return __awaiter(void 0, void 0, void 0, function () {
        return __generator(this, function (_a) {
            console.log(msg.card);
            if (msg.card) {
                card = JSON.parse(msg.card);
                (0, adaptiveCard_1.setCardTitle)(card, msg.title);
                (0, adaptiveCard_1.setCardImageLink)(card, msg.imageLink);
                (0, adaptiveCard_1.setCardSummary)(card, msg.summary);
                (0, adaptiveCard_1.setCardAuthor)(card, msg.author);
                (0, adaptiveCard_1.setCardDeptTitle)(card, msg.department);
                (0, adaptiveCard_1.setCardVideoPlayerUrl)(card, msg.videoLink);
                (0, adaptiveCard_1.setCardVideoPlayerPoster)(card, msg.posterLink);
                (0, adaptiveCard_1.setCardLogo)(card, defaultsState.logoLink);
                (0, adaptiveCard_1.setCardBanner)(card, defaultsState.bannerLink);
                if (msg.buttonTitle && msg.buttonLink) {
                    (0, adaptiveCard_1.setCardBtn)(card, msg.buttonTitle, msg.buttonLink);
                }
                setIsCardReady(true);
            }
            return [2 /*return*/];
        });
    }); };
    var onClose = function () {
        microsoftTeams.tasks.submitTask();
    };
    var onExport = function () { return __awaiter(void 0, void 0, void 0, function () {
        var payload;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    setExportDisabled(true);
                    payload = {
                        id: messageState.id,
                        teamId: statusState.teamId,
                    };
                    return [4 /*yield*/, (0, messageListApi_2.exportNotification)(payload)
                            .then(function () {
                            setStatusState(__assign(__assign({}, statusState), { page: 'SuccessPage' }));
                        })
                            .catch(function () {
                            setStatusState(__assign(__assign({}, statusState), { page: 'ErrorPage' }));
                        })
                            .finally(function () {
                            setExportDisabled(false);
                        })];
                case 1:
                    _a.sent();
                    return [2 /*return*/];
            }
        });
    }); };
    var getItemList = function (items, secondaryText, shape) {
        var resultedTeams = [];
        if (items) {
            items.map(function (element) {
                resultedTeams.push(React.createElement("li", { key: element + 'key' },
                    React.createElement(react_components_1.Persona, { name: element, secondaryText: secondaryText, avatar: { shape: shape, color: 'colorful' } })));
            });
        }
        return resultedTeams;
    };
    var renderAudienceSelection = function () {
        if (messageState.teamNames && messageState.teamNames.length > 0) {
            return (React.createElement(react_components_1.Field, { size: 'large', label: t('SentToGeneralChannel') },
                React.createElement("ul", { className: 'ul-no-bullets' }, getItemList(messageState.teamNames, 'Team', 'square'))));
        }
        else if (messageState.rosterNames && messageState.rosterNames.length > 0) {
            return (React.createElement(react_components_1.Field, { size: 'large', label: t('SentToRosters') },
                React.createElement("ul", { className: 'ul-no-bullets' }, getItemList(messageState.rosterNames, 'Team', 'square'))));
        }
        else if (messageState.groupNames && messageState.groupNames.length > 0) {
            return (React.createElement(react_components_1.Field, { size: 'large', label: t('SentToGroups1') },
                React.createElement("span", null, t('SentToGroups2')),
                React.createElement("ul", { className: 'ul-no-bullets' }, getItemList(messageState.groupNames, 'Group', 'circular'))));
        }
        else if (messageState.allUsers) {
            return (React.createElement(React.Fragment, null,
                React.createElement(react_components_1.Text, { size: 500 }, t('SendToAllUsers'))));
        }
        else {
            return React.createElement("div", null);
        }
    };
    var renderErrorMessage = function () {
        if (messageState.errorMessage) {
            return (React.createElement("div", null,
                React.createElement(react_components_1.Field, { size: 'large', label: t('Errors') },
                    React.createElement(react_components_1.Text, { className: 'info-text' }, messageState.errorMessage))));
        }
        else {
            return React.createElement("div", null);
        }
    };
    var renderWarningMessage = function () {
        if (messageState.warningMessage) {
            return (React.createElement("div", null,
                React.createElement(react_components_1.Field, { size: 'large', label: t('Warnings') },
                    React.createElement(react_components_1.Text, { className: 'info-text' }, messageState.warningMessage))));
        }
        else {
            return React.createElement("div", null);
        }
    };
    return (React.createElement(React.Fragment, null,
        loader && React.createElement(react_components_1.Spinner, null),
        statusState.page === 'ViewStatus' && (React.createElement(React.Fragment, null,
            React.createElement("span", { role: 'alert', "aria-label": t('ViewMessageStatus') }),
            React.createElement("div", { className: 'adaptive-task-grid' },
                React.createElement("div", { className: 'form-area' }, !loader && (React.createElement(React.Fragment, null,
                    React.createElement("div", { style: { paddingBottom: '16px' } },
                        React.createElement(react_components_1.Field, { size: 'large', label: t('TitleText') },
                            React.createElement(react_components_1.Text, { style: { overflowWrap: 'anywhere' } }, messageState.title))),
                    React.createElement("div", { style: { paddingBottom: '16px' } },
                        React.createElement(react_components_1.Field, { className: 'spacingVerticalM', size: 'large', label: t('SendingStarted') },
                            React.createElement(react_components_1.Text, null, messageState.sendingStartedDate))),
                    React.createElement("div", { style: { paddingBottom: '16px' } },
                        React.createElement(react_components_1.Field, { size: 'large', label: t('Completed') },
                            React.createElement(react_components_1.Text, null, messageState.sentDate))),
                    React.createElement("div", { style: { paddingBottom: '16px' } },
                        React.createElement(react_components_1.Field, { size: 'large', label: t('CreatedBy') },
                            React.createElement(react_components_1.Persona, { name: messageState.createdBy, secondaryText: 'Member', avatar: { color: 'colorful' } }))),
                    React.createElement("div", { style: { paddingBottom: '16px' } },
                        React.createElement(react_components_1.Field, { size: 'large', label: t('Duration') },
                            React.createElement(react_components_1.Text, null, messageState.sendingDuration))),
                    React.createElement("div", { style: { paddingBottom: '16px' } },
                        React.createElement(react_components_1.Field, { size: 'large', label: t('Seen') },
                            React.createElement(react_components_1.Text, null, messageState.seen))),
                    React.createElement("div", { style: { paddingBottom: '16px' } },
                        React.createElement(react_components_1.Field, { size: 'large', label: t('Results') },
                            React.createElement(react_components_1.Text, null, t('Success', { SuccessCount: messageState.succeeded })),
                            React.createElement(react_components_1.Text, null, t('Failure', { FailureCount: messageState.failed })),
                            messageState.unknown && (React.createElement(React.Fragment, null,
                                React.createElement(react_components_1.Text, null, t('Unknown', { UnknownCount: messageState.unknown })))))),
                    React.createElement("div", { style: { paddingBottom: '16px' } },
                        renderAudienceSelection(),
                        renderErrorMessage(),
                        renderWarningMessage())))),
                React.createElement("div", { className: 'card-area' },
                    React.createElement("div", { className: cardAreaBorderClass },
                        React.createElement("div", { className: 'card-area-1' })))),
            React.createElement("div", null,
                React.createElement("div", { className: 'fixed-footer' },
                    React.createElement("div", { className: 'footer-action-right' },
                        React.createElement("div", { className: 'footer-actions-flex' },
                            exportDisabled && React.createElement(react_components_1.Spinner, { role: 'alert', size: 'small', label: t('ExportLabel'), labelPosition: 'after' }),
                            React.createElement(react_components_1.Button, { icon: React.createElement(react_icons_1.ArrowDownload24Regular, null), style: { marginLeft: '16px' }, title: exportDisabled || messageState.canDownload === false ? t('ExportButtonProgressText') : t('ExportButtonText'), disabled: exportDisabled || messageState.canDownload === false, onClick: onExport, appearance: 'primary' }, t('ExportButtonText')))))))),
        !loader && statusState.page === 'SuccessPage' && (React.createElement(React.Fragment, null,
            React.createElement("span", { role: 'alert', "aria-label": t('ExportSuccessView') }),
            React.createElement("div", { className: 'wizard-page' },
                React.createElement("h2", null,
                    React.createElement(react_icons_1.CheckmarkSquare24Regular, { style: { color: '#22bb33', verticalAlign: 'top', paddingRight: '4px' } }),
                    t('ExportQueueTitle')),
                React.createElement(react_components_1.Text, null, t('ExportQueueSuccessMessage1')),
                React.createElement("br", null),
                React.createElement("br", null),
                React.createElement(react_components_1.Text, null, t('ExportQueueSuccessMessage2')),
                React.createElement("br", null),
                React.createElement("br", null),
                React.createElement(react_components_1.Text, null, t('ExportQueueSuccessMessage3')),
                React.createElement("br", null),
                React.createElement("br", null),
                React.createElement("div", { className: 'fixed-footer' },
                    React.createElement("div", { className: 'footer-action-right' },
                        React.createElement(react_components_1.Button, { id: 'closeBtn', onClick: onClose, appearance: 'primary' }, t('CloseText'))))))),
        !loader && statusState.page === 'ErrorPage' && (React.createElement(React.Fragment, null,
            React.createElement("span", { role: 'alert', "aria-label": t('ExportFailureView') }),
            React.createElement("div", { className: 'wizard-page' },
                React.createElement("h2", null,
                    React.createElement(react_icons_1.ShareScreenStop24Regular, { style: { color: '#bb2124', verticalAlign: 'top', paddingRight: '4px' } }),
                    t('ExportErrorTitle')),
                React.createElement(react_components_1.Text, null, t('ExportErrorMessage')),
                React.createElement("br", null),
                React.createElement("div", { className: 'fixed-footer' },
                    React.createElement("div", { className: 'footer-action-right' },
                        React.createElement(react_components_1.Button, { id: 'closeBtn', onClick: onClose, appearance: 'primary' }, t('CloseText')))))))));
};
exports.ViewStatusTask = ViewStatusTask;
//# sourceMappingURL=viewStatusTask.js.map