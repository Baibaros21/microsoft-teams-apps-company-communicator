// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import {
  getDraftNotifications,
  getGroups,
  getSentNotifications,
  getTeams,
  searchGroups,
  verifyGroupAccess,
} from "./apis/messageListApi";
import { formatDate } from "./i18n";
import {
  draftMessages,
  groups,
  isDraftMessagesFetchOn,
  isSentMessagesFetchOn,
  queryGroups,
  selectedMessage,
  sentMessages,
  teamsData,
  verifyGroup,
  filteredMessages,
  updateSearchQuery

} from "./messagesSlice";
import { store, RootState, useAppSelector } from "./store";
import { createSelector } from 'reselect';

type Notification = {
  createdDateTime: string;
  failed: number;
  id: string;
  isCompleted: boolean;
  sentDate: string;
  sendingStartedDate: string;
  sendingDuration: string;
  succeeded: number;
  throttled: number;
    title: string;
    seen: number;
    like: number;
    heart: number;
    surprise: number;
    laugh: number;
  totalMessageCount: number;
    createdBy: string;
}; 

const fetchSentMessages = (state: RootState) => state.messages.sentMessages;
const fetchSearchQuery = (state: RootState) => state.messages.searchQuery;

const selectFilteredMessages = createSelector([fetchSentMessages, fetchSearchQuery], (messages, searchQuery) => {

        return messages.payload.filter((message) =>
            JSON.stringify(message).includes(searchQuery.payload)
        );
    });

export const SearchQueryAction = (dispatch: typeof store.dispatch, searchQuery: string) => {
    dispatch(updateSearchQuery({ type:"FETCH_SEARCH_QUERY", searchQuery }));
};

export const FilteredMessagesAction = (dispatch: typeof store.dispatch) => {
    const fetchFilteredMessages = useAppSelector(selectFilteredMessages);
    dispatch(filteredMessages({ type: "FETCH_FILTERED_MESSAGES", fetchFilteredMessages })); 

};

export const SelectedMessageAction = (dispatch: typeof store.dispatch, payload: any) => {
  dispatch(selectedMessage({ type: "MESSAGE_SELECTED", payload }));
};

export const GetSentMessagesAction = (dispatch: typeof store.dispatch) => {
  SentMessageFetchStatusAction(dispatch, true);
  getSentNotifications()
    .then((response) => {
      const notificationList: Notification[] = response?.data || [];
      notificationList.forEach((notification) => {
        notification.sendingStartedDate = formatDate(notification.sendingStartedDate);
        notification.sentDate = formatDate(notification.sentDate);
      });
      dispatch(sentMessages({ type: "FETCH_MESSAGES", payload: notificationList || [] }));
    })
    .finally(() => {
      SentMessageFetchStatusAction(dispatch, false);
    });
};

export const GetSentMessagesSilentAction = (dispatch: typeof store.dispatch) => {
  getSentNotifications().then((response) => {
    const notificationList: Notification[] = response?.data || [];
    notificationList.forEach((notification) => {
      notification.sendingStartedDate = formatDate(notification.sendingStartedDate);
      notification.sentDate = formatDate(notification.sentDate);
    });
    dispatch(sentMessages({ type: "FETCH_MESSAGES", payload: notificationList || [] }));
  });
};

export const GetDraftMessagesAction = (dispatch: typeof store.dispatch) => {
  DraftMessageFetchStatusAction(dispatch, true);
  getDraftNotifications()
    .then((response) => {
      dispatch(draftMessages({ type: "FETCH_DRAFT_MESSAGES", payload: response?.data || [] }));
    })
    .finally(() => {
      DraftMessageFetchStatusAction(dispatch, false);
    });
};

export const GetDraftMessagesSilentAction = (dispatch: typeof store.dispatch) => {
  getDraftNotifications().then((response) => {
    dispatch(draftMessages({ type: "FETCH_DRAFT_MESSAGES", payload: response?.data || [] }));
  });
};

export const GetTeamsDataAction = (dispatch: typeof store.dispatch) => {
  getTeams().then((response) => {
    dispatch(teamsData({ type: "GET_TEAMS_DATA", payload: response?.data || [] }));
  });
};

export const GetGroupsAction = (dispatch: typeof store.dispatch, payload: { id: number }) => {
  getGroups(payload.id).then((response) => {
    dispatch(groups({ type: "GET_GROUPS", payload: response?.data || [] }));
  });
};

export const SearchGroupsAction = (dispatch: typeof store.dispatch, payload: { query: string }) => {
  searchGroups(payload.query).then((response) => {
    dispatch(queryGroups({ type: "SEARCH_GROUPS", payload: response?.data || [] }));
  });
};

export const VerifyGroupAccessAction = (dispatch: typeof store.dispatch) => {
  verifyGroupAccess()
    .then((response) => {
      dispatch(verifyGroup({ type: "VERIFY_GROUP_ACCESS", payload: true }));
    })
    .catch((error) => {
      const errorStatus = error.response.status;
      if (errorStatus === 403) {
        dispatch(verifyGroup({ type: "VERIFY_GROUP_ACCESS", payload: false }));
      } else {
        throw error;
      }
    });
};

export const DraftMessageFetchStatusAction = (dispatch: typeof store.dispatch, payload: boolean) => {
  dispatch(isDraftMessagesFetchOn({ type: "DRAFT_MESSAGES_FETCH_STATUS", payload }));
};

export const SentMessageFetchStatusAction = (dispatch: typeof store.dispatch, payload: boolean) => {
  dispatch(isSentMessagesFetchOn({ type: "SENT_MESSAGES_FETCH_STATUS", payload }));
};
