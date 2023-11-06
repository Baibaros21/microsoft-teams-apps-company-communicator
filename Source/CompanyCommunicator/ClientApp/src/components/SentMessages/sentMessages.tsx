// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import { useTranslation } from "react-i18next";
import { Spinner, makeStyles } from "@fluentui/react-components";
import { GetSentMessagesAction, GetSentMessagesSilentAction } from "../../actions";
import { RootState, useAppDispatch, useAppSelector } from "../../store";
import { SentMessageDetail } from "../MessageDetail/sentMessageDetail";
import * as CustomHooks from "../../useInterval";
import { Search12Regular } from "@fluentui/react-icons";
import {
    Input,
    Label,
} from "@fluentui/react-components";


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
const useStyles = makeStyles({
    root: {
        display: "inline-block",
        float: "right",
        position: "relative"
    }
});

export const SentMessages = () => {
    const { t } = useTranslation();
    const sentMessages = useAppSelector((state: RootState) => state.messages).sentMessages.payload;
    const loader = useAppSelector((state: RootState) => state.messages).isSentMessagesFetchOn.payload;
    const dispatch = useAppDispatch();
    const delay = 60000;
    const [filteredMessages, setFilteredMessages] = React.useState<any>([]);
    const [filterQuery, setFilterQuery] = React.useState<string>(""); 

    React.useEffect(() => {
        if (sentMessages && sentMessages.length === 0) {
            GetSentMessagesAction(dispatch);
            setFilteredMessages(sentMessages);
        }
    }, []);

    React.useEffect(() => {

        if (sentMessages && sentMessages.length > 0) {

            if (filterQuery !== "") { }
            var filter = sentMessages.filter((message: Notification) => {
                if (message.title.toLocaleLowerCase() === filterQuery.toLocaleLowerCase()) {
                    return true;
                }

                return message.title.toLocaleLowerCase().includes(filterQuery.toLocaleLowerCase());
            })
            setFilteredMessages(filter);
        } else {
            setFilteredMessages(sentMessages);
        }
        
        
    }, [sentMessages, filterQuery]);

    const onQueryChange = (e : any) => {
         setFilterQuery(e.target.value);
  
    }

    CustomHooks.useInterval(() => {
        GetSentMessagesSilentAction(dispatch);
    }, delay);

    var classes = useStyles();

    return (
        <>
            {loader && <Spinner labelPosition="below" label="Fetching..." />}
            {sentMessages && sentMessages.length === 0 && !loader && <div>{t("EmptySentMessages")}</div>}
            {sentMessages && sentMessages.length > 0 && !loader &&
                (<>
                <div className={classes.root}>

                    <Input onChange={onQueryChange} placeholder="search" contentBefore={<Search12Regular />} id={"search"} />
                </div>
                <SentMessageDetail sentMessages={filteredMessages} />

            </>
                )
            }
            </>
  );
};
