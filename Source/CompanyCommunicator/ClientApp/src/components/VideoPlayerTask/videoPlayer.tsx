import * as AdaptiveCards from 'adaptivecards';
import * as React from 'react';
import './videoPlayer.scss';
import { useTranslation } from 'react-i18next';
import { useParams } from 'react-router-dom';
import { Button, Field, Label, Persona, Spinner, Text } from '@fluentui/react-components';
import * as microsoftTeams from '@microsoft/teams-js';
import {
    getConsentSummaries, getDraftNotification, sendDraftNotification
} from '../../apis/messageListApi';
import {
    getInitAdaptiveCard, setCardVideoPlayerPoster, setCardVideoUrl, setCardVideoPoster
} from '../AdaptiveCard/adaptiveCard';
import { AvatarShape } from '@fluentui/react-avatar';
import { getBaseUrl } from '../../configVariables';
import { TemplateSelection } from "../../store";

let card: any;


export const VideoPlayer = () => {

    const { t } = useTranslation();
    const { id } = useParams() as any;
    const [loader, setLoader] = React.useState(true);

    const videoUrl = "https://samplelib.com/lib/preview/mp4/sample-5s.mp4";
    const imageLink = "https://adaptivecards.io/content/poster-video.png";
    const buttonLink = "";

    React.useEffect(() => {

        card = getInitAdaptiveCard("Tittle", "videoPlayer");
        setCardVideoPoster(card, imageLink);
        setCardVideoUrl(card, videoUrl);
        console.log("videoUrl: " + videoUrl);
        let adaptiveCard = new AdaptiveCards.AdaptiveCard();
        adaptiveCard.parse(card);
        let renderedCard = adaptiveCard.render();
        document.getElementsByClassName('card-area')[0].innerHTML = "";
        document.getElementsByClassName('card-area')[0].appendChild(renderedCard!);


    }, []);
    return (
        <>
            <div className="taskModule">
                <div className="adaptive-task-grid">
                    <div className="card-area">

                    </div>
                </div>
            </div>
        </>
    );

};