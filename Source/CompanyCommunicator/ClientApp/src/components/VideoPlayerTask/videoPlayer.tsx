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
    getInitAdaptiveCard, setCardAuthor, setCardBtn, setCardImageLink, setCardSummary, setCardTitle, setCardVideoPlayerUrl, setCardVideoPlayerPoster
} from '../AdaptiveCard/adaptiveCard';
import { AvatarShape } from '@fluentui/react-avatar';
import { getBaseUrl } from '../../configVariables';

let card: any

enum TemplateSelection {
    Default = 'Default',
    infromational = 'infromational',
    infoVideo = 'infoVideo',
    department = 'department',
    departmentVideo = 'departmentVideo',
    allIn = "allIn"
}

export const VideoPlayer = () => {

    const { t } = useTranslation();
    const { id } = useParams() as any;
    const [loader, setLoader] = React.useState(true);

    const videoUrl = "https://www.youtube.com/embed/xxkCJKpU3vA";
    const imageLink = "https://adaptivecards.io/content/poster-video.png"
    const buttonLink = ""

    let card: any;
    
    React.useEffect(() => {

        card = getInitAdaptiveCard("Tittle", TemplateSelection.departmentVideo);
        console.log("videoUrl: " + videoUrl);
            setCardVideoPlayerUrl(card, videoUrl);
            setCardVideoPlayerPoster(card, imageLink);
            let adaptiveCard = new AdaptiveCards.AdaptiveCard();
            adaptiveCard.parse(card);
        let renderedCard = adaptiveCard.render();
        document.getElementsByClassName('card-area')[0].innerHTML = "";
         document.getElementsByClassName('card-area')[0].appendChild(renderedCard!);
            let link = buttonLink;
            adaptiveCard.onExecuteAction = function (action) { window.open(link, '_blank'); }

        
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