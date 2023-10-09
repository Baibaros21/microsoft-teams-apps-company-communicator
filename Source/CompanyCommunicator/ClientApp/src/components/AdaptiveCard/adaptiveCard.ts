// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as AdaptiveCards from 'adaptivecards';
import MarkdownIt from 'markdown-it';
import { getBaseUrl } from '../../configVariables';
import { TemplateSelection } from "../../store"



AdaptiveCards.AdaptiveCard.onProcessMarkdown = function (text, result) {
  result.outputHtml = new MarkdownIt().render(text);
  result.didProcess = true;
};

export const getInitAdaptiveCard = (titleText: string = "title", type: string = TemplateSelection.Default) => {

    switch (type) {
        case TemplateSelection.Default: {
            return {
                type: 'AdaptiveCard',
                body: [
                    {
                        type: 'TextBlock',
                        weight: 'Bolder',
                        text: titleText,
                        size: 'ExtraLarge',
                        wrap: true,
                        name: "title"
                    },
                    {
                        type: 'Image',
                        spacing: 'Default',
                        url: getBaseUrl() + "/image/imagePlaceholder.png",
                        altText: 'Image',
                        size: 'Auto',
                        name: "image"
                    },
                    {
                        type: 'TextBlock',
                        text: 'Author',
                        wrap: true,
                        name: "author"
                    },
                    {
                        type: 'TextBlock',
                        size: 'Small',
                        weight: 'Lighter',
                        text: 'Summary',
                        wrap: true,
                        name: "summary"
                    },
                ],
                $schema: 'https://adaptivecards.io/schemas/adaptive-card.json',
                version: '1.0',
            };


        }

        case TemplateSelection.infromational: {
            return {
                "type": "AdaptiveCard",
                "body": [
                    {
                        "type": "TextBlock",
                        "weight": "Bolder",
                        "text": "Department",
                        "size": "medium",
                        "wrap": true,
                        "horizontalAlignment": "Center",
                        name: "department"
                    },
                    {
                        "type": "TextBlock",
                        "weight": "Bolder",
                        "spacing": "None",
                        "text": titleText,
                        "size": "ExtraLarge",
                        "wrap": true,
                        "horizontalAlignment": "Center",
                        name: "title"
                    },
                    {
                        "type": "Image",
                        "spacing": "Default",
                        "url": getBaseUrl() + "/image/imagePlaceholder.png",
                        "size": "Stretch",
                        "width": "Auto",
                        "altText": "",
                        name: "image"
                    },
                    {
                        "type": "TextBlock",
                        "text": "Author",
                        "wrap": true,
                        "horizontalAlignment": "Left",
                        name: "author"
                    },
                    {
                        "type": "TextBlock",
                        "wrap": true,
                        "size": "Small",
                        "weight": "Lighter",
                        "text": "summary",
                        name: "summary"
                    },
                    {
                        "type": "ColumnSet",
                        "columns": [
                            {
                                "type": "Column",
                                "width": "stretch",
                                "items": [
                                    {
                                        "type": "Image",
                                        "url": getBaseUrl() + "/image/UAE.png",
                                        "size": "Large"
                                    }
                                ],
                                "verticalContentAlignment": "Center"
                            },
                            {
                                "type": "Column",
                                "width": "stretch",
                                "items": [
                                    {
                                        "type": "Image",
                                        "url": getBaseUrl() + "/image/Customs.png"
                                    }
                                ],
                                "verticalContentAlignment": "Bottom"
                            }
                        ]
                    }
                ],
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "version": "1.2"
            };

        }
        case TemplateSelection.infoVideo: {

            return {
                "type": "AdaptiveCard",
                "body": [
                    {
                        "type": "TextBlock",
                        "weight": "Bolder",
                        "text": "Department",
                        "size": "medium",
                        "wrap": true,
                        "horizontalAlignment": "Center",
                        "name": "Department"
                    },
                    {
                        "type": "TextBlock",
                        "weight": "Bolder",
                        "spacing": "None",
                        "text": titleText,
                        "size": "ExtraLarge",
                        "wrap": true,
                        "horizontalAlignment": "Center",
                        "name": "title"
                    },
                    {
                        "type": "Media",
                        "poster": getBaseUrl() + "/image/imagePlaceholder.png",

                        name: "Video",
                        "sources": [
                            {
                                "mimeType": "video/mp4",
                                "url": ""
                            }
                        ]
                    },
                    {
                        "type": "TextBlock",
                        "text": "Author",
                        "wrap": true,
                        "horizontalAlignment": "Left",
                        name: "author"
                    },
                    {
                        "type": "TextBlock",
                        "wrap": true,
                        "size": "Small",
                        "weight": "Lighter",
                        "text": "Summary",
                        name: "summary"
                    },
                    {
                        "type": "ColumnSet",
                        "columns": [
                            {
                                "type": "Column",
                                "width": "stretch",
                                "items": [
                                    {
                                        "type": "Image",
                                        "url": getBaseUrl() + "/image/UAE.png",
                                        "size": "Large"
                                    }
                                ],
                                "verticalContentAlignment": "Center"
                            },
                            {
                                "type": "Column",
                                "width": "stretch",
                                "items": [
                                    {
                                        "type": "Image",
                                        "url": getBaseUrl() + "/image/Customs.png"
                                    }
                                ],
                                "verticalContentAlignment": "Bottom"
                            }
                        ]
                    }
                ],
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "version": "1.2"
            };

        }
        case TemplateSelection.departmentVideo: {
            return {
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "type": "AdaptiveCard",
                "version": "1.5",
                "fallbackText": "This card requires CaptionSource to be viewed. Ask your platform to update to Adaptive Cards v1.6 for this and more!",
                "body": [
                    {
                        "type": "TextBlock",
                        "weight": "Bolder",
                        "text": titleText,
                        "size": "medium",
                        "wrap": true,
                        "horizontalAlignment": "Center",
                        name: "department"
                    },
                    {
                        "type": "Media",
                        "poster": "https://adaptivecards.io/content/poster-video.png",
                        name: "Video",
                        "sources": [
                            {
                                "mimeType": "video/mp4",
                                "url": ""
                            }
                        ]

                    }
                ]
            };
        }
        case TemplateSelection.department: {

            return {
                "type": "AdaptiveCard",
                "body": [
                    {
                        "type": "TextBlock",
                        "weight": "Bolder",
                        "text": "department",
                        "size": "medium",
                        "wrap": true,
                        "horizontalAlignment": "Center",
                        name: "department"
                    },
                    {
                        "type": "TextBlock",
                        "weight": "Bolder",
                        "spacing": "None",
                        "text": titleText,
                        "size": "ExtraLarge",
                        "altText": '',
                        "wrap": true,
                        "horizontalAlignment": "Center",
                        name: "title"
                    },

                    {
                        "type": "TextBlock",
                        "text": "Author",
                        "wrap": true,
                        "horizontalAlignment": "left",
                        name: 'author'
                    },
                    {
                        "type": "TextBlock",
                        "wrap": true,
                        "size": "Small",
                        "weight": "Lighter",
                        "text": "summary",
                        "horizontalAlignment": "left",
                        name: "summary"
                    },
                    {
                        "type": "ColumnSet",
                        "columns": [
                            {
                                "type": "Column",
                                "width": "stretch",
                                "items": [
                                    {
                                        "type": "Image",
                                        "url": getBaseUrl() + "/image/UAE.png",
                                        "size": "Large"
                                    }
                                ],
                                "verticalContentAlignment": "Center"
                            },
                            {
                                "type": "Column",
                                "width": "stretch",
                                "items": [
                                    {
                                        "type": "Image",
                                        "url": getBaseUrl() + "/image/Customs.png"
                                    }
                                ],
                                "verticalContentAlignment": "Bottom"
                            }
                        ]
                    }
                ],
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "version": "1.2"
            };

        }
        case TemplateSelection.allIn: {

            return {
                "type": "AdaptiveCard",
                "body": [
                    {
                        "type": "TextBlock",
                        "weight": "Bolder",
                        "text": "department",
                        "size": "medium",
                        "wrap": true,
                        "horizontalAlignment": "Center",
                        name: "department"
                    },
                    {
                        "type": "TextBlock",
                        "weight": "Bolder",
                        "spacing": "None",
                        "text": titleText,
                        "size": "ExtraLarge",
                        "wrap": true,
                        "horizontalAlignment": "Center",
                        name: "title"
                    },
                    {
                        "type": "Image",
                        "spacing": "Default",
                        "url": getBaseUrl() + "/image/imagePlaceholder.png",
                        "size": "Stretch",
                        "width": "Auto",
                        "altText": "",
                        name: "image"
                    },

                    {
                        "type": "Media",
                        "poster": getBaseUrl() + "/image/imagePlaceholder.png",
                        name: "video",
                        "sources": [
                            {
                                "mimeType": "video/mp4",
                                "url": "",
                            }
                        ]
                    },
                    {
                        "type": "TextBlock",
                        "text": "Author",
                        "wrap": true,
                        "horizontalAlignment": "Right",
                        name: "author"
                    },
                    {
                        "type": "TextBlock",
                        "wrap": true,
                        "size": "Small",
                        "weight": "Lighter",
                        "text": "Summary",
                        "horizontalAlignment": "Right",
                        name: "summary"
                    },
                    {
                        "type": "ColumnSet",
                        "columns": [
                            {
                                "type": "Column",
                                "width": "stretch",
                                "items": [
                                    {
                                        "type": "Image",
                                        "url": getBaseUrl() + "/image/UAE.png",
                                        "size": "Large"
                                    }
                                ],
                                "verticalContentAlignment": "Center"
                            },
                            {
                                "type": "Column",
                                "width": "stretch",
                                "items": [
                                    {
                                        "type": "Image",
                                        "url": getBaseUrl() + "/image/Customs.png"
                                    }
                                ],
                                "verticalContentAlignment": "Bottom"
                            }
                        ]
                    }
                ],
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "version": "1.2"
            };

        }


    };
}



    var getProperty = (body: any, value: string): any => {
        return body.filter((prop: any) =>
            prop.name === value
        );
    }

    export const getCardTitle = (card: any) => {

        var filteredcomp = getProperty(card.body, "title");
        if (filteredcomp.length > 0) {
            return filteredcomp.text
        }
    };

    export const setCardTitle = (card: any, title: string) => {
        var filteredcomp = getProperty(card.body, "title");
        console.log("In title");
        if (filteredcomp.length > 0) {

            filteredcomp[0].text = title

        }

    };


    export const getDeptTitle = (card: any) => {

        var filteredcomp = getProperty(card.body, "department");
        if (filteredcomp.length > 0) {
            return filteredcomp[0].text
        }
    };

export const setCardDeptTitle = (card: any, title?: string) => {
        var filteredcomp = getProperty(card.body, "department");
        if (filteredcomp.length > 0) {

            filteredcomp[0].text = title

        }

    };

    export const getCardImageLink = (card: any) => {
        const filteredcomp = getProperty(card.body, "image");
        if (filteredcomp.length > 0) {
            return filteredcomp[0].url
        }
    };

    export const setCardImageLink = (card: any, imageLink?: string) => {
        const filteredcomp = getProperty(card.body, "image");
        if (filteredcomp.length > 0) {
            filteredcomp[0].url = imageLink
        }
    };
    export const setCardVideoPlayerUrl = (card: any, videoLink?: string) => {

        const filteredcomp = getProperty(card.body, "video");
        if (filteredcomp.length > 0) {
            filteredcomp[0].sources[0].url = videoLink;
        }
    };
    export const setCardVideoPlayerPoster = (card: any, imageLink?: string) => {

        const filteredcomp = getProperty(card.body, "video");
        if (filteredcomp.length > 0) {
            filteredcomp[0].poster = imageLink;
        }
    }

    export const getCardSummary = (card: any) => {
        const filteredcomp = getProperty(card.body, "summary");
        if (filteredcomp.length > 0) {
            return filteredcomp[0].text
        }
    };

    export const setCardSummary = (card: any, summary?: string) => {
        const filteredcomp = getProperty(card.body, "summary");
        if (filteredcomp.length > 0) {
            filteredcomp[0].text = summary
        }
    };

    export const getCardAuthor = (card: any) => {
        const filteredcomp = getProperty(card.body, "author");
        if (filteredcomp.length > 0) {
            return filteredcomp[0].text
        }
    };

    export const setCardAuthor = (card: any, author?: string) => {
        const filteredcomp = getProperty(card.body, "author");
        if (filteredcomp.length > 0) {
            filteredcomp[0].text = author
        }
    };

    export const getCardBtnTitle = (card: any) => {
        return card.actions[0].title;
    };

    export const getCardBtnLink = (card: any) => {
        return card.actions[0].url;
    };

    export const setCardBtn = (card: any, buttonTitle?: string, buttonLink?: string) => {
        if (buttonTitle && buttonLink) {
            card.actions = [
                {
                    type: 'Action.OpenUrl',
                    title: buttonTitle,
                    url: buttonLink,
                },
            ];
        } else {
            delete card.actions;
        }
    };
