﻿import * as AdaptiveCards from 'adaptivecards';
import * as React from 'react';
import { TemplateSelection, useAppSelector, RootState, useAppDispatch } from "../../store";
import { useTranslation } from 'react-i18next';
import {
    Button,
    Field,

    makeStyles,
    tokens,
    Label,

    Radio,
    RadioGroup,
    Spinner,
    RadioGroupOnChangeData
} from '@fluentui/react-components';
import * as microsoftTeams from '@microsoft/teams-js';
import { getDefaultData, } from '../../apis/messageListApi';
import { setCardLogo, setCardBanner, saveAdaptiveCard } from '../AdaptiveCard/adaptiveCard';
import AceEditor from 'react-ace';
import * as ACData from 'adaptivecards-templating';
import { GetAllCardTemplatesAction } from "../../actions";
import 'brace/mode/javascript';
import 'brace/mode/json';

import 'brace/theme/monokai';
const validPropNames = ['title', 'logo', 'banner', 'department', 'summary', 'author', 'image', 'video'];
enum CurrentPageSelection {
    TemplateChoice = "TemplateChoice",
    JsonEditor = "JsonEditor"
}

interface IDefaults {

    logoFileName: string;
    logoLink: string;
    bannerFileName: string;
    bannerLink: string;
}
const useFieldStyles = makeStyles({
    styles: {
        marginBottom: tokens.spacingVerticalM,
        gridGap: tokens.spacingHorizontalXXS,
    },
});

let card: any;
export const ModifyTemplatesTask = () => {
    const field_styles = useFieldStyles();
    const { t } = useTranslation();
    const dispatch = useAppDispatch();
    const Templates: any = useAppSelector((state: RootState) => state.messages).cardTemplates.payload; const [pageSelection, setPageSelection] = React.useState(CurrentPageSelection.TemplateChoice);
    const [cardAreaBorderClass, setCardAreaBorderClass] = React.useState('card-area-border');
    const [json, setJson] = React.useState<any>();
    const [selectedTemplate, setSelectedTemplate] = React.useState(TemplateSelection.Default);
    const [showMsgDraftingSpinner, setShowMsgDraftingSpinner] = React.useState(false);
    const [isBackBtnDisabled, setIsBackBtnDisabled] = React.useState(false);
    const [isSaveBtnDisabled, setIsSaveBtnDisabled] = React.useState(false);
    const [jsonErrorMessage, setjsonErrorMessage] = React.useState('');
    const [defaultsState, setDefaultState] = React.useState<IDefaults>({
        logoFileName: "",
        logoLink: "",
        bannerLink: "",
        bannerFileName: ""
    });
    React.useEffect(() => {
        getDefaultsItem();
        GetAllCardTemplatesAction(dispatch);
    }, []);

    React.useEffect(() => {
        if (Templates && Templates.length > 0) {
            getCurrentCardTemplate(selectedTemplate);
        }

    }, [Templates]);
    React.useEffect(() => {

        if (pageSelection === CurrentPageSelection.JsonEditor) {
            setJson(JSON.stringify(card, null, 2));
            updateAdaptiveCard();
        } else if (pageSelection === CurrentPageSelection.TemplateChoice) {
            if (Templates && Templates.length > 0) {
                getCurrentCardTemplate(selectedTemplate);
            }
        }
    }, [pageSelection]);

    const getCurrentCardTemplate = (cardtemplate: TemplateSelection) => {
        console.log(Templates);
        card = Templates?.find((template: any) => template.name === cardtemplate)?.card;

        console.log(card);
        var cardTemplate = new ACData.Template(JSON.parse(card));
        card = cardTemplate.expand({
            $root: {


            }
        });

        setCardLogo(card, defaultsState.logoLink);
        setCardBanner(card, defaultsState.bannerLink);
        updateAdaptiveCard();

    }


    const getDefaultsItem = async () => {

        try {
            await getDefaultData().then((response) => {

                const defaultImages = response.data;
                console.log(defaultImages);
                setDefaultState({
                    logoFileName: defaultImages.logoFileName,
                    logoLink: defaultImages.logoLink,
                    bannerFileName: defaultImages.bannerFileName,
                    bannerLink: defaultImages.bannerLink
                });
            });

        }
        catch (error) {

        }

    }




    const validateNameExists = (json: any): boolean => {
        var valid = true;
        json.forEach((prop: any) => {
            try {
                console.log(prop.name);
                if (prop.name === null) {
                    setjsonErrorMessage("Error! Name property canoot be null")
                    valid = false;
                    return;
                } else if (prop.name === "") {
                    setjsonErrorMessage("Error! Name property cannot be empty")
                    valid = false;
                    return;

                } else if (!validPropNames.includes(prop.name)) {
                    setjsonErrorMessage("Error! Name should be one of the following: " + validPropNames)
                    valid = false;
                    return;
                }
            } catch {
                valid = false
                return;
            }
        });
        console.log(valid);
        return valid
    }
    const onJSONChange = (event: any) => {
        console.log("change");
        setjsonErrorMessage("");
        try {
            card = JSON.parse(event);
            setJson(event);
            if (validateNameExists(card.body)) {
                setIsSaveBtnDisabled(false);
                updateAdaptiveCard()
            } else {
                setIsSaveBtnDisabled(true);
            }

        } catch (error) {
            // Handle error
            console.log(error);
            setJson(event);
            setjsonErrorMessage("Error!: Invalide Json for an adaptive card");
        }


    };

    const templateSelectionChange = (ev: any, data: RadioGroupOnChangeData) => {
        let input = data.value as keyof typeof TemplateSelection;
        setSelectedTemplate(TemplateSelection[input]);

        getCurrentCardTemplate(TemplateSelection[input]);
    };

    const updateAdaptiveCard = () => {
        var adaptiveCard = new AdaptiveCards.AdaptiveCard();
        adaptiveCard.parse(card);
        const renderCard = adaptiveCard.render();

        if (renderCard && pageSelection === CurrentPageSelection.TemplateChoice) {
            document.getElementsByClassName('card-area-1')[0].innerHTML = '';
            document.getElementsByClassName('card-area-1')[0].appendChild(renderCard);
            setCardAreaBorderClass('card-area-border');
        } else if (renderCard && pageSelection === CurrentPageSelection.JsonEditor) {
            document.getElementsByClassName('card-area-2')[0].innerHTML = '';
            document.getElementsByClassName('card-area-2')[0].appendChild(renderCard);
            setCardAreaBorderClass('card-area-border');
        }
        adaptiveCard.onExecuteAction = function (action: any) {
            window.open(action.url, '_blank');
        };
    };

    const onNext = (event: any) => {
        switch (pageSelection) {
            case (CurrentPageSelection.TemplateChoice):
                setPageSelection(CurrentPageSelection.JsonEditor);
                setIsBackBtnDisabled(false);
                break;
            default:


        }

    };

    const onBack = (event: any) => {
        switch (pageSelection) {
            case (CurrentPageSelection.JsonEditor):
                setPageSelection(CurrentPageSelection.TemplateChoice);
                setIsBackBtnDisabled(true);
                break;
            default:


        }

    };

    const onSave = (event: any): void => {
        setShowMsgDraftingSpinner(true);
        saveAdaptiveCard(card, selectedTemplate).then(() => {
            microsoftTeams.tasks.submitTask();
        }).finally(() => {
            setShowMsgDraftingSpinner(false);
        })
    }

    return (
        <>
            {pageSelection === CurrentPageSelection.TemplateChoice && (
                <>
                    <span role='alert' aria-label={t('NewMessageStep2')} />
                    <div className='adaptive-task-grid'>
                        <div className='form-area'>
                            <Label size='large' id='TemplateSelectionGroupLabelId'>
                                {t('SendHeadingText')}
                            </Label>
                            <RadioGroup defaultValue={selectedTemplate} aria-labelledby='TemplateSelectionGroupLabelId' onChange={templateSelectionChange}>
                                <Radio id='radio1' value="Default" label={TemplateSelection.Default} />

                                <Radio id='radio2' value="infromational" label={TemplateSelection.infromational} />

                                <Radio id='radio4' value="department" label={TemplateSelection.department} />

                                <Radio id='radio5' value="departmentVideo" label={TemplateSelection.departmentVideo} />

                                <Radio id='radio7' value="Default_ar" label={TemplateSelection.Default_ar} />

                                <Radio id='radio10' value="department_ar" label={TemplateSelection.department_ar} />

                                <Radio id='radio11' value="departmentVideo_ar" label={TemplateSelection.departmentVideo_ar} />

                                <Radio id='radio12' value="uae50" label={TemplateSelection.uae50} />

                            </RadioGroup>
                        </div>
                        <div className='card-area'>
                            <div className={cardAreaBorderClass}>
                                <div className='card-area-1'></div>
                            </div>
                        </div>
                    </div>
                    <div>
                        <div className='fixed-footer'>
                            <div className='footer-action-right'>
                                <div className='footer-actions-flex'>
                                    <Button
                                        style={{ marginLeft: '16px' }}
                                        disabled={isBackBtnDisabled}
                                        id='backBtn'
                                        onClick={onBack}
                                        appearance='primary'
                                    >
                                        {t('Back')}
                                    </Button>
                                    {showMsgDraftingSpinner && (
                                        <Spinner
                                            role='alert'
                                            id='draftingLoader'
                                            size='small'
                                            label={t('DraftingMessageLabel')}
                                            labelPosition='after'
                                        />
                                    )}
                                    <Button
                                        style={{ marginLeft: '16px' }}

                                        id='saveBtn'
                                        onClick={onNext}
                                        appearance='primary'
                                    >
                                        {t('ModifyTemplate')}
                                    </Button>
                                </div>
                            </div>
                        </div>
                    </div>
                </>
            )}

            {pageSelection === CurrentPageSelection.JsonEditor && (
                <>
                    <div className='adaptive-task-grid'>


                        <AceEditor
                            mode="json"
                            theme="monokai"
                            className='form-area'
                            fontSize={12}
                            width="700"
                            showPrintMargin={true}
                            showGutter={true}
                            highlightActiveLine={true}
                            onChange={onJSONChange}

                            name="UNIQUE_ID_OF_DIV"
                            editorProps={{ $blockScrolling: true, wrapEnabled: true }}
                            setOptions={
                                {
                                    enableBasicAutocompletion: false,
                                    enableLiveAutocompletion: false,
                                    enableSnippets: false,
                                    showLineNumbers: true,
                                    tabSize: 2,
                                }}
                            value={json}
                        />

                        <div className='card-area'>
                            <Field
                                size='large'
                                className={field_styles.styles}
                                label={t("modifyTemplate")}
                                validationMessage={jsonErrorMessage}

                            >
                                <div className={cardAreaBorderClass}>

                                    <div className='card-area-2'></div>

                                </div>
                            </Field>
                        </div>
                    </div>
                    <div>
                        <div className='fixed-footer'>
                            <div className='footer-action-right'>
                                <div className='footer-actions-flex'>
                                    <Button
                                        style={{ marginLeft: '16px' }}
                                        disabled={isBackBtnDisabled}
                                        id='saveBtn'
                                        onClick={onBack}
                                        appearance='primary'
                                    >
                                        {t('Back')}
                                    </Button>
                                    {showMsgDraftingSpinner && (
                                        <Spinner

                                            role='alert'
                                            id='draftingLoader'
                                            size='small'
                                            label={t('DraftingMessageLabel')}
                                            labelPosition='after'
                                        />
                                    )}
                                    <Button
                                        style={{ marginLeft: '16px' }}
                                        id='saveBtn'
                                        onClick={onSave}
                                        appearance='primary'
                                        disabled={isSaveBtnDisabled || showMsgDraftingSpinner}
                                    >
                                        {t('Save')}
                                    </Button>
                                </div>
                            </div>
                        </div>
                    </div>
                </>
            )}



        </>
    );
}