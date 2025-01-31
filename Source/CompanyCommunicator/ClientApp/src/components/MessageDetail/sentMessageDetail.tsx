/* eslint-disable jsx-a11y/alt-text */
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from 'react';
import { useTranslation } from 'react-i18next';
import {
  Button,
  Menu,
  MenuItem,
  MenuList,
  MenuPopover,
  MenuTrigger,
  Persona,
  Table,
  TableBody,
  TableCell,
  TableCellLayout,
  TableHeader,
  TableHeaderCell,
  TableRow,
  Tooltip,
  useArrowNavigationGroup,
} from '@fluentui/react-components';
import {
  BookExclamationMark24Regular,
  CalendarCancel24Regular,
  CheckmarkSquare24Regular,
  DocumentCopyRegular,
  Chat20Regular,
  MoreHorizontal24Filled,
  ChatMultiple24Regular,
  ShareScreenStop24Regular,
    Warning24Regular,
    Heart12Filled
} from '@fluentui/react-icons';
import * as microsoftTeams from '@microsoft/teams-js';
import { cancelSentNotification, duplicateDraftNotification } from '../../apis/messageListApi';
import { getBaseUrl } from '../../configVariables';
import { formatNumber } from '../../i18n';
import { ROUTE_PARTS, ROUTE_QUERY_PARAMS } from '../../routes';
import { GetDraftMessagesSilentAction, GetSentMessagesSilentAction } from '../../actions';
import { useAppDispatch } from '../../store';
import './sentMessageDetails.scss';
import { UrlDialogInfo } from '@microsoft/teams-js';

export const SentMessageDetail = (sentMessages: any) => {
  const { t } = useTranslation();
  const keyboardNavAttr = useArrowNavigationGroup({ axis: 'grid' });
  const dispatch = useAppDispatch();
  const statusUrl = (id: string) => getBaseUrl() + `/${ROUTE_PARTS.VIEW_STATUS}/${id}?${ROUTE_QUERY_PARAMS.LOCALE}={locale}`;

  const renderSendingText = (message: any) => {
    var text = '';
    switch (message.status) {
      case 'Queued':
        text = t('Queued');
        break;
      case 'SyncingRecipients':
        text = t('SyncingRecipients');
        break;
      case 'InstallingApp':
        text = t('InstallingApp');
        break;
      case 'Sending':
        let sentCount = (message.succeeded ? message.succeeded : 0) + (message.failed ? message.failed : 0) + (message.unknown ? message.unknown : 0);
        text = t('SendingMessages', {
          SentCount: formatNumber(sentCount),
          TotalCount: formatNumber(message.totalMessageCount),
        });
        break;
      case 'Canceling':
        text = t('Canceling');
        break;
      case 'Canceled':
      case 'Sent':
      case 'Failed':
        text = '';
    }

    return text;
  };

  const countStatusMsg = () => {
    return sentMessages?.sentMessages?.filter((x: any) => x.status && x.status !== 'Canceled' && x.status !== 'Sent' && x.status !== 'Failed').length;
  };

  const shouldNotShowCancel = (msg: any) => {
    let cancelState = false;
    if (msg !== undefined && msg.status !== undefined) {
      const status = msg.status.toUpperCase();
      cancelState = status === 'SENT' || status === 'UNKNOWN' || status === 'FAILED' || status === 'CANCELED' || status === 'CANCELING';
    }
    return cancelState;
  };

    const onOpenTaskModule = (event: any, url: string, title: string) => {
        const dialogInfo: UrlDialogInfo = {
            url,
            title,
            size: { height: microsoftTeams.DialogDimension.Large, width: microsoftTeams.DialogDimension.Large },
            fallbackUrl: url,
        };

        // now open the dialog
        microsoftTeams.dialog.url.open(dialogInfo);
    };

  const duplicateDraftMessage = async (id: number) => {
    try {
      await duplicateDraftNotification(id);
      GetDraftMessagesSilentAction(dispatch);
    } catch (error) {
      return error;
    }
  };

  const cancelSentMessage = async (id: number) => {
    try {
      await cancelSentNotification(id);
      GetSentMessagesSilentAction(dispatch);
    } catch (error) {
      return error;
    }
  };

  return (
    <>
      {sentMessages?.sentMessages && (
        <Table {...keyboardNavAttr} role='grid' className='sent-messages' aria-label='Sent messages table with grid keyboard navigation'>
          <TableHeader>
            <TableRow>
              <TableHeaderCell key='title' style={{ width: '40%' }}>
                <b>{t('TitleText')}</b>
              </TableHeaderCell>
              {countStatusMsg() > 0 && <TableHeaderCell key='status' aria-hidden='true' />}
              <TableHeaderCell key='recipients'>
                <b>{t('Recipients')}</b>
                          </TableHeaderCell>
                <TableHeaderCell key='reactions'>
                    <b>{t('Reactions')}</b>
                </TableHeaderCell>
              <TableHeaderCell key='sent'>
                <b>{t('Sent')}</b>
              </TableHeaderCell>
              <TableHeaderCell key='createdBy'>
                <b className='big-screen-visible'>{t('CreatedBy')}</b>
              </TableHeaderCell>
              <TableHeaderCell key='actions' style={{ width: '50px' }}>
                <b>Actions</b>
              </TableHeaderCell>
            </TableRow>
          </TableHeader>
          <TableBody>
            {sentMessages!.sentMessages!.map((item: any) => (
              <TableRow key={item.id + 'key'}>
                <TableCell tabIndex={0} role='gridcell'>
                  <TableCellLayout
                    truncate
                    media={<Chat20Regular />}
                    title={item.title}
                    style={{ cursor: 'pointer' }}
                    onClick={() => onOpenTaskModule(null, statusUrl(item.id), t('ViewStatus'))}
                  >
                    {item.title}
                  </TableCellLayout>
                    </TableCell>

                {countStatusMsg() > 0 && (
                  <TableCell tabIndex={0} role='gridcell'>
                    <TableCellLayout truncate>
                      <span className='big-screen-visible'>{renderSendingText(item)}</span>
                    </TableCellLayout>
                  </TableCell>
                    )}
                <TableCell tabIndex={0} role='gridcell'>
                  <TableCellLayout>
                    <div style={{ display: 'inline-block' }}>
                      <Tooltip content={t('TooltipSuccess') || ''} relationship='label'>
                        <Button
                          appearance='subtle'
                          icon={<CheckmarkSquare24Regular style={{ color: '#22bb33', verticalAlign: 'middle' }} />}
                          size='small'
                        ></Button>
                      </Tooltip>
                      <span className='recipient-text'>{formatNumber(item.succeeded)}</span>
                    </div>
                    <div style={{ display: 'inline-block' }}>
                      <Tooltip content={t('TooltipFailure') || ''} relationship='label'>
                        <Button
                          appearance='subtle'
                          icon={<ShareScreenStop24Regular style={{ color: '#bb2124', verticalAlign: 'middle' }} />}
                          size='small'
                        ></Button>
                      </Tooltip>
                      <span className='recipient-text'>{formatNumber(item.failed)}</span>
                    </div>
                    {item.canceled && (
                      <div style={{ display: 'inline-block' }}>
                        <Tooltip content='Canceled' relationship='label'>
                          <Button
                            appearance='subtle'
                            icon={<BookExclamationMark24Regular style={{ color: '#f0ad4e', verticalAlign: 'middle' }} />}
                            size='small'
                          ></Button>
                        </Tooltip>
                        <span className='recipient-text'>{formatNumber(item.canceled)}</span>
                      </div>
                    )}
                    {item.unknown && (
                      <div style={{ display: 'inline-block' }}>
                        <Tooltip content='Unknown' relationship='label'>
                          <Button
                            appearance='subtle'
                            icon={<Warning24Regular style={{ color: '#e9835e', verticalAlign: 'middle' }} />}
                            size='small'
                          ></Button>
                        </Tooltip>
                        <span className='recipient-text'>{formatNumber(item.unknown)}</span>
                      </div>
                    )}
                  </TableCellLayout>
                    </TableCell>

                    <TableCell tabIndex={0} role='gridcell'>
                        <TableCellLayout >

                        <div className="reactions">
                                <div className="reactionContainer">
                            <Tooltip content={t('TooltipLike') || ''} relationship='label'>
                                    <img src={getBaseUrl()+"/image/like.png"} className="reaction" />
                            </Tooltip>
                                <span className='recipient-text'>{item.like}</span>
                                </div>
                                <div className = "reactionContainer">
                                <Tooltip content={t('TooltipHeart') || ''} relationship='label'>
                                    <img src={getBaseUrl() + "/image/heart.png"} className="reaction" />
                                </Tooltip>
                                <span className='recipient-text'>{item.heart}</span>
                            </div>
                                <div className="reactionContainer">
                                <Tooltip content={t('TooltipLaugh') || ''} relationship='label'>
                                    <img src={getBaseUrl() + "/image/laugh.png"} className="reaction" />
                                </Tooltip>
                                <span className='recipient-text'>{item.laugh}</span>
                            </div>
                                <div className="reactionContainer">
                                <Tooltip content={t('TooltipSurpised') || ''} relationship='label'>
                                    <img src={getBaseUrl() + "/image/surprised.png"} className="reaction" />
                                </Tooltip>
                                    <span className='recipient-text'>{item.laugh}</span>
                                </div>
                                </div>
                            
                    </TableCellLayout>
                </TableCell>
                <TableCell tabIndex={0} role='gridcell'>
                  <TableCellLayout truncate>{item.sentDate}</TableCellLayout>
                </TableCell>
                <TableCell tabIndex={0} role='gridcell'>
                  <TableCellLayout truncate title={item.createdBy}>
                    <span className='big-screen-visible'>
                      <Persona
                        size='extra-small'
                        textAlignment='center'
                        name={item.createdBy}
                        secondaryText={'Member'}
                        avatar={{ color: 'colorful' }}
                      />
                    </span>
                  </TableCellLayout>
                </TableCell>
                <TableCell role='gridcell' style={{ width: '50px' }}>
                  <TableCellLayout>
                    <Menu>
                      <MenuTrigger disableButtonEnhancement>
                        <Button aria-label='Actions menu' icon={<MoreHorizontal24Filled />} />
                      </MenuTrigger>
                      <MenuPopover>
                        <MenuList>
                          <MenuItem
                            icon={<ChatMultiple24Regular />}
                            key={'viewStatusKey'}
                            onClick={() => onOpenTaskModule(null, statusUrl(item.id), t('ViewStatus'))}
                          >
                            {t('ViewStatus')}
                          </MenuItem>
                          <MenuItem key={'duplicateKey'} icon={<DocumentCopyRegular />} onClick={() => duplicateDraftMessage(item.id)}>
                            {t('Duplicate')}
                          </MenuItem>
                          {!shouldNotShowCancel(item) && (
                            <MenuItem key={'cancelKey'} icon={<CalendarCancel24Regular />} onClick={() => cancelSentMessage(item.id)}>
                              {t('Cancel')}
                            </MenuItem>
                          )}
                        </MenuList>
                      </MenuPopover>
                    </Menu>
                  </TableCellLayout>
                </TableCell>
              </TableRow>
            ))}
          </TableBody>
        </Table>
      )}
    </>
  );
};
