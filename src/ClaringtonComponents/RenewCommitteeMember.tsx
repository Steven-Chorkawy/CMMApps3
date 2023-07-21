import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Panel, PanelType } from 'office-ui-fabric-react';
import * as React from 'react';
import IMemberListItem from '../ClaringtonInterfaces/IMemberListItem';
import { GetMember, OnFormatDate } from '../HelperMethods/MyHelperMethods';
import { MyShimmer } from './MyShimmer';
import { Dashboard, WidgetSize } from '@pnp/spfx-controls-react/lib/Dashboard';
import { CommitteeMemberContactDetails, CommitteeMemberTermHistory } from './MemberDetailsComponent';
import { RenewMemberComponent } from './RenewMemberComponent';

export interface IRenewCommitteeMemberProps {
    description?: string;
    memberId: number;
    context: WebPartContext;
}

export interface IRenewCommitteeMemberState {
    selectedMember?: IMemberListItem;
}

export class RenewCommitteeMember extends React.Component<IRenewCommitteeMemberProps, IRenewCommitteeMemberState> {

    constructor(props: any) {
        super(props);
        this.state = {
            selectedMember: undefined
        };

        if (this.props.memberId) {
            GetMember(this.props.memberId).then(member => this.setState({ selectedMember: member }));
        }
    }

    public render(): React.ReactElement<any> {
        if (this.state.selectedMember) {
            return (
                <Dashboard
                    widgets={[
                        {
                            title: this.state.selectedMember.Title,
                            desc: `Last updated ${OnFormatDate(new Date(this.state.selectedMember.Modified))}`,
                            size: WidgetSize.Box,
                            body: [
                                {
                                    id: "t1",
                                    title: "Tab 1",
                                    content: (
                                        <div style={{ overflow: 'scroll' }}>
                                            <CommitteeMemberContactDetails member={this.state.selectedMember} />
                                            <hr />
                                            <CommitteeMemberTermHistory memberID={this.state.selectedMember.ID} context={this.props.context} />
                                        </div>
                                    ),
                                }
                            ]
                        },
                        {
                            title: "Renew Committee Member",
                            size: WidgetSize.Box,
                            body: [{
                                id: 'renewMemberId',
                                title: 'Renew Member',
                                content: (<RenewMemberComponent context={this.props.context} />)
                            }]
                        }
                    ]} />
            );
        }
        else {
            return <MyShimmer />;
        }
    }
}

export class RenewCommitteeMemberPanel extends React.Component<any, any> {
    private _context: any = null;
    private _memberId: number = null;

    constructor(props: any) {
        super(props);
        this._context = props.context;
        this._memberId = props.memberId;
        this.state = {
            isOpen: true
        }
    }
    public render(): React.ReactElement<any, any> {
        return <Panel
            isLightDismiss={false}
            isOpen={this.state.isOpen}
            type={PanelType.large}
            allowTouchBodyScroll={false}
            onDismiss={() => this.setState({ isOpen: false })}
        >
            <RenewCommitteeMember memberId={this._memberId} context={this._context} />
        </Panel>;
    }
}