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
    context?: WebPartContext;
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
        const LINK_MEMBER_EDIT_FORM = { href: `/sites/CMM/Lists/Members/EditForm.aspx?ID=${this.props.memberId}` };
        const LINK_MEMBER_HISTORY = { href: `/sites/CMM/Lists/CommitteeMemberHistory?FilterField1=MemberID&FilterValue1=${this.props.memberId}&sortField=_EndDate&isAscending=false` };
        const LINK_EXAMPLE = { href: `#` };
        if (this.state.selectedMember) {
            return (
                <div>
                    <div>
                        <Dashboard
                            widgets={[
                                {
                                    title: "Renew Committee Member",
                                    size: WidgetSize.Triple,
                                    body: [{
                                        id: 'renewMemberId',
                                        title: 'Renew Member',
                                        content: (<RenewMemberComponent />)
                                    }]
                                },
                                {
                                    title: this.state.selectedMember.Title,
                                    desc: `Last updated ${OnFormatDate(new Date(this.state.selectedMember.Modified))}`,
                                    size: WidgetSize.Single,
                                    body: [
                                        {
                                            id: "t1",
                                            title: "Tab 1",
                                            content: (
                                                <CommitteeMemberContactDetails member={this.state.selectedMember} />
                                            ),
                                        }
                                    ],
                                    link: LINK_MEMBER_EDIT_FORM,
                                },
                                // {
                                //     title: "Committee History",
                                //     size: WidgetSize.Single,
                                //     body: [{
                                //         id: 'id',
                                //         title: 'Committee History',
                                //         content: (<CommitteeMemberTermHistory memberID={this.state.selectedMember.ID} context={this.props.context} />)
                                //     }],
                                //     link: LINK_MEMBER_HISTORY,
                                // },

                            ]} />
                    </div>
                    {/* <hr />
                    <div>
                        {this.state.selectedMember && JSON.stringify(this.state.selectedMember)}
                    </div> */}
                </div>
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
            onDismiss={() => this.setState({ isOpen: false })}
        >
            <RenewCommitteeMember memberId={this._memberId} context={this._context} />
        </Panel>;
    }
}