import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Panel, PanelType } from 'office-ui-fabric-react';
import * as React from 'react';
import IMemberListItem from '../ClaringtonInterfaces/IMemberListItem';
import { GetMember } from '../HelperMethods/MyHelperMethods';
import { MyShimmer } from './MyShimmer';

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
        if (this.state.selectedMember) {
            return (
                <div>
                    <div>hello world... Member ID: {this.props.memberId}</div>
                    <div>
                        {this.state.selectedMember && JSON.stringify(this.state.selectedMember)}
                    </div>
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