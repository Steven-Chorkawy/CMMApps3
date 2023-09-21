import * as React from 'react';
import { Panel, PanelType } from 'office-ui-fabric-react';
import { CommitteeMemberDashboard } from './CommitteeMemberDashboard';

export class CommitteeMemberDashboardPanel extends React.Component<any, any> {
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
            <CommitteeMemberDashboard memberId={this._memberId} context={this._context} />
        </Panel>;
    }
}

