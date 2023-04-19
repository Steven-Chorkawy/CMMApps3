import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { Panel, PanelType } from 'office-ui-fabric-react';
import { BaseDialog } from '@microsoft/sp-dialog';
import { CommitteeMemberDashboard } from './CommitteeMemberDashboard';

export class CommitteeMemberDashboardPanel extends BaseDialog {
    private _context: any = null;
    private _memberId: number = null;
    constructor(props: any) {
        super(props);
        this._context = props.context;
        this._memberId = props.memberId;
    }
    public render(): void {
        debugger;
        ReactDOM.render(<Panel
            isLightDismiss={false}
            isOpen={true}
            type={PanelType.large}
            onDismissed={() => this.close()}
        >
            <CommitteeMemberDashboard memberId={this._memberId} context={this._context} />
        </Panel>, this.domElement);
    }
}

