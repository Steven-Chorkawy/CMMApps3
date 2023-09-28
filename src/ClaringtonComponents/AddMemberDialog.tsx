import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { DefaultButton, DialogFooter, Panel, PanelType, PrimaryButton } from '@fluentui/react';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import { DialogContent } from 'office-ui-fabric-react';
import { CommitteeMemberDashboard } from './CommitteeMemberDashboard';

export interface IAddMemberDialogProps {
    hideDialog?: boolean;
    close: any;
}

export interface IIAddMemberDialogState {
    hideDialog: boolean;
}

export class AddMemberDialogContent extends React.Component<any, any> {
    constructor(props: any) {
        super(props);
        console.log(this.props.context);
        debugger;
    }

    public render(): React.ReactElement<any, any> {
        return <DialogContent
            title={'Add Member to Committee'}
            onDismiss={this.props.close}
            showCloseButton={true}
        >
            <div>
                <h3>TESTING...</h3>
                <CommitteeMemberDashboard context={this.props.context} />
            </div>
            <DialogFooter>
                <PrimaryButton text="Send" />
                <DefaultButton text="Don't send" />
            </DialogFooter>
        </DialogContent>;
    }
}

export class AddMemberDialogBase extends BaseDialog {
    public message: string;

    public render(): void {
        ReactDOM.render(<AddMemberDialogContent close={this.close} />, this.domElement);
    }

    public getConfig(): IDialogConfiguration {
        return { isBlocking: false };
    }

    protected onAfterClose(): void {
        super.onAfterClose();
        // Clean up the element for the next dialog
        ReactDOM.unmountComponentAtNode(this.domElement);
    }
}


export class AddMemberPanel extends React.Component<any, any> {
    /**
     *
     */
    constructor(props: any) {
        super(props);
        this.state = {
            isOpen: this.props.isOpen ? this.props.isOpen : false
        }
    }

    public render(): React.ReactElement<any, any> {
        return <Panel isLightDismiss={false} isOpen={this.state.isOpen} type={PanelType.large} onDismiss={() => this.setState({ isOpen: false })}>
            <h3>Add member Panel</h3>
        </Panel>;
    }
}