import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { DefaultButton, Dialog, DialogFooter, PrimaryButton } from '@fluentui/react';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import { DialogContent } from 'office-ui-fabric-react';

export interface IAddMemberDialogProps {
    hideDialog?: boolean;
    close: any;
}

export interface IIAddMemberDialogState {
    hideDialog: boolean;
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

export class AddMemberDialogContent extends React.Component<any, any> {
    constructor(props: any) {
        super(props);
    }

    public render(): React.ReactElement<any, any> {
        return <DialogContent
            title={'Add Member to Committee'}
            onDismiss={this.props.close}
            showCloseButton={true}
        >
            <div>
                <h4>Add Member Form will go here...</h4>
            </div>
            <DialogFooter>
                <PrimaryButton text="Send" />
                <DefaultButton text="Don't send" />
            </DialogFooter>
        </DialogContent>;
    }
}

