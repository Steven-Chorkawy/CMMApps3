import * as React from 'react';
import { DefaultButton, IStackTokens, Panel, PanelType, PrimaryButton, Stack } from '@fluentui/react';
import { CommitteeMemberDashboard } from './CommitteeMemberDashboard';
import { NewCommitteeMemberFormComponent } from './NewCommitteeMemberFormComponent';
import { FieldArray, Form, FormElement, FormRenderProps } from '@progress/kendo-react-form';
import { MyShimmer } from './MyShimmer';
import { CreateNewCommitteeMember, GetListOfActiveCommittees } from '../HelperMethods/MyHelperMethods';
import { ListViewCommandSetContext } from '@microsoft/sp-listview-extensibility';


export interface IAddMemberPanelState {
    activeCommittees?: any[];
    isOpen: boolean;
    failedToLoadActiveCommittees: boolean;
    selectedMember?: any;
}

export interface IAddMemberPanelProps {
    context: ListViewCommandSetContext;
}

export class AddMemberPanel extends React.Component<IAddMemberPanelProps, IAddMemberPanelState> {
    constructor(props: any) {
        super(props);
        console.log('AddMemberPanel Props:');
        console.log(props);
        this.state = {
            isOpen: true,
            failedToLoadActiveCommittees: false
        };

        GetListOfActiveCommittees()
            .then(value => {
                this.setState({ activeCommittees: value });
            })
            .catch(reason => {
                console.error('1: Something went wrong while getting list of active committees!');
                console.error(reason);
                console.log('Attempting to reload list of active committess.');

                // Call the same method again. But if this method fails do not call it again.
                GetListOfActiveCommittees()
                    .then(value => {
                        this.setState({ activeCommittees: value });
                    })
                    .catch(reason => {
                        console.error('2: Something went wrong while getting list of active committees!');
                        console.error(reason);
                        console.log('Will not attempt to reload the list of active committees.');
                        this.setState({ failedToLoadActiveCommittees: true });
                    });
            });
    }

    private _onSubmit = async (values: any): Promise<void> => {
        alert('form submit');
        console.log(values);

        for (let committeeIndex = 0; committeeIndex < values.Committees.length; committeeIndex++) {
            const currentCommittee = values.Committees[committeeIndex];
            await CreateNewCommitteeMember(values.selectedMember.ID, currentCommittee)
                .then(value => {
                    alert('Success! New member has been added.');
                    this.setState({ isOpen: false });
                })
                .catch(reason => {
                    alert('Failed to save new committee member!');
                    console.error(reason);
                })
        }
    }

    private _buttons(): React.ReactElement<any, any> {
        const stackTokens: IStackTokens = { childrenGap: 40 };
        return <div>
            <Stack horizontal tokens={stackTokens}>
                <PrimaryButton text="Add Member" type="submit" allowDisabledFocus disabled={this.state.selectedMember ? false : true} />
                <DefaultButton text="Cancel" onClick={() => this.setState({ isOpen: false })} allowDisabledFocus disabled={false} />
            </Stack >
        </div >;
    }

    public render(): React.ReactElement<any, any> {
        return <Panel
            isLightDismiss={false}
            isOpen={this.state.isOpen}
            type={PanelType.large}
            onDismiss={() => this.setState({ isOpen: false })}
        >
            <Form
                onSubmit={this._onSubmit}
                render={(formRenderProps: FormRenderProps) => (
                    <FormElement>
                        <h2>Add Member</h2>
                        {this._buttons()}
                        <CommitteeMemberDashboard
                            context={this.props.context}
                            selectMemberCallback={(member: any) => {
                                console.log(member);
                                this.setState({ selectedMember: member });
                                formRenderProps.onChange('selectedMember', { value: member });
                            }}
                        />
                        {
                            (this.state.activeCommittees) && (this.state.activeCommittees.length > 0 && this.state.failedToLoadActiveCommittees === false) ?
                                <div>
                                    <h2>Add '{this.state.selectedMember?.FirstName} {this.state.selectedMember?.LastName}' to Committee</h2>
                                    <FieldArray
                                        name={'Committees'}
                                        component={NewCommitteeMemberFormComponent}
                                        context={this.props.context}
                                        activeCommittees={this.state.activeCommittees}
                                        formRenderProps={formRenderProps}
                                    />
                                </div> :
                                <div>
                                    <MyShimmer />
                                </div>
                        }
                        {this._buttons()}
                    </FormElement>
                )}
            />
        </Panel >;
    }
}