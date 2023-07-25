import { ComboBox, DatePicker, DefaultButton, IComboBoxOption, PrimaryButton } from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import * as React from 'react';
import { CONSOLE_LOG_ERROR, CalculateTermEndDate, GetChoiceColumn, GetCommitteeByName, GetListOfActiveCommittees, OnFormatDate, RenewCommitteeMember, getSP } from '../HelperMethods/MyHelperMethods';
import { MyShimmer } from './MyShimmer';
import { Field, Form, FormElement, FormRenderProps } from '@progress/kendo-react-form';
import { MyComboBox, MyDatePicker } from './MyFormComponents';
import ICommitteeFileItem from '../ClaringtonInterfaces/ICommitteeFileItem';
import { FilePicker, IFilePickerResult } from '@pnp/spfx-controls-react';

export interface IRenewMemberComponentProps {
    context: WebPartContext;
    memberId: number;
}

export interface IRenewMemberComponentState {
    activeCommittees?: IComboBoxOption[];
    statusOptions?: any[];
    positionOptions?: any[];
    committeeFileItem?: ICommitteeFileItem;
    selectedStartDate?: Date;
    calculatedEndDate?: Date;
}

export class RenewMemberComponent extends React.Component<IRenewMemberComponentProps, IRenewMemberComponentState> {
    constructor(props: any) {
        super(props);

        this.state = {
            activeCommittees: [],
            committeeFileItem: undefined
        };

        GetListOfActiveCommittees().then(committees => {
            let committeeMapHolder: IComboBoxOption[] = [];
            committees.map((committee: any) => {
                committeeMapHolder.push({ key: committee.Title, text: committee.Title });
            });
            // TODO: Delete this line.
            committeeMapHolder.push({ key: 'sample', text: 'test_DELETE_ME' });
            this.setState({ activeCommittees: committeeMapHolder });
        });
    }

    private _onSubmit = async (values: any): Promise<void> => {
        console.log('_onSubmit');
        console.log(values);
        RenewCommitteeMember();
    }

    public render(): React.ReactElement<any> {
        if (this.state.activeCommittees) {
            return (
                <div>
                    <Form
                        onSubmit={this._onSubmit}
                        initialValues={{
                            MemberID: this.props.memberId,
                            MemberLookUpId: this.props.memberId,
                        }}
                        render={(formRenderProps: FormRenderProps) => (
                            <FormElement>
                                <Field
                                    name={'committeeName'}
                                    label={'Committee'}
                                    component={MyComboBox}
                                    validator={value => value ? "" : "Please Select a Committee."}
                                    required={true}
                                    options={this.state.activeCommittees}
                                    onChange={e => {
                                        GetChoiceColumn(e.value, 'Status').then(statusValue => this.setState({ statusOptions: statusValue })).catch(CONSOLE_LOG_ERROR);
                                        GetChoiceColumn(e.value, 'Position').then(positionValue => this.setState({ positionOptions: positionValue })).catch(CONSOLE_LOG_ERROR);
                                        GetCommitteeByName(e.value).then(f => this.setState({ committeeFileItem: f })).catch(CONSOLE_LOG_ERROR);
                                    }}
                                />
                                <Field
                                    name={'_Status'}
                                    label={'Status'}
                                    component={MyComboBox}
                                    validator={value => value ? "" : "Please Select a Status."}
                                    required={true}
                                    options={this.state.statusOptions ? this.state.statusOptions.map((f: any) => { return { key: f, text: f }; }) : []}
                                    disabled={!this.state.committeeFileItem}
                                />
                                <Field
                                    name={'Position'}
                                    label={'Position'}
                                    component={MyComboBox}
                                    required={true}
                                    options={this.state.positionOptions ? this.state.positionOptions.map((f: any) => { return { key: f, text: f }; }) : []}
                                    disabled={!this.state.committeeFileItem}
                                    validator={value => value ? "" : "Please Select a Position."}
                                />
                                <Field
                                    name={'StartDate'}
                                    label={'Term Start Date'}
                                    validator={value => value ? "" : "Please Select a Start Date."}
                                    formatDate={OnFormatDate}
                                    component={MyDatePicker}
                                    onChange={e => {
                                        const CALC_END_DATE = CalculateTermEndDate(e.value, this.state.committeeFileItem.TermLength);
                                        this.setState({ calculatedEndDate: CALC_END_DATE });
                                        formRenderProps.onChange(`_EndDate`, { value: CALC_END_DATE });
                                    }}
                                    required={true}
                                    disabled={!this.state.committeeFileItem}
                                />
                                <Field
                                    name={`_EndDate`}
                                    label={'Term End Date'}
                                    validator={value => value ? "" : "Please Select a End Date."}
                                    formatDate={OnFormatDate}
                                    component={MyDatePicker}
                                    required={true}
                                    disabled={!this.state.committeeFileItem}
                                />
                                <Field
                                    component={FilePicker}
                                    name={`Files`}
                                    buttonIcon="Attach"
                                    buttonLabel='Select Files'
                                    label={'Upload Attachments'}
                                    context={this.props.context}
                                    hideStockImages={true}
                                    hideLinkUploadTab={true}
                                    hideLocalUploadTab={true}
                                    hideRecentTab={true}
                                    disabled={!this.state.committeeFileItem}
                                    // onSave={(filePickerResult: IFilePickerResult[]) => this._pushFileAttachment(filePickerResult)}
                                    onSave={(filePickerResult: IFilePickerResult[]) => {
                                        console.log(filePickerResult);
                                        formRenderProps.onChange(`Files`, { value: filePickerResult });
                                    }}
                                />
                                <div style={{ marginTop: "10px" }}>
                                    <PrimaryButton
                                        text='Submit'
                                        type="submit"
                                        style={{ margin: '5px' }}
                                    // disabled={(this.state.saveStatus === NewMemberFormSaveStatus.Success || this.state.saveStatus === NewMemberFormSaveStatus.Error)}
                                    />
                                    <DefaultButton
                                        text='Clear'
                                        style={{ margin: '5px' }}
                                        onClick={e => {
                                            formRenderProps.onFormReset();
                                            this.setState({ committeeFileItem: undefined });
                                            // this.setState({ saveStatus: NewMemberFormSaveStatus.NewForm, linkToCommitteeDocSet: [] });
                                        }}
                                    />
                                </div>
                            </FormElement>
                        )}
                    />
                </div>
            );
        }
        else {
            // Display shimmer while everything is loading.
            return <MyShimmer />;
        }
    }
}