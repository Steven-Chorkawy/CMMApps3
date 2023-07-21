import { ComboBox, DatePicker, IComboBoxOption } from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import * as React from 'react';
import { CONSOLE_LOG_ERROR, CalculateTermEndDate, GetChoiceColumn, GetCommitteeByName, GetListOfActiveCommittees, OnFormatDate, getSP } from '../HelperMethods/MyHelperMethods';
import { MyShimmer } from './MyShimmer';
import { Field, Form, FormElement, FormRenderProps } from '@progress/kendo-react-form';
import { MyComboBox, MyDatePicker } from './MyFormComponents';
import ICommitteeFileItem from '../ClaringtonInterfaces/ICommitteeFileItem';
import { FilePicker, IFilePickerResult } from '@pnp/spfx-controls-react';

export interface IRenewMemberComponentProps {
    context: WebPartContext;
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
    }

    public render(): React.ReactElement<any> {
        if (this.state.activeCommittees) {
            return (
                <div>
                    <Form
                        onSubmit={this._onSubmit}
                        render={(formRenderProps: FormRenderProps) => (
                            <FormElement>
                                <Field
                                    name={'committee'}
                                    label={'Committee'}
                                    component={MyComboBox}
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
                                    options={this.state.statusOptions ? this.state.statusOptions.map((f: any) => { return { key: f, text: f }; }) : []}
                                    disabled={!this.state.committeeFileItem}
                                />
                                <Field
                                    name={'position'}
                                    label={'Position'}
                                    component={MyComboBox}
                                    options={this.state.positionOptions ? this.state.positionOptions.map((f: any) => { return { key: f, text: f }; }) : []}
                                    disabled={!this.state.committeeFileItem}
                                />
                                <Field
                                    name={'StartDate'}
                                    label={'Term Start Date'}
                                    //allowTextInput={true}
                                    formatDate={OnFormatDate}
                                    component={MyDatePicker}
                                    onChange={e => {
                                        debugger;
                                        const CALC_END_DATE = CalculateTermEndDate(e.value, this.state.committeeFileItem.TermLength);
                                        this.setState({ calculatedEndDate: CALC_END_DATE });
                                        formRenderProps.onChange(`_EndDate`, { value: CALC_END_DATE });
                                    }}
                                    required={true}
                                    validator={value => value ? "" : "Please Select a Start Date."}
                                    disabled={!this.state.committeeFileItem}
                                />
                                <Field
                                    name={`_EndDate`}
                                    label={'Term End Date'}
                                    formatDate={OnFormatDate}
                                    component={DatePicker}
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