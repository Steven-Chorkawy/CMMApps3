import * as React from 'react';
import styles from './NewCommitteeMember.module.scss';
import { INewCommitteeMemberProps } from './INewCommitteeMemberProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Field, FieldArray, Form, FormElement, FormRenderProps } from '@progress/kendo-react-form';
import { DatePicker, DefaultButton, getTheme, Link, MessageBar, MessageBarType, PrimaryButton, ProgressIndicator, Separator, TextField } from '@fluentui/react';
import { emailValidator } from '../../../HelperMethods/Validators';
import { GetChoiceColumn, GetListOfActiveCommittees, OnFormatDate } from '../../../HelperMethods/MyHelperMethods';
import { MyComboBox, PhoneInput, PostalCodeInput } from '../../../ClaringtonComponents/MyFormComponents';
import { NewCommitteeMemberFormComponent } from '../../../ClaringtonComponents/NewCommitteeMemberFormComponent';



export enum NewMemberFormSaveStatus {
  NewForm = -1,
  CreatingNewMember = 0,
  AddingMemberToCommittee = 1,
  Success = 2,
  Error = 3
}

export interface INewMemberFormState {
  activeCommittees: any[];
  // provinces: any[];
  saveStatus: NewMemberFormSaveStatus;
  linkToCommitteeDocSet: any[];
}

export default class NewCommitteeMember extends React.Component<INewCommitteeMemberProps, INewMemberFormState> {

  constructor(props: any) {
    super(props);

    this.state = {
      activeCommittees: [],
      // provinces: [],
      saveStatus: NewMemberFormSaveStatus.NewForm,
      linkToCommitteeDocSet: []
    };

    GetListOfActiveCommittees().then(value => {
      console.log('Active Committees');
      console.log(value);
      this.setState({ activeCommittees: value });
    });

    // GetChoiceColumn('Members', 'Province').then(value => {
    //   this.setState({ provinces: value });
    // });
  }

  public render(): React.ReactElement<INewCommitteeMemberProps> {

    const reactTheme = getTheme();

    const EmailInput = (fieldRenderProps: any) => {
      const { validationMessage, visited, ...others } = fieldRenderProps;
      return <TextField {...others} errorMessage={visited && validationMessage && validationMessage} />;
    };

    return (
      <div>
        <Form
          onSubmit={dateItem => console.log(dateItem)}
          render={(formRenderProps: FormRenderProps) => (
            <FormElement>
              <h2>Add New Member</h2>
              <div style={{ padding: '10px', marginBottom: '10px', boxShadow: reactTheme.effects.elevation16 }}>
                {/* <Field name={'Member.Salutation'} label={'Salutation'} component={TextField} /> */}
                <Field name={'Member.FirstName'} label={'First Name'} required={true} component={TextField} />
                <Field name={'Member.MiddleName'} label={'Middle Name'} component={TextField} />
                <Field name={'Member.LastName'} label={'Last Name'} required={true} component={TextField} />
                {/* <Field name={'Member.Birthday'} label={'Date of Birth'} component={DatePicker} formatDate={OnFormatDate} /> */}

                <Field name={'Member.EMail'} label={'Email'} validator={emailValidator} component={EmailInput} />
                <Field name={'Member.Email2'} label={'Email 2'} validator={emailValidator} component={EmailInput} />

                <Field name={'Member.CellPhone1'} label={'Cell Phone'} component={PhoneInput} onChange={e => formRenderProps.onChange(e.name, e.value)} />
                <Field name={'Member.WorkPhone'} label={'Work Phone'} component={PhoneInput} onChange={e => formRenderProps.onChange(e.name, e.value)} />
                <Field name={'Member.HomePhone'} label={'Home Phone'} component={PhoneInput} onChange={e => formRenderProps.onChange(e.name, e.value)} />

                <Field name={'Member.WorkAddress'} label={'Street Address'} component={TextField} />
                <Field name={'Member.WorkCity'} label={'City'} component={TextField} />
                <Field name={'Member.PostalCode'} label={'Postal Code'} component={PostalCodeInput} onChange={e => formRenderProps.onChange(e.name, e.value)} />
                {/** !!! TODO: Get these values from SharePoint, not hard coded.  */}
                {/* <Field name={'Member.Province'}
                  label={'Province'}
                  component={MyComboBox}
                  options={this.state.provinces ? this.state.provinces.map(f => { return { key: f, text: f }; }) : []}
                /> */}
              </div>
              <h2>Add "{formRenderProps.valueGetter('Member.FirstName')} {formRenderProps.valueGetter('Member.LastName')}" to Committee</h2>
              {
                this.state.activeCommittees.length > 0 &&
                <FieldArray
                  name={'Committees'}
                  component={NewCommitteeMemberFormComponent}
                  context={this.props.context}
                  activeCommittees={this.state.activeCommittees}
                  formRenderProps={formRenderProps}
                />
              }
              {
                (this.state.saveStatus === NewMemberFormSaveStatus.CreatingNewMember || this.state.saveStatus === NewMemberFormSaveStatus.AddingMemberToCommittee) &&
                <div style={{ marginTop: '20px' }}>
                  <ProgressIndicator
                    label="Saving New Committee Member..."
                    description={<div>
                      {this.state.saveStatus === NewMemberFormSaveStatus.CreatingNewMember && "Saving Member Contact Information..."}
                      {this.state.saveStatus === NewMemberFormSaveStatus.AddingMemberToCommittee && "Adding Member to Committee..."}
                    </div>}
                  />
                </div>
              }
              {
                (this.state.saveStatus === NewMemberFormSaveStatus.Success) &&
                <MessageBar messageBarType={MessageBarType.success} isMultiline={true}>
                  <div>
                    Success! New Committee Member has been saved.
                    {
                      this.state.linkToCommitteeDocSet.map(l => {
                        return <div>
                          <Separator />
                          <Link href={`${l.Link}`} target="_blank" underline>Click to View: {l.MemberName} - {l.CommitteeName}</Link>
                        </div>;
                      })
                    }
                  </div>
                </MessageBar>
              }
              {
                (this.state.saveStatus === NewMemberFormSaveStatus.Error) &&
                <MessageBar messageBarType={MessageBarType.error}>
                  Something went wrong!  Cannot save new Committee Member.
                </MessageBar>
              }
              <div style={{ marginTop: "10px" }}>
                <PrimaryButton
                  text='Submit'
                  type="submit"
                  style={{ margin: '5px' }}
                  disabled={(this.state.saveStatus === NewMemberFormSaveStatus.Success || this.state.saveStatus === NewMemberFormSaveStatus.Error)}
                />
                <DefaultButton
                  text='Clear'
                  style={{ margin: '5px' }}
                  onClick={e => {
                    formRenderProps.onFormReset();
                    this.setState({ saveStatus: NewMemberFormSaveStatus.NewForm, linkToCommitteeDocSet: [] });
                  }}
                />
              </div>
            </FormElement>
          )}
        />
      </div>
    );
  }
}
