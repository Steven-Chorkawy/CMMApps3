import * as React from 'react';
import styles from './NewCommitteeMember.module.scss';
import { INewCommitteeMemberProps } from './INewCommitteeMemberProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Field, FieldArray, Form, FormElement, FormRenderProps } from '@progress/kendo-react-form';
import { DatePicker, DefaultButton, getTheme, Link, MessageBar, MessageBarType, PrimaryButton, ProgressIndicator, Separator, TextField } from '@fluentui/react';
import { emailValidator } from '../../../HelperMethods/Validators';
import { CreateNewMember, FormatDocumentSetPath, GetChoiceColumn, GetListOfActiveCommittees, OnFormatDate } from '../../../HelperMethods/MyHelperMethods';
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
      this.setState({ activeCommittees: value });
    });

    // GetChoiceColumn('Members', 'Province').then(value => {
    //   this.setState({ provinces: value });
    // });
  }

  private _onSubmit = async (values: any) => {
    try {
      console.log('_onSubmit');
      console.log(values);

      this.setState({ saveStatus: NewMemberFormSaveStatus.CreatingNewMember });

      // Step 1: Add the new member to the Members List.
      let newMember_IAR = await CreateNewMember(values.Member);

      console.log('new member add results.');
      console.log(newMember_IAR);

      // Step 2: Add the new member to committess if any are provided. 
      if (values.Committees) {
        this.setState({ saveStatus: NewMemberFormSaveStatus.AddingMemberToCommittee });
        for (let committeeIndex = 0; committeeIndex < values.Committee.length; committeeIndex++) {
          const currentCommittee = values.Committee[committeeIndex];
          await CreateNewCommitteeMember(newMember_IAR.data.ID, currentCommittee);
          let linkToDocSet = await FormatDocumentSetPath(currentCommittee.CommitteeName, newMember_IAR.data.Title);
          this.setState({
            linkToCommitteeDocSet: [
              ...this.state.linkToCommitteeDocSet,
              {
                CommitteeName: currentCommittee.CommitteeName,
                MemberName: newMember_IAR.data.Title,
                Link: linkToDocSet
              }
            ]
          });
        }
      }
      this.setState({ saveStatus: NewMemberFormSaveStatus.Success });
    } catch (error) {
      this.setState({ saveStatus: NewMemberFormSaveStatus.Error });
      console.log("Something went wrong while saving new member!");
      console.error(error);
    }
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
          onSubmit={this._onSubmit}
          render={(formRenderProps: FormRenderProps) => (
            <FormElement>
              <h2>Add New Member</h2>
              <div style={{ padding: '10px', marginBottom: '10px', boxShadow: reactTheme.effects.elevation16 }}>
                <Field name={'Member.FirstName'} label={'First Name'} required={true} component={TextField} />
                <Field name={'Member.LastName'} label={'Last Name'} required={true} component={TextField} />

                <Field name={'Member.EMail'} label={'Email'} validator={emailValidator} component={EmailInput} />
                <Field name={'Member.Email2'} label={'Email 2'} validator={emailValidator} component={EmailInput} />

                <Field name={'Member.CellPhone'} label={'Cell Phone'} component={PhoneInput} onChange={e => formRenderProps.onChange(e.name, e.value)} />
                <Field name={'Member.WorkPhone'} label={'Work Phone'} component={PhoneInput} onChange={e => formRenderProps.onChange(e.name, e.value)} />
                <Field name={'Member.HomePhone'} label={'Home Phone'} component={PhoneInput} onChange={e => formRenderProps.onChange(e.name, e.value)} />

                <Field name={'Member.WorkAddress'} label={'Street Address'} component={TextField} />
                <Field name={'Member.WorkCity'} label={'City'} component={TextField} />
                <Field name={'Member.PostalCode'} label={'Postal Code'} component={PostalCodeInput} onChange={e => formRenderProps.onChange(e.name, e.value)} />          
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
function CreateNewCommitteeMember(ID: any, arg1: any) {
  throw new Error('Function not implemented.');
}

