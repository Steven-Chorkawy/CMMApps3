import * as React from 'react';
import { INewCommitteeMemberProps } from './INewCommitteeMemberProps';
import { Field, FieldArray, Form, FormElement, FormRenderProps } from '@progress/kendo-react-form';
import { DefaultButton, getTheme, Link, MessageBar, MessageBarType, PrimaryButton, ProgressIndicator, Separator, TextField } from '@fluentui/react';
import { emailValidator } from '../../../HelperMethods/Validators';
import { CreateNewCommitteeMember, CreateNewMember, FormatDocumentSetPath, GetListOfActiveCommittees } from '../../../HelperMethods/MyHelperMethods';
import { EmailInput, PhoneInput, PostalCodeInput } from '../../../ClaringtonComponents/MyFormComponents';
import { NewCommitteeMemberFormComponent } from '../../../ClaringtonComponents/NewCommitteeMemberFormComponent';
import PackageSolutionVersion from '../../../ClaringtonComponents/PackageSolutionVersion';
import { MyShimmer } from '../../../ClaringtonComponents/MyShimmer';

export enum NewMemberFormSaveStatus {
  NewForm = -1,
  CreatingNewMember = 0,
  AddingMemberToCommittee = 1,
  Success = 2,
  Error = 3
}

export interface INewMemberFormState {
  activeCommittees: any[];
  failedToLoadActiveCommittees: boolean;  // True: Display error messages.  False: Hide error messages.
  // provinces: any[];
  saveStatus: NewMemberFormSaveStatus;
  linkToCommitteeDocSet: any[];
}

export default class NewCommitteeMember extends React.Component<INewCommitteeMemberProps, INewMemberFormState> {

  constructor(props: any) {
    super(props);

    this.state = {
      activeCommittees: [],
      failedToLoadActiveCommittees: false,  // ! Must be set to false or else error messages will be displayed by default.
      // provinces: [],
      saveStatus: NewMemberFormSaveStatus.NewForm,
      linkToCommitteeDocSet: []
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
    try {
      this.setState({ saveStatus: NewMemberFormSaveStatus.CreatingNewMember });

      // Step 1: Add the new member to the Members List.
      const newMember_IAR = await CreateNewMember(values.Member);

      // Step 2: Add the new member to committess if any are provided. 
      if (values.Committees) {
        this.setState({ saveStatus: NewMemberFormSaveStatus.AddingMemberToCommittee });
        for (let committeeIndex = 0; committeeIndex < values.Committees.length; committeeIndex++) {
          const currentCommittee = values.Committees[committeeIndex];
          await CreateNewCommitteeMember(newMember_IAR.data.ID, currentCommittee);
          const linkToDocSet = await FormatDocumentSetPath(currentCommittee.CommitteeName, newMember_IAR.data.Title);
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
      console.error("Something went wrong while saving new member!");
      console.error(error);
    }
  }

  public render(): React.ReactElement<INewCommitteeMemberProps> {

    const reactTheme = getTheme();

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
                {/* <Field name={'Member.Email2'} label={'Email 2'} validator={emailValidator} component={EmailInput} /> */}

                <Field name={'Member.CellPhone1'} label={'Cell Phone'} component={PhoneInput} onChange={e => formRenderProps.onChange(e.name, e.value)} />
                {/* <Field name={'Member.WorkPhone'} label={'Work Phone'} component={PhoneInput} onChange={e => formRenderProps.onChange(e.name, e.value)} /> */}
                <Field name={'Member.HomePhone'} label={'Home Phone'} component={PhoneInput} onChange={e => formRenderProps.onChange(e.name, e.value)} />

                <Field name={'Member.WorkAddress'} label={'Street Address'} component={TextField} />
                <Field name={'Member.WorkCity'} label={'City'} component={TextField} />
                <Field name={'Member.PostalCode'} label={'Postal Code'} component={PostalCodeInput} onChange={e => formRenderProps.onChange(e.name, e.value)} />
              </div>
              {
                (this.state.activeCommittees.length > 0 && this.state.failedToLoadActiveCommittees === false) ?
                  <div>
                    <h2>Add '{formRenderProps.valueGetter('Member.FirstName')} {formRenderProps.valueGetter('Member.LastName')}' to Committee</h2>
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
              {
                this.state.failedToLoadActiveCommittees === true &&
                <div>
                  <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
                    Failed to load active committees. <Link href={window.location.href} underline>Click here to try again.</Link>
                  </MessageBar>
                </div>
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
                      this.state.linkToCommitteeDocSet.map((l, index) => {
                        return <div key={`${l.MemberName}${index}`}>
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
                  title={this.state.failedToLoadActiveCommittees ? 'Cannot save without a list of committees.' : 'Click to Save New Member'}
                  disabled={(this.state.saveStatus === NewMemberFormSaveStatus.Success || this.state.failedToLoadActiveCommittees === true)}
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
        <PackageSolutionVersion />
      </div>
    );
  }
}
