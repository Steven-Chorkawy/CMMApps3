import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetExecuteEventParameters,
  ListViewStateChangedEventArgs,
  RowAccessor
} from '@microsoft/sp-listview-extensibility';
import { MyCommandSets } from '../../HelperMethods/MyCommandSets';
import { GetMemberIdFromSelectedRow, getSP } from '../../HelperMethods/MyHelperMethods';
import * as ReactDOM from 'react-dom';
import * as React from 'react';
import { CommitteeMemberDashboardPanel } from '../../ClaringtonComponents/CommitteeMemberDashboardPanel';
import { AddMemberPanel } from '../../ClaringtonComponents/AddMemberDialog';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ICommitteeMemberDetailsCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'CommitteeMemberDetailsCommandSet';

export default class CommitteeMemberDetailsCommandSet extends BaseListViewCommandSet<ICommitteeMemberDetailsCommandSetProperties> {

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized CommitteeMemberDetailsCommandSet');

    getSP(this.context);

    // initial state of the command's visibility
    const compareMemberDetailsCommand: Command = this.tryGetCommand(MyCommandSets.MemberDetails);
    const compareAddMemberCommand: Command = this.tryGetCommand(MyCommandSets.AddMember);

    compareMemberDetailsCommand.visible = false;
    compareAddMemberCommand.visible = true;   // This command should always be visible.

    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);

    return Promise.resolve();
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case MyCommandSets.MemberDetails: {
        const selectedRow: RowAccessor = event.selectedRows[0];
        GetMemberIdFromSelectedRow(selectedRow)
          .then(value => {
            const memberDetailPanel: React.ReactComponentElement<any> = React.createElement(CommitteeMemberDashboardPanel, { context: this.context, memberId: value });
            ReactDOM.render(memberDetailPanel, document.createElement('div'));
          })
          .catch(reason => {
            console.error('Failed to Get Member ID from Selected Row');
            console.error(reason);
          });
        break;
      }
      case MyCommandSets.AddMember: {
        const addMemberPanelElement: React.ReactComponentElement<any> = React.createElement(AddMemberPanel, { context: this.context })
        ReactDOM.render(addMemberPanelElement, document.createElement('div'));
        break;
      }
      default:
        throw new Error('Unknown command');
    }
  }

  private _onListViewStateChanged = (args: ListViewStateChangedEventArgs): void => {
    Log.info(LOG_SOURCE, 'List view state changed');

    const compareMemberDetailsCommand: Command = this.tryGetCommand(MyCommandSets.MemberDetails)

    if (compareMemberDetailsCommand) {
      // No need to check if MemberLookup has a value here.  The display form will handle that with error messages.
      compareMemberDetailsCommand.visible = this.context.listView.selectedRows?.length === 1;
    }

    // You should call this.raiseOnChage() to update the command bar
    this.raiseOnChange();
  }
}
