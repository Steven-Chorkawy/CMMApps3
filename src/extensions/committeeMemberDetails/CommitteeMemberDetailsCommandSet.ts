import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetExecuteEventParameters,
  ListViewStateChangedEventArgs,
  RowAccessor
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import { MyCommandSets } from '../../HelperMethods/MyCommandSets';
import { GetMemberIdFromSelectedRow, getSP } from '../../HelperMethods/MyHelperMethods';
import { MyLists } from '../../HelperMethods/MyLists';
import * as ReactDOM from 'react-dom';
import * as React from 'react';
import { CommitteeMemberDashboardPanel } from '../../ClaringtonComponents/CommitteeMemberDashboardPanel';

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

  private panelPlaceHolder: HTMLDivElement = null;

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized CommitteeMemberDetailsCommandSet');

    getSP(this.context);

    this.panelPlaceHolder = document.body.appendChild(document.createElement("div"));

    // initial state of the command's visibility
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    const compareMemberDetailsCommand: Command = this.tryGetCommand(MyCommandSets.MemberDetails);

    compareOneCommand.visible = false;
    compareMemberDetailsCommand.visible = false;

    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);

    return Promise.resolve();
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'COMMAND_1':
        Dialog.alert(`${this.properties.sampleTextOne}`).catch(() => {
          /* handle error */
        });
        break;
      case 'COMMAND_2':
        Dialog.alert(`${this.properties.sampleTextTwo}`).catch(() => {
          /* handle error */
        });
        break;
      case MyCommandSets.MemberDetails:
        console.log(event);
        const selectedRow: RowAccessor = event.selectedRows[0];

        GetMemberIdFromSelectedRow(selectedRow).then(value => {
          const memberDetailPanel: React.ReactComponentElement<any> = React.createElement(CommitteeMemberDashboardPanel, { context: this.context, memberId: value });
          const panelDiv = document.createElement('div');
          ReactDOM.render(memberDetailPanel, panelDiv);
        });
        break;
      default:
        throw new Error('Unknown command');
    }
  }

  private _onListViewStateChanged = (args: ListViewStateChangedEventArgs): void => {
    Log.info(LOG_SOURCE, 'List view state changed');

    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    const compareMemberDetailsCommand: Command = this.tryGetCommand(MyCommandSets.MemberDetails)
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = this.context.listView.selectedRows?.length === 1;
    }

    if (compareMemberDetailsCommand) {
      if (this.context.listView.list.title === MyLists.Members || this.context.listView.columns.some(e => e.field.internalName === 'MemberLookup')) {
        // Only show when in the Members list, when the current list has a 'MemberLookup' field, and when one row is selected.
        // No need to check if MemberLookup has a value here.  The display form will handle that with error messages.
        compareMemberDetailsCommand.visible = this.context.listView.selectedRows?.length === 1;
      }
    }

    // You should call this.raiseOnChage() to update the command bar
    this.raiseOnChange();
  }
}
