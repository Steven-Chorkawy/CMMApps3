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
import * as React from 'react';
import { RenewCommitteeMemberPanel } from '../../ClaringtonComponents/RenewCommitteeMember';
import * as ReactDOM from 'react-dom';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IRenewCommitteeMemberCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'RenewCommitteeMemberCommandSet';

export default class RenewCommitteeMemberCommandSet extends BaseListViewCommandSet<IRenewCommitteeMemberCommandSetProperties> {

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized RenewCommitteeMemberCommandSet');
    getSP(this.context);

    // initial state of the command's visibility
    const compareOneCommand: Command = this.tryGetCommand(MyCommandSets.RenewCommitteeMember);
    compareOneCommand.visible = false;

    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);

    return Promise.resolve();
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case MyCommandSets.RenewCommitteeMember:
        const selectedRow: RowAccessor = event.selectedRows[0];
          GetMemberIdFromSelectedRow(selectedRow).then(value => {
          const memberDetailPanel: React.ReactComponentElement<any> = React.createElement(RenewCommitteeMemberPanel, { context: this.context, memberId: value });
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
 
    const compareOneCommand: Command = this.tryGetCommand(MyCommandSets.RenewCommitteeMember);
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = this.context.listView.selectedRows?.length === 1;
    }

    // You should call this.raiseOnChage() to update the command bar
    this.raiseOnChange();
  }
}
