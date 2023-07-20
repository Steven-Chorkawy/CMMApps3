import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetExecuteEventParameters,
  ListViewStateChangedEventArgs
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import { MyCommandSets } from '../../HelperMethods/MyCommandSets';

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

    // initial state of the command's visibility
    const compareOneCommand: Command = this.tryGetCommand(MyCommandSets.RenewCommitteeMember);
    compareOneCommand.visible = false;

    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);

    return Promise.resolve();
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    console.log('onExecute......');
    console.log(event);
    switch (event.itemId) {
      case MyCommandSets.RenewCommitteeMember:
        Dialog.alert(`Renew Committee Member Clicked!`).catch(() => {
          /* handle error */
        });
        break;
      // case 'COMMAND_2':
      //   Dialog.alert(`${this.properties.sampleTextTwo}`).catch(() => {
      //     /* handle error */
      //   });
      //   break;
      default:
        throw new Error('Unknown command');
    }
  }

  private _onListViewStateChanged = (args: ListViewStateChangedEventArgs): void => {
    Log.info(LOG_SOURCE, 'List view state changed');
    console.log('_onListViewStateChanged......');
    console.log(args);

    const compareOneCommand: Command = this.tryGetCommand(MyCommandSets.RenewCommitteeMember);
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = this.context.listView.selectedRows?.length === 1;
    }

    // TODO: Add your logic here

    // You should call this.raiseOnChage() to update the command bar
    this.raiseOnChange();
  }
}
