import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'SendEmailCommandSetStrings';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ISendEmailCommandSetProperties {
  // This is an example; replace with your own properties
  listFieldTitle: string;
  listFieldsEmail: string[];
}

const LOG_SOURCE: string = 'SendEmailCommandSet';

export default class SendEmailCommandSet extends BaseListViewCommandSet<ISendEmailCommandSetProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized SendEmailCommandSet');
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const command: Command = this.tryGetCommand('spfxEmailTo');
    if (command) {
      // This command should be hidden unless exactly one row is selected.
      command.visible = event.selectedRows.length >= 1;
    }
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'spfxEmailTo':
        Dialog.alert(`${this.properties.listFieldTitle}Test1`);
        break;
      default:
        throw new Error('Unknown command');
    }
  }
}
