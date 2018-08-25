import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';

import { sp } from "@pnp/sp";

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ISendEmailCommandSetProperties {
  // This is an example; replace with your own properties
  listFieldsEmail: string[];
  defaultTo: string;
}

const LOG_SOURCE: string = 'SendEmailCommandSet';

export default class SendEmailCommandSet extends BaseListViewCommandSet<ISendEmailCommandSetProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized SendEmailCommandSet');

    //Provide PnP JS-Core with the proper context (needed in SPFx Components)
    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context
      });
    });
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

        // We'll request all the selected items in a single batch
        let itemBatch: any = sp.createBatch();

        //This will be our cleansed items to clone array
        let items: Array<any> = new Array<any>();

        //Batch up each item for retrieval
        for (let row of event.selectedRows) {

          //grab the item ID
          let itemId: number = row.getValueByName('ID');

          //Add the item to the batch
          sp.web.lists.getById(this.context.pageContext.list.id.toString()).items.getById(itemId).select(...this.properties.listFieldsEmail).inBatch(itemBatch).get<Array<any>>()
            .then((result: any) => {
              items.push(result);
            })
            .catch((error: any): void => {
              Log.error(LOG_SOURCE, error);
              this.safeLog(error);
            });
        }

        //Execute the batch
        itemBatch.execute()
          .then(() => {

            let emailAddresses: string[] = [];
            // send command to browser to open default mail client       
            this.safeLog("all items retrieved");
            this.safeLog(items);
            items.forEach((item: any)=>{
              this.properties.listFieldsEmail.forEach((listField:string)=>{
                if(item[listField]){
                  emailAddresses.push(item[listField]);
                }
              });
            });
            this.safeLog("all email addresses");
            this.safeLog(emailAddresses);
            let mailto: string = `mailto:${this.properties.defaultTo}?bcc=${emailAddresses.join(';')}`;
            this.safeLog(mailto);
            window.location.href = mailto;
          })
          .catch((error: any): void => {
            Log.error(LOG_SOURCE, error);
            this.safeLog(error);
          });

        break;
      default:
        throw new Error('Unknown command');
    }
  }
  /** Logs messages to the console if the console is available */
  private safeLog(message: any): void {
    if (window.console && window.console.log) {
      window.console.log(message);
    }
  }
}
