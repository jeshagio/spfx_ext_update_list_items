import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetExecuteEventParameters,
  ListViewStateChangedEventArgs
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface INewListViewCExtCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'NewListViewCExtCommandSet';

export default class NewListViewCExtCommandSet extends BaseListViewCommandSet<INewListViewCExtCommandSetProperties> {

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized NewListViewCExtCommandSet');

    // initial state of the command's visibility
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    compareOneCommand.visible = false;

    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);

    return Promise.resolve();
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'COMMAND_1':
        Dialog.alert(`${this.properties.sampleTextOne}`).catch(() => {
          /* handle error */
        });
        // execute the action for command on async for updating rows
 
        
        //const res: SPHttpClientResponse = await request;

       break;
      case 'COMMAND_2':
        this._onSave().catch()
        .then(() => console.log('this will succeed'));
      /*      
      Dialog.prompt(`Clicked ${this.properties.sampleTextOne}. Enter something to alert:`).then((value: string) => {
          Dialog.alert(value);
        });*/
        break;
        
      default:
        throw new Error('Unknown command');
    }
  }

  private _onSave = async (): Promise<void> => {

  
    

    // this.context.listView.selectedRows[0].getValueByName("ID").toString()+
    this.context.listView.selectedRows.forEach(
      async row => {
        let request: Promise<SPHttpClientResponse>;
        var uniqueIdL=row.getValueByName("UniqueId").toLowerCase();
        var etag='"'+uniqueIdL.replace('{','').replace('}','') +","+row.getValueByName("owshiddenversion")+'"';
        const now=new Date(); 
        request=this._updateItem(now.toString(),row.getValueByName("ID"),etag);
        
      
        const res: SPHttpClientResponse = await request;
      
        if (res.ok) {
          // Refresh view listing
          location.reload();
          console.log(`Item updated! ${uniqueIdL}`);
        }
        else {
          const error: { error: { message: string } } = await res.json();
          console.log(`An error has occurred while saving the item. Please try again. Error: ${error.error.message}`);

        }

      }
    );
    
  }
  
  private _onListViewStateChanged = (args: ListViewStateChangedEventArgs): void => {
    Log.info(LOG_SOURCE, 'List view state changed');

    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = this.context.listView.selectedRows?.length === 1;
    }

    // TODO: Add your logic here

    // You should call this.raiseOnChage() to update the command bar
    this.raiseOnChange();
  }

  private _updateItem(title: string, idItem :number, etag: string): Promise<SPHttpClientResponse> {
    var urlV=this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getByTitle('Testlist')/items(${idItem})`;
    return this.context.spHttpClient
      .post(urlV, SPHttpClient.configurations.v1, {
        headers: {
          'content-type': 'application/json;odata.metadata=none',
          'if-match': etag,
          'x-http-method': 'MERGE',
        },
        body: JSON.stringify({
          Title: title
        })
      });
  }
}
