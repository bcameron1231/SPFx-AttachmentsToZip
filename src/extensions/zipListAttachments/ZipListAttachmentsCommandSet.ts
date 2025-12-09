import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  type Command,
  type IListViewCommandSetExecuteEventParameters,
  type ListViewStateChangedEventArgs
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import { SPFI, SPFx, spfi } from "@pnp/sp";
import { IItem } from "@pnp/sp/items/types";
import { IAttachmentInfo } from "@pnp/sp/attachments";
import "@pnp/sp/items";
import "@pnp/sp/webs";
import "@pnp/sp/attachments";
import "@pnp/sp/lists";
import * as JSZip from 'jszip';
import { saveAs } from 'file-saver';
/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IZipListAttachmentsCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'ZipListAttachmentsCommandSet';

export default class ZipListAttachmentsCommandSet extends BaseListViewCommandSet<IZipListAttachmentsCommandSetProperties> {
  private sp:SPFI;
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized ZipListAttachmentsCommandSet');

    this.sp = spfi().using(SPFx(this.context));

    // initial state of the command's visibility
    const compareOneCommand: Command = this.tryGetCommand('ATTACHMENT_ZIP_COMMAND');
    compareOneCommand.visible = false;

    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);

    return Promise.resolve();
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'ATTACHMENT_ZIP_COMMAND':
        const itemId = parseInt(event.selectedRows?.[0].getValueByName("ID"));
        this.ConvertAttachmentsToZip(itemId).catch((error) => {
          Dialog.alert("Unable to zip attachments");
        });  
        break;
      default:
        throw new Error('Unknown command');
    }
  }

  private _onListViewStateChanged = (args: ListViewStateChangedEventArgs): void => {
    Log.info(LOG_SOURCE, 'List view state changed');

    const compareOneCommand: Command = this.tryGetCommand('ATTACHMENT_ZIP_COMMAND');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = this.context.listView.selectedRows?.length === 1;
    }

    // You should call this.raiseOnChage() to update the command bar
    this.raiseOnChange();
  }

  private async ConvertAttachmentsToZip(itemId:number): Promise<void> {
    try{

        if(this.context.listView.list?.title){

        const item:IItem = this.sp.web.lists.getByTitle(this.context.listView.list?.title).items.getById(itemId);
        const attachments:IAttachmentInfo[] = await item.attachmentFiles();

        if(attachments.length > 0){
          var zip = new JSZip();
          // Logic to zip attachments goes here
          for(const attachment of attachments){
            const blob = await item.attachmentFiles.getByName(attachment.FileName).getBlob();

            zip.file(attachment.FileName, blob);
          }

          await zip.generateAsync({type:"blob"}).then(function(content) {
            // see FileSaver.js
            saveAs(content, `Attachments_Item_${itemId}.zip`);
          });
        }
      }
    }
    catch(error){
      throw error;
    }
  }
}
