import * as React from 'react';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { PrimaryButton, DefaultButton, ActionButton, IconButton } from 'office-ui-fabric-react/lib/Button';
import { getId } from 'office-ui-fabric-react/lib/Utilities';
import { hiddenContentStyle, mergeStyles } from 'office-ui-fabric-react/lib/Styling';
// import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { ContextualMenu } from 'office-ui-fabric-react/lib/ContextualMenu';
import { sp, Item, Web } from "@pnp/sp";
// import { AttachmentFile, AttachmentFiles, AttachmentFileInfo } from '@pnp/sp/src/attachmentfiles'; 
import { IEventCardProps } from '../EventCard';


const screenReaderOnly = mergeStyles(hiddenContentStyle);

export interface IDialogBasicExampleState {
  hideDialog: boolean;
  isDraggable: boolean;
}

export class DialogBasicExample extends React.Component<IEventCardProps , IDialogBasicExampleState> {
  public state: IDialogBasicExampleState = {
    hideDialog: true,
    isDraggable: false
  };

  // Use getId() to ensure that the IDs are unique on the page.
  // (It's also okay to use plain strings without getId() and manually ensure uniqueness.)
  private _labelId: string = getId('dialogLabel');
  private _subTextId: string = getId('subTextLabel');
  private _dragOptions = {
    moveMenuItemText: 'Move',
    closeMenuItemText: 'Close',
    menu: ContextualMenu
  };

//Retrieve all the attachments from all items in a SharePoint List  
    private getSPData(): void {  
        let attachmentfiles: string = "";  
        // console.log(this.props.event.id)
        // console.log(this.testVariables.identity);
        // let urlList = "https://keckmedicine.sharepoint.com/sites/_api/web/lists/getByTitle('Events')/items?$expand=AttachmentFiles&$filter=Attachments%20eq%201"
        const urlList = new Web("https://keckmedicine.sharepoint.com/sites/KSOM-Intranet/");
        // console.log(urlList.lists.getByTitle("Events").items);
        // sp.web.lists
        urlList.lists.getByTitle("Events").items  
        .select("Id,Title,Attachments,AttachmentFiles")  
        .expand("AttachmentFiles")  
        .filter('Attachments eq 1')  
        .get()
        .then((response: Item[]) => {  
        // console.log(response);
        response.forEach((listItem: any) => {  
            if(listItem.Id == this.props.event.id){
                listItem.AttachmentFiles.forEach((afile: any) => {  
                    // console.log(afile.FileName)
                    // work on the functionality to get the list item id and pass that before the filename
                    let downloadUrl = "https://keckmedicine.sharepoint.com/sites/KSOM-Intranet/Lists/Events/Attachments/" + `${listItem.Id}/` + afile.FileName;
                    // console.log(downloadUrl)
                    // let downloadUrl = this.context.pageContext.web.absoluteUrl + "/_layouts/download.aspx?sourceurl=" + afile.ServerRelativeUrl;  
                    // attachmentfiles += `<li>(${listItem.Id}) ${listItem.Title} - <a href='${downloadUrl}'>${afile.FileName}</a></li>`;  
                    attachmentfiles += `<li><a href='${downloadUrl}'>${afile.FileName}</a></li>`;  
                });
            }
              
        });      
        attachmentfiles = `<ul>${attachmentfiles}</ul>`;
        // console.log(attachmentfiles); 
        this.renderData(attachmentfiles);  
        });  
    } 

    private renderData(strResponse: string): void {  
        const htmlElement = document.getElementById("attachmentFiles");  
        htmlElement.innerHTML = strResponse;  
    } 

  public render() {
    const { hideDialog, isDraggable } = this.state;

    return (
      <span onClick = {this.handleChildClick}>
        {/* <Checkbox label="Is draggable" onChange={this._toggleDraggable} checked={isDraggable} /> */}
        <IconButton
            // className={styles.addToMyCalendar}
            iconProps={{ iconName: "PhotoVideoMedia" }}
            onClick={this._showDialog}
            // ariaLabel={strings.AddToCalendarAriaLabel}
            // onClick={this._onAddToMyCalendar}
        >
            {/* {strings.AddToCalendarButtonLabel} */}
        </IconButton>
        {/* <DefaultButton secondaryText="Opens the Sample Dialog" onClick={this._showDialog} text="Open Dialog" /> */}
        <label id={this._labelId} className={screenReaderOnly}>
          My sample Label
        </label>
        <label id={this._subTextId} className={screenReaderOnly}>
          My Sample description
        </label>

        <Dialog
          hidden={hideDialog}
          onDismiss={this._closeDialog}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Attachments',
            closeButtonAriaLabel: 'Close',
            subText: 'Click on the links to open documents'
          }}
          modalProps={{
            titleAriaId: this._labelId,
            subtitleAriaId: this._subTextId,
            isBlocking: false,
            styles: { main: { maxWidth: 450 } },
            dragOptions: isDraggable ? this._dragOptions : undefined
          }}
        >
        <div id ="attachmentFiles"></div>
          {/* <DialogFooter>
            <PrimaryButton onClick={this._closeDialog} text="Send" />
            <DefaultButton onClick={this._closeDialog} text="Don't send" />
          </DialogFooter> */}
        </Dialog>
        {this.getSPData()}
      </span>
    );
  }

  private handleChildClick = (e): void => {
    e.stopPropagation();
    // console.log("handleChildClick");
  }

  private _showDialog = (): void => {
    this.setState({ hideDialog: false });
  }

  private _closeDialog = (): void => {
    this.setState({ hideDialog: true });
  }

  private _toggleDraggable = (): void => {
    this.setState({ isDraggable: !this.state.isDraggable });
  }
}
