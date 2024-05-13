import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import { SPPermission } from '@microsoft/sp-page-context';
import { BaseListViewCommandSet, Command, IListViewCommandSetListViewUpdatedParameters, IListViewCommandSetExecuteEventParameters } from '@microsoft/sp-listview-extensibility';
import HtmlDialog from './component/HtmlDialog';
import { SPHttpClient, ISPHttpClientOptions, SPHttpClientResponse } from '@microsoft/sp-http';
import SharePointServices from './services/SharepointServices';

export interface IBulkUpdateDocLibProperties {
}

// Constant values
const LOG_SOURCE: string = 'BulkUpdate';
const cancel = 'Cancel';
const ok = 'OK';
const update = 'Update';
const validate = 'Validate';
const COMMANDBULKUPDATE = 'COMMAND_BULK_UPDATE';
const BulkUpdate = 'Bulk Update';
const documentTitle = 'documents';
const materialsListInternalName = 'materialslist';
var title: string = '';
var message: string = '';
var textAreaValue: string = '';
var invalidMaterials: string = '';
var dialog: any;
var uploadIds: string[] = [];
var validCollection: any[] = [];
var uniqueTitle: string[] = [];

export default class BulkUpdateDocLib extends BaseListViewCommandSet<IBulkUpdateDocLibProperties> {
  constructor() {
    super();
    this.send = this.send.bind(this);
    this.validate = this.validate.bind(this);
  }

  private _fileRefRelativeUrl: string;
  private _triggerType: string;
  private _allValidMaterials: boolean;
  private _updateDone: boolean;
  private _webAbsoluteUrl: string;
  private _documentLibServerRelativeUrl: string;
  private _documentLibIdObj: any;
  private _documentLibId: string;
  private _documentLibTitle: string;
  private _materialsListID: string;
  private _userEmail: string;
  private _spHttpClient: SPHttpClient;

  @override
  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized BulkUpdateCommandSet');
    this._webAbsoluteUrl = this.context.pageContext.web.absoluteUrl;
    console.log(`WebAbsolute Url => ${this._webAbsoluteUrl}`);
    this._documentLibServerRelativeUrl = this.context.pageContext.list.serverRelativeUrl;
    console.log(`Document Library ServerRelative Url => ${this._documentLibServerRelativeUrl}`);
    this._documentLibTitle = this.context.pageContext.list.title;
    console.log(`Document Library Title => ${this._documentLibTitle}`);
    this._documentLibIdObj = this.context.pageContext.list.id;
    this._documentLibId = this._documentLibIdObj._guid;
    console.log(`Document Library Id => ${this._documentLibId}`);
    this._userEmail = this.context.pageContext.user.email;
    console.log(`User Email => ${this._userEmail}`);
    this._spHttpClient = this.context.spHttpClient;
    await this.Initiate();
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    // Check for permissions
    let permission = new SPPermission(this.context.pageContext.web.permissions.value);
    const fullcontrol = permission.hasPermission(SPPermission.fullMask);
    console.log(`FullControl permission => ${fullcontrol}`);
    const approveItems = permission.hasPermission(SPPermission.approveItems);
    console.log(`Design permission => ${approveItems}`);
    let currentExportPermission = this.context.pageContext.web.permissions.hasAnyPermissions(SPPermission.fullMask, SPPermission.approveItems);
    console.log(`Logged in user Permission => ${currentExportPermission}`);
    let isSiteProjekt: boolean = false;
    if (this._webAbsoluteUrl.indexOf('Site') !== -1) {
      isSiteProjekt = true;
    }
    console.log(`Is Site Projekt => ${isSiteProjekt}`);
    let externalGuestUser = this.context.pageContext.user.isExternalGuestUser;
    console.log(`External User => ${externalGuestUser}`);
    // Bulk Update Command
    let currentBulkUpdatePermission = this.context.pageContext.web.permissions.hasPermission(SPPermission.editListItems);
    console.log(`Logged in user Bulk Update Permission => ${currentBulkUpdatePermission}`);
    const bulkUpdateCommand: Command = this.tryGetCommand(COMMANDBULKUPDATE);
    if (bulkUpdateCommand && currentBulkUpdatePermission && isSiteProjekt && (this._documentLibTitle.toLocaleLowerCase() == documentTitle)) {
      // Show the Bulk Update button only if its via Site Web and the logged in internal user has approve item permission.
      bulkUpdateCommand.visible = true;
      console.info('Valid internal user having design or higher permissions & accesing within Site Project.');
    }
    else {
      // Hide the Bulk Update as user is no longer valid.
      bulkUpdateCommand.visible = false;
      console.info('Invalid user not having the viewing permission or not a single document selected for the action.');
    }
  }

  // Relevant send Method
  public async send(): Promise<void> {
    switch (this._triggerType) {
      // Bulk Update
      case BulkUpdate:
        if (!!uploadIds && uploadIds.length > 0 && (!!textAreaValue && textAreaValue != '')) {
          console.log(`Trigerring the bulk update event => ${this._triggerType}`);
          this._allValidMaterials = uniqueTitle.length == validCollection.length ? true : false;
          this._updateDone = false;
          var batchuploadIds = [...uploadIds];
          console.log(`Batch Id for Doc Object length ${batchuploadIds.length} & the values are ${batchuploadIds}`);
          while (batchuploadIds.length > 0) {
            // Maxm 100 Ids can be updated for the given Docs.
            var batchuploadIdsObj = batchuploadIds.splice(0, 100);
            console.log(`Batch Id for Doc Object length ${batchuploadIdsObj.length} & the values are ${batchuploadIdsObj}`);
            if (batchuploadIds.length == 0) {
              this._updateDone = true;
            }
            dialog.close();
            await this._updateMultipleBatchRequest(this._triggerType, batchuploadIdsObj, validCollection.map((doc: { ID: string; }) => doc.ID.toString()));
          }
        }
        break;
      default:
        throw new Error('Unknown send action.');
    }
  }

  // Relevant Validate Method for Multi Bulk Update
  public async validate(): Promise<void> {
    if (!!uploadIds && uploadIds.length > 0) {
      validCollection = [];
      console.log(`Trigerring the bulk update event => ${this._triggerType}`);
      var multiLineTextBox = document.querySelector('textarea');
      if (!!multiLineTextBox && multiLineTextBox.textLength > 0) {
        let textAreaTextContent = multiLineTextBox.value;
        const materialValuesArr = textAreaTextContent.split(';');
        var sanitizedArray: string[] = [];
        materialValuesArr.map((materialValue) => {
          let trimMaterialValue = !!materialValue ? materialValue.trim() : materialValue;
          if (trimMaterialValue !== '') {
            sanitizedArray.push(trimMaterialValue);
          }
        });
        console.log(`Number of sanitizedArray are => ${sanitizedArray.length}`);
        uniqueTitle = sanitizedArray
          .map((e, i, final) => final.indexOf(e) === i && i)
          .filter(obj => sanitizedArray[obj])
          .map(e => sanitizedArray[e]);
        console.log(`Number of Unique Materials title are => ${uniqueTitle.length} that needs to be inserted`);

        var duplicateTitle = sanitizedArray
          .map((e, i, final) => final.indexOf(e) !== i && i)
          .filter(obj => sanitizedArray[obj])
          .map(e => sanitizedArray[e]);
        console.log(`Number of Duplicate Materials title are => ${duplicateTitle.length} that won't be inserted`);
        var batchedArray = [...uniqueTitle];
        var batched: string[];
        var validValues: any[] = [];
        while (batchedArray.length > 0) {
          // Maxm 500 values can be retrieved with IN operator.
          batched = batchedArray.splice(0, 500);
          console.log(`Number of Materials batched together => ${batched.length}.`);
          var lookupValues = await SharePointServices.retrieveLookUpValues(this._materialsListID, this._webAbsoluteUrl, this._spHttpClient, batched);
          console.log(`Number of lookupValues retrieved => ${lookupValues.Row.length}.`);
          if (!!lookupValues && !!lookupValues.Row && lookupValues.Row.length > 0) {
            validValues = validValues.concat(lookupValues.Row);
          }
        }
        if (!!validValues && validValues.length > 0) {
          validValues.map((mat: any) => {
            let Obj = { ID: mat.ID.toString(), Title: mat.Title.toString() };
            validCollection.push(Obj);
          });
          console.log(`Total Number of ${validCollection.length} lookup Title that would be inserted.`);
        }
        dialog.close();
        // Only Valid sent for Update
        if (!!validCollection && validCollection.length > 0) {
          textAreaValue = validCollection.map((valid) => { return valid.Title.toString(); }).join(';');
          let invalidCollection = uniqueTitle.filter(unique => !validCollection.some(valid => unique === valid.Title));
          console.log(`Total number of Invalid  are => ${invalidCollection.length}.`);
          if (!!invalidCollection && invalidCollection.length > 0) {
            invalidMaterials = invalidCollection.map((valid) => { return valid.toString(); }).join(';');
            message = 'Please see below valid and invalid materials, only valid ones would be assigned.';
            dialog = new HtmlDialog(this.send, title, message, cancel, update, null, null, false, { backgroundColor: '#0078d4' }, { width: 'inherit', height: '100px' }, textAreaValue, invalidMaterials, { width: 'inherit', height: '100px' }, true, true);
          }
          else {
            message = 'All materials are valid, would be assigned.';
            dialog = new HtmlDialog(this.send, title, message, cancel, update, null, null, false, { backgroundColor: '#0078d4' }, { width: 'inherit', height: '100px' }, textAreaValue, invalidMaterials, { display: 'none' }, true, true);
          }
          dialog.show();
        }
        // Not a single valid Materials
        else {
          title = `${this._triggerType} not possible!`;
          message = `All materials entered are invalid, please only enter valid materials.`;
          dialog = new HtmlDialog(null, title, message, ok, null, null, null, false, { display: 'none' }, { display: 'none' }, null, null, { display: 'none' });
          dialog.show();
        }
      }
    }
  }

  @override
  public async onExecute(event: IListViewCommandSetExecuteEventParameters): Promise<void> {
    switch (event.itemId) {
      case COMMANDBULKUPDATE:
        // Send selected documents for Bulk Update
        var selectedFileObject: any[] = [];
        var allSelectedFileObject: any[] = [];
        this._triggerType = BulkUpdate;
        if (event.selectedRows.length > 0) {
          selectedFileObject = await this.extractNestedFileIDs(event);
          if (!!selectedFileObject && selectedFileObject.length > 0) {
            uploadIds = selectedFileObject.map(doc => doc.ID);
            console.log(`Selected File IDs from Bulk Update => ${uploadIds}`);
            title = `${this._triggerType}`;
            message = 'Please enter valid materials, separated by semicolons, for bulk update.';
            textAreaValue = '';
            dialog = new HtmlDialog(this.validate, title, message, cancel, validate, null, null, false, { backgroundColor: '#0078d4' }, { width: 'inherit', height: '100px' }, textAreaValue, invalidMaterials, { display: 'none' }, false, true);
            dialog.show();
          }
          // No docs underneath the selected Folder
          else {
            title = `${this._triggerType} not possible!`;
            message = 'There are no documents for this action in the selected folder.';
            dialog = new HtmlDialog(null, title, message, ok, null, null, null, false, { display: 'none' }, { display: 'none' }, null, null, { display: 'none' });
            dialog.show();
          }
        }
        // If not selected send all the documents for Bulk Update
        else {
          allSelectedFileObject = await this.retrieveDefaultVisibleIDs();
          // Only if the documents are present
          if (!!allSelectedFileObject && allSelectedFileObject.length > 0) {
            console.log(`Total selected Doc IDs for Bulk Update are => ${allSelectedFileObject.length}`);
            uploadIds = allSelectedFileObject.map(doc => doc.ID);
            console.log(`Selected File IDs from Bulk Update => ${uploadIds}`);
            // Building the Dialog from HtmlDialog
            title = `${this._triggerType}`;
            message = 'Do you really want to update the metadata for all the documents?';
            dialog = new HtmlDialog(this.validate, title, message, cancel, validate, null, null, false, { backgroundColor: '#0078d4' }, { width: 'inherit', height: '100px' }, textAreaValue, invalidMaterials, { display: 'none' }, false, true);
            dialog.show();
          }
          // Not a single document present..
          else {
            title = `${this._triggerType} not possible!`;
            message = 'There are no documents for this action in the selected folder.';
            dialog = new HtmlDialog(null, title, message, ok, null, null, null, false, { display: 'none' }, { display: 'none' }, null, null, { display: 'none' });
            dialog.show();
          }
        }
        break;
      default:
        throw new Error('Unknown command.');
    }
  }

  private async retrieveDefaultVisibleIDs(): Promise<string[]> {
    var allDocIDs: any[] = [];
    if (window.location.href.indexOf('?id=') !== -1) {
      this._fileRefRelativeUrl = decodeURIComponent(window.location.href.split('?id=')[1].split('&')[0]);
    }
    else if (window.location.href.indexOf('&id=') !== -1) {
      this._fileRefRelativeUrl = decodeURIComponent(window.location.href.split('&id=')[1].split('&')[0]);
    }
    else {
      this._fileRefRelativeUrl = this._documentLibServerRelativeUrl;
    }
    console.log(`FileRef RelativeUrl => ${this._fileRefRelativeUrl}`);
    do {
      var batchFileIds = await SharePointServices.retrieveFileProps(this._documentLibId, this._webAbsoluteUrl, this._spHttpClient, this._fileRefRelativeUrl, !!batchFileIds ? batchFileIds.NextHref : null);
      if (!!batchFileIds && !!batchFileIds.Row && batchFileIds.Row.length > 0) {
        allDocIDs = allDocIDs.concat(batchFileIds.Row);
      }
    } while (!!batchFileIds && (batchFileIds.NextHref != null));
    return allDocIDs;
  }

  private async extractNestedFileIDs(event: IListViewCommandSetExecuteEventParameters): Promise<any[]> {
    var allSelectedFileIDs: any[] = [];
    for (var selectedItem of event.selectedRows) {
      var rowSelectedId = selectedItem.getValueByName('ID');
      var rowSelectedObjType = selectedItem.getValueByName('FSObjType');
      var rowSelectedFileRef = selectedItem.getValueByName('FileRef');
      var rowSelectedFileType = selectedItem.getValueByName('File_x0020_Type');
      // Retrieve Ids for all the folder types
      if (rowSelectedObjType == 1) {
        // dont just await, generate the all the selected docs ids in parallel
        var recursiveFileIds: any[] = [];
        do {
          var batchFileIds = await SharePointServices.retrieveFileProps(this._documentLibId, this._webAbsoluteUrl, this._spHttpClient, rowSelectedFileRef, !!batchFileIds ? batchFileIds.NextHref : null);
          if (!!batchFileIds && !!batchFileIds.Row && batchFileIds.Row.length > 0) {
            recursiveFileIds = recursiveFileIds.concat(batchFileIds.Row);
          }
        } while (!!batchFileIds && (batchFileIds.NextHref != null));
        recursiveFileIds.map((file) => {
          var fileMetadata = { ID: file.ID, FileExtens: file.File_x0020_Type };
          allSelectedFileIDs.push(fileMetadata);
        });
      }
      else {
        var fileObj = { ID: rowSelectedId, FileExtens: rowSelectedFileType };
        allSelectedFileIDs.push(fileObj);
      }
    }
    console.log(`Total Number of selected File IDs length ${allSelectedFileIDs.length}`);
    return allSelectedFileIDs;
  }

  private async _updateMultipleBatchRequest(updateField: string, itemsIDUpdate: string | any[], valuesToUpdate: any[]): Promise<boolean> {
    var libraryId = this._documentLibId;
    // generate a batch boundary
    var batchGuid = this.generateGUID();

    // creating the body
    var batchContents = new Array();
    var changeSetId = this.generateGUID();

    // for each item...
    for (var index = 0; index < itemsIDUpdate.length; index++) {
      var itemID = itemsIDUpdate[index];
      var data = null;
      if (updateField == BulkUpdate) {
        data = {
          '__metadata': { 'type': 'SP.Data.Shared_x0020_DocumentsItem' },
          'PlantMaterialId': {
            '__metadata': { 'type': 'Collection(Edm.Int32)' },
            'results': valuesToUpdate
          }
        };
      }
      var endpoint = this._webAbsoluteUrl + `/_api/web/lists(guid'` + libraryId + `')` + `/items(` + itemID + `)`;
      // create the changeset
      batchContents.push('--changeset_' + changeSetId);
      batchContents.push('Content-Type: application/http');
      batchContents.push('Content-Transfer-Encoding: binary');
      batchContents.push('');
      batchContents.push('PATCH ' + endpoint + ' HTTP/1.1');
      batchContents.push('Content-Type: application/json;odata=verbose');
      batchContents.push('Accept: application/json;odata=verbose');
      batchContents.push('If-Match: *');
      batchContents.push('');
      batchContents.push(JSON.stringify(data));
      batchContents.push('');
    }
    // END changeset to create data
    batchContents.push('--changeset_' + changeSetId + '--');
    // generate the body of the batch
    var batchBody = batchContents.join('\r\n');
    // start with a clean array
    batchContents = new Array();

    // create batch for creating items
    batchContents.push('--batch_' + batchGuid);
    batchContents.push('Content-Type: multipart/mixed; boundary="changeset_' + changeSetId + '"');
    batchContents.push('Content-Length: ' + batchBody.length);
    batchContents.push('Content-Transfer-Encoding: binary');
    batchContents.push('');
    batchContents.push(batchBody);
    batchContents.push('');
    batchContents.push('--batch_' + batchGuid);
    // batch end
    batchBody = batchContents.join('\r\n');
    // create the request endpoint 
    var urlEndpoint = this._webAbsoluteUrl + '/_api/$batch';

    const spHttpClientPostOptions: ISPHttpClientOptions = {
      headers: {
        'odata-version': '3.0',
        'Content-Type': 'multipart/mixed; boundary="batch_' + batchGuid + '"'
      },
      body: batchBody
    };
    return this.context.spHttpClient.post(urlEndpoint, SPHttpClient.configurations.v1, spHttpClientPostOptions)
      .then(async (response: SPHttpClientResponse) => {
        return response.text()
          .then((batchResponse) => {
            console.log(`batchResponse => ${batchResponse}`);
            if (response.ok && (response.status === 200)) {
              console.log(`success => ${response}`);
              // File is locked cant be updated for Bulk Update
              if ((this._triggerType == BulkUpdate) && (batchResponse.indexOf('odata.error') !== -1) && (batchResponse.indexOf('locked') !== -1)) {
                title = `${this._triggerType} nicht möglich!`;
                message = `One or more documents are open. Please close them and repeat the process.`;
                dialog = new HtmlDialog(null, title, message, ok, null, null, null, false, { display: 'none' }, { display: 'none' }, null, null, { display: 'none' });
                dialog.show();
                return false;
              }
              // All the input Material are valid ones & updated successfully
              else if ((this._triggerType == BulkUpdate) && this._allValidMaterials && this._updateDone) {
                title = `${this._triggerType} successful!`;
                message = `All documents have been successfully updated. The display is updated automatically with a time delay. Press <F5> for an immediate update.`;
                dialog = new HtmlDialog(null, title, message, ok, null, null, null, false, { display: 'none' }, { display: 'none' }, null, null, { display: 'none' });
                dialog.show();
              }
              // Only the valid input Materials are updated successfully
              else if ((this._triggerType == BulkUpdate) && !this._allValidMaterials && this._updateDone) {
                title = `${this._triggerType} successful!`;
                message = `Only valid materials have been updated. The display is updated automatically with a time delay. Press <F5> for an immediate update..`;
                dialog = new HtmlDialog(null, title, message, ok, null, null, null, false, { display: 'none' }, { display: 'none' }, null, null, { display: 'none' });
                dialog.show();
              }
              return response.ok;
            }
            else {
              console.info('http Reponse follows ...');
              console.log(`There is a problem while doing the bulk Update for the items => ${response}`);
              if (response.status === 403) {
                title = `${this._triggerType} fehlgeschlagen!`;
                message = `The currently logged in session is invalid and has expired. Press <F5> and repeat the process again. Error:- Code: ${response.status} Message: ${response.statusText}`;
                dialog = new HtmlDialog(null, title, message, ok, null, null, null, false, { display: 'none' }, { display: 'none' }, null, null, { display: 'none' });
                dialog.show();
              }
              else {
                title = `${this._triggerType} fehlgeschlagen!`;
                message = `Please contact us and provide the error code shown below. Error:- Code: ${response.status} Message: ${response.statusText}`;
                dialog = new HtmlDialog(null, title, message, ok, null, null, null, false, { display: 'none' }, { display: 'none' }, null, null, { display: 'none' });
                dialog.show();
              }
              return response.ok;
            }
          })
          .catch((error: any) => {
            console.error(`An error occurred on the server. Error message => ${error}`);
            title = `${this._triggerType} fehlgeschlagen!`;
            message = `Please contact us and provide the error code shown below. Error Message: ${error}`;
            dialog = new HtmlDialog(null, title, message, ok, null, null, null, false, { display: 'none' }, { display: 'none' }, null, null, { display: 'none' });
            dialog.show();
            return false;
          });
      });
  }

  private async Initiate() {
    // Retrieve Pruefungen List details
    let allListDetails = await SharePointServices.getAllListDetails(this._webAbsoluteUrl, this._spHttpClient);
    this._materialsListID = !!allListDetails ? allListDetails.filter((material: { EntityTypeName: string; }) => { return material.EntityTypeName.toLocaleLowerCase() == materialsListInternalName; })[0].Id : '';
    console.log(`Pruefungen List ID => ${this._materialsListID}`);
  }

  private generateGUID() {
    var d = new Date().getTime();
    var uuid = 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => {
      var r = (d + Math.random() * 16) % 16 | 0;
      d = Math.floor(d / 16);
      return (c == 'x' ? r : (r & 0x7 | 0x8)).toString(16);
    });
    return uuid;
  }

}