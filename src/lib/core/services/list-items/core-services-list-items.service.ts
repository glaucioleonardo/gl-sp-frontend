import { SpCore } from '../setup/core-services-setup.service';
import { IAttachment, IAttachmentAddResult, IAttachmentFileInfo, IAttachmentInfo } from '@pnp/sp/attachments/index';
import { IAttachmentBlob, IAttachmentData, IAttachmentMultipleBlobs, IListDatabaseResults, ItemAddResult } from './core-services-list-items.interface';
import { IItem, sp } from '@pnp/sp/presets/core';

import "@pnp/sp/attachments";
import { ITypedHash } from '@pnp/common';
import { ISpCoreResult } from '../setup/core-services-setup.interface';
import { IComboBoxData } from 'gl-w-combobox-frontend';
import { ArrayRemove } from 'gl-w-array-frontend';
import { AttachmentIcon } from 'gl-w-attachment-frontend';

class Core {

  fieldsToStringArray(fields: string[]): string {
    return fields.toString().replace('[', '').replace(']', '');
  }

  /**
   * Retrieve a search list based on fields, filter and ordering
   * @param listName
   * @param fieldsArray (optional)
   * @param baseUrl (optional) In case it is necessary to gather data from another url.
   */
  async retrieve(listName: string, fieldsArray: string[] = [], baseUrl?: string): Promise<any[]> {
    const fields: string = this.fieldsToStringArray(fieldsArray);
    const base = baseUrl == null ? SpCore.baseUrl : baseUrl;

    try {
      return await sp.configure(SpCore.config, base).web.lists.getByTitle(listName).items.select(fields).getAll();
    } catch (reason) {
      const error = SpCore.onError(reason)
      SpCore.showErrorLog(reason);
      throw new Error(error.code.toString())
    }
  }

  /**
   * Retrieve a search list based on fields from external sharepoint, filter and ordering
   * @param listName
   * @param fieldsArray
   * @param baseUrl: Necessary to gather data from another url.
   * @param top: Retrieve the number of first n items.
   */
  async retrieveExternal(listName: string, fieldsArray: string[] = [], baseUrl: string, top: number = 0): Promise<any[]> {
    const fields: string = this.fieldsToStringArray(fieldsArray);

    try {
      const listUrl: string = encodeURI(`${baseUrl}/_api/web/lists/GetByTitle('${listName}')?$select=ItemCount`);

      let count: number = top;

      if (top == null || top === 0) {
        const result = await fetch(listUrl, await SpCore.fetchHeader());
        const list = await result.json();
        count = await list.d.ItemCount;
      }

      const itemsUrl: string = encodeURI(`${baseUrl}/_api/web/lists/GetByTitle('${listName}')/items?$select=${fields}&$top=${count}`);
      const resultItems = await fetch(itemsUrl, await SpCore.fetchHeader());
      const fetchItems = await resultItems.json();
      return fetchItems.d.results;
    } catch (reason) {
      SpCore.showErrorLog(reason);
      return [];
    }
  }

  /**
   * Retrieve a single search list item based on fields from external sharepoint, filter and ordering
   * @param id
   * @param listName
   * @param fieldsArray
   * @param baseUrl: Necessary to gather data from another url.
   */
  async retrieveExternalSingle(id: number, listName: string, fieldsArray: string[] = [], baseUrl: string): Promise<any[]> {
    const fields: string = this.fieldsToStringArray(fieldsArray);

    try {
      const itemsUrl: string = encodeURI(`${baseUrl}/_api/web/lists/GetByTitle('${listName}')/items(${id})?$select=${fields}`);
      const resultItems = await fetch(itemsUrl, await SpCore.fetchHeader());
      const fetchItems = await resultItems.json();
      return fetchItems.d;
    } catch (reason) {
      SpCore.showErrorLog(reason);
      return [];
    }
  }

  /**
   * Retrieve a search list based on fields from external sharepoint, filter and ordering
   * @param listName
   * @param baseUrl: Necessary to gather data from another url.
   * @param valueField: Value of combobox.
   * @param textField: Inner value of combobox.
   */
  async retrieveForCombobox(listName: string, baseUrl: string, valueField: string = 'value', textField: string = 'text'): Promise<IComboBoxData[]> {
    try {
      const items: IComboBoxData[] = [];

      if (listName != null && listName.length > 0) {
        const fields: string[] = [valueField, textField];
        try {
          const itemsResult = await this.retrieve(listName, fields, baseUrl);
          for (const item of itemsResult) {
            items.push({
              text: item[textField],
              value: item[valueField].toString()
            });
          }
          return items;
        } catch (reason) {
          SpCore.showErrorLog(reason);
          return reason;
        }
      } else {
        return [];
      }

    } catch (reason) {
      SpCore.showErrorLog(reason);
      return reason;
    }
  }


  /**
   * Retrieve a search list based on fields from external sharepoint, filter and ordering
   * @param listName
   * @param baseUrl: Necessary to gather data from another url.
   * @param valueField: Value of combobox.
   * @param textField: Inner value of combobox.
   * @param top: Retrieve the number of first n items.
   */
  async retrieveExternalForCombobox(listName: string, baseUrl: string, valueField: string = 'value', textField: string = 'text', top: number = 0): Promise<IComboBoxData[]> {
    try {
      const fields = [valueField, textField];
      const listItems = await this.retrieveExternal(listName, fields, baseUrl, top);

      const comboBox: IComboBoxData[] = [];
      for (const item of listItems) {
        comboBox.push({
          text: item[textField],
          value: item[valueField].toString()
        })
      }

      return comboBox;

    } catch (reason) {
      SpCore.showErrorLog(reason);
      return reason;
    }
  }

  /**
   * Retrieve a unique list item based on fields, filter and ordering
   * @param listItemId
   * @param listName
   * @param fieldsArray (optional)
   * @param baseUrl (optional) In case it is necessary to gather data from another url.

   */
  async retrieveSingle(listItemId: number, listName: string, fieldsArray: string[] = [], baseUrl?: string): Promise<any> {
    const fields: string = this.fieldsToStringArray(fieldsArray);
    const base = baseUrl == null ? SpCore.baseUrl : baseUrl;

    try {
      return await sp.configure(SpCore.config, base).web.lists.getByTitle(listName).items.getById(listItemId).select(fields).get();
    } catch (reason) {
      const error = SpCore.onError(reason)
      SpCore.showErrorLog(reason);
      throw new Error(error.code.toString())
    }
  }

  /**
   * Move items to recycle bin. The user will be able to restore the information.
   * @param listItemId
   * @param listName
   * @param baseUrl (optional) In case it is necessary to gather data from another url.
   */
  recycle(listItemId: number, listName: string, baseUrl?: string) {
    const base = baseUrl == null ? SpCore.baseUrl : baseUrl;

    return new Promise((resolve, reject) => {
      sp.configure(SpCore.config, base).web.lists.getByTitle(listName).items.getById(listItemId).recycle().then(() => {
        resolve({ code: 200, description: 'Success!', message: 'The item has been moved to recycle bin.' })
      })
      .catch(reason => {
        const error = SpCore.showErrorLog(reason);
        reject(error)
      });
    });
  }

  /**
   * Move duplicated items to recycle bin. The user will be able to restore the information.
   * @param listName
   * @param field The comparison field
   * @param baseUrl (optional) In case it is necessary to gather data from another url.
   */
  async recycleDuplicated(listName: string, field: string,  baseUrl?: string): Promise<any[]> {
    const base = baseUrl == null ? SpCore.baseUrl : baseUrl;

    try {
      const data = await this.retrieve(listName, ['Id', field ], base);
      const duplicates = await ArrayRemove.notDuplicatedByKey(data, field, field, true);

      for (const duplicate of duplicates) {
        const items = data.filter(x => x[field] === duplicate);
        for (let i = 1; i < items.length; i++) {
          await this.recycle(items[i].Id, listName, base);
        }
      }

      return duplicates;
    } catch (reason) {
      SpCore.showErrorLog(reason);
      return [];
    }
  }

  /**
   * Delete items permanently. The user will not be able to restore the information.
   * @param listItemId
   * @param listName
   * @param baseUrl (optional) In case it is necessary to gather data from another url.
   */
  delete(listItemId: number, listName: string, baseUrl?: string) {
    const base = baseUrl == null ? SpCore.baseUrl : baseUrl;

    return new Promise((resolve, reject) => {
      sp.configure(SpCore.config, base).web.lists.getByTitle(listName).items.getById(listItemId).delete().then(() => {
        resolve({ code: 200, description: 'Success!', message: 'The item has been deleted successfully.' })
      })
      .catch(reason => {
        const error = SpCore.showErrorLog(reason);
        reject(error)
      });
    });
  }

  /**
   * Delete duplicated items permanently. The user will be able to restore the information.
   * @param listName
   * @param field The comparison field
   * @param baseUrl (optional) In case it is necessary to gather data from another url.
   */
  async deleteDuplicated(listName: string, field: string,  baseUrl?: string): Promise<any[]> {
    const base = baseUrl == null ? SpCore.baseUrl : baseUrl;

    try {
      const data = await this.retrieve(listName, ['Id', field ], base);
      const duplicates = await ArrayRemove.notDuplicatedByKey(data, field, field, true);

      for (const duplicate of duplicates) {
        const items = data.filter(x => x[field] === duplicate);
        for (let i = 1; i < items.length; i++) {
          await this.delete(items[i].Id, listName, base);
        }
      }

      return duplicates;
    } catch (reason) {
      SpCore.showErrorLog(reason);
      return [];
    }
  }

  /**
   * Adds a new item to the collection
   * @param listName
   * @param data
   * @param baseUrl (optional) In case it is necessary to gather data from another url.
   */
  async add(listName: string, data:  ITypedHash<any>, baseUrl?: string): Promise<ItemAddResult | ISpCoreResult> {
    const base = baseUrl == null ? SpCore.baseUrl : baseUrl;

    try {
      return await sp.configure(SpCore.config, base).web.lists.getByTitle(listName).items.add(data);
    } catch (reason) {
      SpCore.showErrorLog(reason)
      throw new Error(reason)
    }
  }

  /**
   * Update a new item to the collection
   * @param listItemId
   * @param listName
   * @param data
   * @param baseUrl (optional) In case it is necessary to gather data from another url.
   */
  update(listItemId: number, listName: string, data:  ITypedHash<any>, baseUrl?: string): Promise<ItemAddResult> {
    const base = baseUrl == null ? SpCore.baseUrl : baseUrl;

    return new Promise((resolve, reject) => {
      sp.configure(SpCore.config, base).web.lists.getByTitle(listName).items.getById(listItemId).update(data)
      .then((iar: ItemAddResult) => { resolve(iar); })
      .catch(reason => {
        const error = SpCore.showErrorLog(reason);
        reject(error)
      })
    })
  }
}
export const ListItemsCore = new Core();

class Attachment {
  /**
   * Before using this method, you need defining the base url on setup (SpCore).
   * Adds multiple new attachment to the collection. Not supported for batching.
   * @param listItemId
   * @param listName
   * @param attachments
   * @param baseUrl (optional) In case it is necessary to gather data from another url.
   */
  add(listItemId: number, listName: string, attachments: IAttachmentFileInfo[], baseUrl?: string): Promise<any> {
    const base = baseUrl == null ? SpCore.baseUrl : baseUrl;

    return new Promise((resolve, reject) => {
      sp.configure(SpCore.config, base).web.lists.getByTitle(listName).items.getById(listItemId).attachmentFiles.addMultiple(attachments)
        .then(() => {
          resolve({ code: 200, description: 'Success!', message: 'The attachments has been added successfully.' });
        })
        .catch(reason => {
          SpCore.showErrorLog(reason);
          reject([])
        });
    });
  }

  /**
   * Before using this method, you need defining the base url on setup (SpCore).
   * Delete multiple attachments from the collection. Not supported for batching.
   * @param listItemId
   * @param listName
   * @param attachments
   * @param baseUrl (optional) In case it is necessary to gather data from another url.
   */
  delete(listItemId: number, listName: string, attachments: string[], baseUrl?: string): Promise<any> {
    const base = baseUrl == null ? SpCore.baseUrl : baseUrl;

    return new Promise((resolve, reject) => {
      sp.configure(SpCore.config, base).web.lists.getByTitle(listName).items.getById(listItemId).attachmentFiles.deleteMultiple(...attachments)
        .then(() => {
          resolve({ code: 200, description: 'Success!', message: 'The attachments has been deleted successfully.' });
        })
        .catch(reason => {
          SpCore.showErrorLog(reason);
          reject([])
        });
    });
  }

  /**
   * Before using this method, you need defining the base url on setup (SpCore).
   * @param listItemId
   * @param listName
   * @param baseUrl (optional) In case it is necessary to gather data from another url.
   */
  retrieve(listItemId: number, listName: string, baseUrl?: string): Promise<IAttachmentInfo[]> {
    const base = baseUrl == null ? SpCore.baseUrl : baseUrl;

    return new Promise((resolve, reject) => {
      sp.configure(SpCore.config, base).web.lists.getByTitle(listName).items.getById(listItemId).attachmentFiles.get()
        .then((attachments: IAttachmentInfo[]) => {
          resolve(attachments);
        })
        .catch(reason => {
          SpCore.showErrorLog(reason);
          reject([])
        });
    });
  }

  /**
   * Before using this method, you need defining the base url on setup (SpCore).
   * @param listItemId
   * @param listName
   * @param fileName
   * @param baseUrl (optional) In case it is necessary to gather data from another url.
   */
  async retrieveBlob(listItemId: number, listName: string, fileName: string, baseUrl?: string): Promise<Blob | null> {
    const base = baseUrl == null ? SpCore.baseUrl : baseUrl;

    try {
      const item: IItem = sp.configure(SpCore.config, base).web.lists.getByTitle(listName).items.getById(listItemId);
      return await item.attachmentFiles.getByName(fileName).getBlob();
    } catch (reason) {
      SpCore.showErrorLog(reason);
      return null;
    }
  }

  async retrieveMultipleBlobSameItem(listItemId: number, listName: string, fileNames: string[], baseUrl?: string): Promise<IAttachmentBlob[]> {
    const base = baseUrl == null ? SpCore.baseUrl : baseUrl;
    const blobs: IAttachmentBlob[] = [];

    try {
      for (const fileName of fileNames) {
        const item: IItem = sp.configure(SpCore.config, base).web.lists.getByTitle(listName).items.getById(listItemId);
        blobs.push({
          fileName,
          file: await item.attachmentFiles.getByName(fileName).getBlob(),
          icon: AttachmentIcon.get(fileName)
        }) ;
      }

     return blobs;
    } catch (reason) {
      SpCore.showErrorLog(reason);
      return [];
    }
  }

  async retrieveMultipleBlob(data: IAttachmentMultipleBlobs[], baseUrl?: string): Promise<IAttachmentBlob[]> {
    const base = baseUrl == null ? SpCore.baseUrl : baseUrl;
    const blobs: IAttachmentBlob[] = [];

    try {
      for (const content of data) {
        const item: IItem = sp.configure(SpCore.config, base).web.lists.getByTitle(content.listName).items.getById(content.id);
        blobs.push({
          fileName: content.fileName,
          file: await item.attachmentFiles.getByName(content.fileName).getBlob(),
          icon: AttachmentIcon.get(content.fileName)
        }) ;
      }

      return blobs;
    } catch (reason) {
      SpCore.showErrorLog(reason);
      return [];
    }
  }

  /**
   * Before using this method, you need defining the base url on setup (SpCore).
   * This method is intended to retrieve all attachments inside a list item and prepare the list to use directly to the user interface (binding)
   * @param listItemId
   * @param listName
   * @param baseUrl (optional) In case it is necessary to gather data from another url.
   */
  async retrieveForBinding(listItemId: number, listName: string, baseUrl?: string): Promise<IAttachmentData[]> {
    let base = baseUrl == null ? SpCore.baseUrl : baseUrl;

    if (!base.includes('http')) {
      base  = 'http://' + base;
    }

    const attachmentList: IAttachmentInfo[] = await this.retrieve(listItemId, listName, base);
    const attachments: IAttachmentData[] = [];

    for (const attachment of attachmentList) {
      const host = base.split('://')[0] + '://' + base.split('://')[1].split('/')[0];
      const serverRelativeUrl = attachment.ServerRelativeUrl;
      const url = encodeURI(host + serverRelativeUrl);

      const file: IAttachmentData = {
        id: attachments.length,
        new: false,
        remove: false,
        name: attachment.FileName,
        url,
        icon: AttachmentIcon.get(attachment.FileName)
      };

      attachments.push(file);
    }

    return attachments;
  }

  /**
   * Before using this method, you need defining the base url on setup (SpCore).
   * @param listItemId
   * @param listName
   * @param fileName Without extension, e.g. "attachment"
   * @param baseUrl (optional) In case it is necessary to gather data from another url.
   */
  retrieveTxtContent(listItemId: number, listName: string, fileName: string = 'attachment', baseUrl?: string): Promise<string> {
    const base = baseUrl == null ? SpCore.baseUrl : baseUrl;

    return new Promise((resolve, reject) => {
      sp.configure(SpCore.config, base).web.lists.getByTitle(listName).items.getById(listItemId).attachmentFiles.getByName(`${fileName}.txt`).getText()
      .then((image: string) => {
        resolve(image);
      })
      .catch(reason => {
        SpCore.showErrorLog(reason);
        reject('');
      });
    })
  }

  /**
   * Sets the content of a file. Not supported for batching
   * @param listItemId
   * @param listName
   * @param fileName Without extension, e.g. "attachment"
   * @param content Content to be added to the text
   * @param baseUrl (optional) In case it is necessary to gather data from another url.
   */
  setTxtContent(listItemId: number, listName: string, fileName: string = 'attachment', content: string, baseUrl?: string): Promise<IAttachment> {
    const base = baseUrl == null ? SpCore.baseUrl : baseUrl;

    return new Promise((resolve, reject) => {
      sp.configure(SpCore.config, base).web.lists.getByTitle(listName).items.getById(listItemId).attachmentFiles.getByName(`${fileName}.txt`).setContent(content)
      .then((attachmentFile: IAttachment) => { resolve(attachmentFile); })
      .catch(reason => {
        const error = SpCore.showErrorLog(reason);
        reject(error)
      })
    })
  }

  /**
   * Adds a new attachment to the collection. Not supported for batching.
   * @param listItemId
   * @param listName
   * @param fileName Without extension, e.g. "attachment"
   * @param content Content to be added to the text
   * @param baseUrl (optional) In case it is necessary to gather data from another url.
   */
  addTxtContent(listItemId: number, listName: string, fileName: string = 'attachment', content: string, baseUrl?: string): Promise<IAttachmentAddResult> {
    const base = baseUrl == null ? SpCore.baseUrl : baseUrl;

    return new Promise((resolve, reject) => {
      sp.configure(SpCore.config, base).web.lists.getByTitle(listName).items.getById(listItemId).attachmentFiles.add(`${fileName}.txt`, content)
      .then((attachmentFile: IAttachmentAddResult) => { resolve(attachmentFile); })
      .catch(reason => {
        const error = SpCore.showErrorLog(reason);
        reject(error)
      })
    })
  }
}
export const ListItemsAttachment = new Attachment();

class Search {
  /**
   * Retrieve a search list based on fields, filter and ordering
   * @param listName
   * @param fieldsArray (optional)
   * @param filter (optional)
   * @param orderBy (optional)
   * @param ascending (optional)
   * @param baseUrl (optional) In case it is necessary to gather data from another url.
   */
  retrieveSearch(listName: string, fieldsArray: string[] = [], filter: string = '', orderBy: string = 'ID', ascending: boolean = true, baseUrl?: string): Promise<IListDatabaseResults[]> {
    const base = baseUrl == null ? SpCore.baseUrl : baseUrl;
    const fields: string = fieldsArray.toString().replace('[', '').replace(']', '');

    return new Promise(async (resolve, reject) => {
      sp.configure(SpCore.config, base).web.lists.getByTitle(listName).items
        .orderBy(orderBy, ascending)
        .select(fields)
        .filter(filter)
        .getAll()
        .then((result: any[]) => {
          resolve(result);
        })
        .catch(reason => {
          SpCore.showErrorLog(reason);
          reject(reason);
        });
    });
  }

  /**
   * Retrieve a search list based on fields, filter and ordering
   * @param listName
   * @param fieldsArray (optional)
   * @param maxItems (optional)
   * @param filter (optional)
   * @param orderBy (optional)
   * @param ascending (optional)
   * @param baseUrl (optional) In case it is necessary to gather data from another url.
   */
  retrieveSearchLimited(listName: string, fieldsArray: string[] = [], maxItems: number = 100 ,filter: string = '', orderBy: string = 'ID', ascending: boolean = true, baseUrl?: string): Promise<IListDatabaseResults[]> {
    const base = baseUrl == null ? SpCore.baseUrl : baseUrl;
    const fields: string = fieldsArray.toString().replace('[', '').replace(']', '');

    return new Promise(async (resolve, reject) => {
      sp.configure(SpCore.config, base).web.lists.getByTitle(listName).items
        .orderBy(orderBy, ascending)
        .top(maxItems)
        .select(fields)
        .filter(filter)
        .get()
        .then((result: any[]) => {
          resolve(result);
        })
        .catch(reason => {
          SpCore.showErrorLog(reason);
          reject(reason);
        });
    });
  }
}
export const ListItemsSearch = new Search();
