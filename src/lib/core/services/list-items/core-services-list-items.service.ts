import pnp, { AttachmentFileAddResult, AttachmentFileInfo, ItemAddResult, TypedHash } from 'sp-pnp-js';
// noinspection ES6PreferShortImport
import { SpCore } from '../setup/core-services-setup.service';
import { IAttachmentData, IAttachmentServerData, IListDatabaseResults } from './core-services-list-items.interface';
import { AttachmentFile } from 'sp-pnp-js/lib/sharepoint/attachmentfiles';
import { AttachmentIcon } from 'gl-w-frontend/lib/es5/scripts/core/services/attachment/core-services-attachment.service';

class Core {

  fieldsToStringArray(fields: string[]): string {
    return fields.toString().replace('[', '').replace(']', '');
  }

  /**
   * Retrieve a search list based on fields, filter and ordering
   * @param listName
   * @param fieldsArray (optional)
   */
  async retrieve(listName: string, fieldsArray: string[] = []): Promise<any[]> {
    const fields: string = this.fieldsToStringArray(fieldsArray);

    try {
      return await pnp.sp.web.lists.getByTitle(listName).items.select(fields).get();
    } catch (reason) {
      const error = SpCore.onError(reason)
      SpCore.showErrorLog(reason);
      throw new Error(error.code.toString())
    }
  }

  /**
   * Retrieve a unique list item based on fields, filter and ordering
   * @param listItemId
   * @param listName
   * @param fieldsArray (optional)
   */
  async retrieveSingle(listItemId: number, listName: string, fieldsArray: string[] = []): Promise<any> {
    const fields: string = this.fieldsToStringArray(fieldsArray);

    try {
      return await pnp.sp.web.lists.getByTitle(listName).items.getById(listItemId).select(fields).get();
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
   */
  recycle(listItemId: number, listName: string) {
    return new Promise((resolve, reject) => {
      pnp.sp.web.lists.getByTitle(listName).items.getById(listItemId).recycle().then(() => {
        resolve({ code: 200, description: 'Success!', message: 'The item has been moved to recycle bin.' })
      })
      .catch(reason => {
        const error = SpCore.showErrorLog(reason);
        reject(error)
      });
    });
  }

  /**
   * Delete items permanently. The user will not be able to restore the information.
   * @param listItemId
   * @param listName
   */
  delete(listItemId: number, listName: string) {
    return new Promise((resolve, reject) => {
      pnp.sp.web.lists.getByTitle(listName).items.getById(listItemId).delete().then(() => {
        resolve({ code: 200, description: 'Success!', message: 'The item has been deleted successfully.' })
      })
      .catch(reason => {
        const error = SpCore.showErrorLog(reason);
        reject(error)
      });
    });
  }

  /**
   * Adds a new item to the collection
   * @param listName
   * @param data
   */
  add(listName: string, data:  TypedHash<any>): Promise<ItemAddResult> {
    return new Promise((resolve, reject) => {
      pnp.sp.web.lists.getByTitle(listName).items.add(data)
      .then((iar: ItemAddResult) => { resolve(iar); })
      .catch(reason => {
        const error = SpCore.showErrorLog(reason);
        reject(error)
      })
    })
  }

  /**
   * Update a new item to the collection
   * @param listItemId
   * @param listName
   * @param data
   */
  update(listItemId: number, listName: string, data:  TypedHash<any>): Promise<ItemAddResult> {
    return new Promise((resolve, reject) => {
      pnp.sp.web.lists.getByTitle(listName).items.getById(listItemId).update(data)
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
   */
  add(listItemId: number, listName: string, attachments: AttachmentFileInfo[]): Promise<any> {
    return new Promise((resolve, reject) => {
      pnp.sp.web.lists.getByTitle(listName).items.getById(listItemId).attachmentFiles.addMultiple(attachments)
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
   */
  delete(listItemId: number, listName: string, attachments: string[]): Promise<any> {
    return new Promise((resolve, reject) => {
      pnp.sp.web.lists.getByTitle(listName).items.getById(listItemId).attachmentFiles.deleteMultiple(...attachments)
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
   */
  retrieve(listItemId: number, listName: string): Promise<IAttachmentServerData[]> {
    return new Promise((resolve, reject) => {
      pnp.sp.web.lists.getByTitle(listName).items.getById(listItemId).attachmentFiles.get()
        .then((attachments: IAttachmentServerData[]) => {
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
   * This method is intended to retrieve all attachments inside a list item and prepare the list to use directly to the user interface (binding)
   * @param listItemId
   * @param listName
   */
  async retrieveForBinding(listItemId: number, listName: string): Promise<IAttachmentData[]> {
    const attachmentList: IAttachmentServerData[] = await this.retrieve(listItemId, listName);
    const attachments: IAttachmentData[] = [];

    for (const attachment of attachmentList) {
      const host = attachment.__metadata.uri.split('/s/')[0];
      const serverRelativeUrl = attachment.ServerRelativeUrl;
      const url = encodeURI(host + serverRelativeUrl);

      const file: IAttachmentData = {
        id: attachmentList.length,
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
   */
  retrieveTxtContent(listItemId: number, listName: string, fileName: string = 'attachment'): Promise<string> {
    return new Promise((resolve, reject) => {
      pnp.sp.web.lists.getByTitle(listName).items.getById(listItemId).attachmentFiles.getByName(`${fileName}.txt`).getText()
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
   */
  setTxtContent(listItemId: number, listName: string, fileName: string = 'attachment', content: string): Promise<AttachmentFile> {
    return new Promise((resolve, reject) => {
      pnp.sp.web.lists.getByTitle(listName).items.getById(listItemId).attachmentFiles.getByName(`${fileName}.txt`).setContent(content)
      .then((attachmentFile: AttachmentFile) => { resolve(attachmentFile); })
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
   */
  addTxtContent(listItemId: number, listName: string, fileName: string = 'attachment', content: string): Promise<AttachmentFileAddResult> {
    return new Promise((resolve, reject) => {
      pnp.sp.web.lists.getByTitle(listName).items.getById(listItemId).attachmentFiles.add(`${fileName}.txt`, content)
      .then((attachmentFile: AttachmentFileAddResult) => { resolve(attachmentFile); })
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
   */
  retrieveSearch(listName: string, fieldsArray: string[] = [], filter: string = '', orderBy: string = 'ID', ascending: boolean = true): Promise<IListDatabaseResults[]> {
    const fields: string = fieldsArray.toString().replace('[', '').replace(']', '');

    return new Promise(async (resolve, reject) => {
      pnp.sp.web.lists.getByTitle(listName).items
        .orderBy(orderBy, ascending)
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
