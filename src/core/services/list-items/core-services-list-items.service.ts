import { default as pnp } from 'sp-pnp-js';
import { SpCore } from '../setup/core-services-setup.service';
import { IAttachmentData, IAttachmentServerData } from './core-services-list-items.interface';
import { AttachmentIcon } from 'gl-w-frontend';

class Attachment {
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
}
export const ListItemsAttachment = new Attachment();
