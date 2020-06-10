var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
import pnp from 'sp-pnp-js';
import { SpCore } from '../setup/core-services-setup.service';
import { AttachmentIcon } from 'gl-w-frontend/lib/es5/scripts/core/services/attachment/core-services-attachment.service';
class Core {
    fieldsToStringArray(fields) {
        return fields.toString().replace('[', '').replace(']', '');
    }
    retrieve(listName, fieldsArray = []) {
        return __awaiter(this, void 0, void 0, function* () {
            const fields = this.fieldsToStringArray(fieldsArray);
            try {
                return yield pnp.sp.web.lists.getByTitle(listName).items.select(fields).get();
            }
            catch (reason) {
                const error = SpCore.onError(reason);
                SpCore.showErrorLog(reason);
                throw new Error(error.code.toString());
            }
        });
    }
    retrieveSingle(listItemId, listName, fieldsArray = []) {
        return __awaiter(this, void 0, void 0, function* () {
            const fields = this.fieldsToStringArray(fieldsArray);
            try {
                return yield pnp.sp.web.lists.getByTitle(listName).items.getById(listItemId).select(fields).get();
            }
            catch (reason) {
                const error = SpCore.onError(reason);
                SpCore.showErrorLog(reason);
                throw new Error(error.code.toString());
            }
        });
    }
    recycle(listItemId, listName) {
        return new Promise((resolve, reject) => {
            pnp.sp.web.lists.getByTitle(listName).items.getById(listItemId).recycle().then(() => {
                resolve({ code: 200, description: 'Success!', message: 'The item has been moved to recycle bin.' });
            })
                .catch(reason => {
                const error = SpCore.showErrorLog(reason);
                reject(error);
            });
        });
    }
    delete(listItemId, listName) {
        return new Promise((resolve, reject) => {
            pnp.sp.web.lists.getByTitle(listName).items.getById(listItemId).delete().then(() => {
                resolve({ code: 200, description: 'Success!', message: 'The item has been deleted successfully.' });
            })
                .catch(reason => {
                const error = SpCore.showErrorLog(reason);
                reject(error);
            });
        });
    }
    add(listName, data) {
        return new Promise((resolve, reject) => {
            pnp.sp.web.lists.getByTitle(listName).items.add(data)
                .then((iar) => { resolve(iar); })
                .catch(reason => {
                const error = SpCore.showErrorLog(reason);
                reject(error);
            });
        });
    }
    update(listItemId, listName, data) {
        return new Promise((resolve, reject) => {
            pnp.sp.web.lists.getByTitle(listName).items.getById(listItemId).update(data)
                .then((iar) => { resolve(iar); })
                .catch(reason => {
                const error = SpCore.showErrorLog(reason);
                reject(error);
            });
        });
    }
}
export const ListItemsCore = new Core();
class Attachment {
    add(listItemId, listName, attachments) {
        return new Promise((resolve, reject) => {
            pnp.sp.web.lists.getByTitle(listName).items.getById(listItemId).attachmentFiles.addMultiple(attachments)
                .then(() => {
                resolve({ code: 200, description: 'Success!', message: 'The attachments has been added successfully.' });
            })
                .catch(reason => {
                SpCore.showErrorLog(reason);
                reject([]);
            });
        });
    }
    delete(listItemId, listName, attachments) {
        return new Promise((resolve, reject) => {
            pnp.sp.web.lists.getByTitle(listName).items.getById(listItemId).attachmentFiles.deleteMultiple(...attachments)
                .then(() => {
                resolve({ code: 200, description: 'Success!', message: 'The attachments has been deleted successfully.' });
            })
                .catch(reason => {
                SpCore.showErrorLog(reason);
                reject([]);
            });
        });
    }
    retrieve(listItemId, listName) {
        return new Promise((resolve, reject) => {
            pnp.sp.web.lists.getByTitle(listName).items.getById(listItemId).attachmentFiles.get()
                .then((attachments) => {
                resolve(attachments);
            })
                .catch(reason => {
                SpCore.showErrorLog(reason);
                reject([]);
            });
        });
    }
    retrieveForBinding(listItemId, listName) {
        return __awaiter(this, void 0, void 0, function* () {
            const attachmentList = yield this.retrieve(listItemId, listName);
            const attachments = [];
            for (const attachment of attachmentList) {
                const host = attachment.__metadata.uri.split('/s/')[0];
                const serverRelativeUrl = attachment.ServerRelativeUrl;
                const url = encodeURI(host + serverRelativeUrl);
                const file = {
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
        });
    }
    retrieveTxtContent(listItemId, listName, fileName = 'attachment') {
        return new Promise((resolve, reject) => {
            pnp.sp.web.lists.getByTitle(listName).items.getById(listItemId).attachmentFiles.getByName(`${fileName}.txt`).getText()
                .then((image) => {
                resolve(image);
            })
                .catch(reason => {
                SpCore.showErrorLog(reason);
                reject('');
            });
        });
    }
    setTxtContent(listItemId, listName, fileName = 'attachment', content) {
        return new Promise((resolve, reject) => {
            pnp.sp.web.lists.getByTitle(listName).items.getById(listItemId).attachmentFiles.getByName(`${fileName}.txt`).setContent(content)
                .then((attachmentFile) => { resolve(attachmentFile); })
                .catch(reason => {
                const error = SpCore.showErrorLog(reason);
                reject(error);
            });
        });
    }
    addTxtContent(listItemId, listName, fileName = 'attachment', content) {
        return new Promise((resolve, reject) => {
            pnp.sp.web.lists.getByTitle(listName).items.getById(listItemId).attachmentFiles.add(`${fileName}.txt`, content)
                .then((attachmentFile) => { resolve(attachmentFile); })
                .catch(reason => {
                const error = SpCore.showErrorLog(reason);
                reject(error);
            });
        });
    }
}
export const ListItemsAttachment = new Attachment();
class Search {
    retrieveSearch(listName, fieldsArray = [], filter = '', orderBy = 'ID', ascending = true) {
        const fields = fieldsArray.toString().replace('[', '').replace(']', '');
        return new Promise((resolve, reject) => __awaiter(this, void 0, void 0, function* () {
            pnp.sp.web.lists.getByTitle(listName).items
                .orderBy(orderBy, ascending)
                .select(fields)
                .filter(filter)
                .get()
                .then((result) => {
                resolve(result);
            })
                .catch(reason => {
                SpCore.showErrorLog(reason);
                reject(reason);
            });
        }));
    }
}
export const ListItemsSearch = new Search();
//# sourceMappingURL=core-services-list-items.service.js.map