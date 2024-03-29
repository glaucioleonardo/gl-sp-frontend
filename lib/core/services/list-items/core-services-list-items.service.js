var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
import { SpCore } from '../setup/core-services-setup.service';
import { sp } from '@pnp/sp/presets/core';
import "@pnp/sp/attachments";
import { ArrayRemove } from 'gl-w-array-frontend';
import { AttachmentIcon } from 'gl-w-attachment-frontend';
class Core {
    fieldsToStringArray(fields) {
        return fields.toString().replace('[', '').replace(']', '');
    }
    retrieve(listName, fieldsArray = [], baseUrl) {
        return __awaiter(this, void 0, void 0, function* () {
            const fields = this.fieldsToStringArray(fieldsArray);
            const base = baseUrl == null ? SpCore.baseUrl : baseUrl;
            try {
                return yield sp.configure(SpCore.config, base).web.lists.getByTitle(listName).items.select(fields).getAll();
            }
            catch (reason) {
                const error = SpCore.onError(reason);
                SpCore.showErrorLog(reason);
                throw new Error(error.code.toString());
            }
        });
    }
    retrieveExternal(listName, fieldsArray = [], baseUrl, top = 0) {
        return __awaiter(this, void 0, void 0, function* () {
            const fields = this.fieldsToStringArray(fieldsArray);
            try {
                const listUrl = encodeURI(`${baseUrl}/_api/web/lists/GetByTitle('${listName}')?$select=ItemCount`);
                let count = top;
                if (top == null || top === 0) {
                    const result = yield fetch(listUrl, yield SpCore.fetchHeader());
                    const list = yield result.json();
                    count = yield list.d.ItemCount;
                }
                const itemsUrl = encodeURI(`${baseUrl}/_api/web/lists/GetByTitle('${listName}')/items?$select=${fields}&$top=${count}`);
                const resultItems = yield fetch(itemsUrl, yield SpCore.fetchHeader());
                const fetchItems = yield resultItems.json();
                return fetchItems.d.results;
            }
            catch (reason) {
                SpCore.showErrorLog(reason);
                return [];
            }
        });
    }
    retrieveExternalSingle(id, listName, fieldsArray = [], baseUrl) {
        return __awaiter(this, void 0, void 0, function* () {
            const fields = this.fieldsToStringArray(fieldsArray);
            try {
                const itemsUrl = encodeURI(`${baseUrl}/_api/web/lists/GetByTitle('${listName}')/items(${id})?$select=${fields}`);
                const resultItems = yield fetch(itemsUrl, yield SpCore.fetchHeader());
                const fetchItems = yield resultItems.json();
                return fetchItems.d;
            }
            catch (reason) {
                SpCore.showErrorLog(reason);
                return [];
            }
        });
    }
    retrieveForCombobox(listName, baseUrl, valueField = 'value', textField = 'text') {
        return __awaiter(this, void 0, void 0, function* () {
            try {
                const items = [];
                if (listName != null && listName.length > 0) {
                    const fields = [valueField, textField];
                    try {
                        const itemsResult = yield this.retrieve(listName, fields, baseUrl);
                        for (const item of itemsResult) {
                            items.push({
                                text: item[textField],
                                value: item[valueField].toString()
                            });
                        }
                        return items;
                    }
                    catch (reason) {
                        SpCore.showErrorLog(reason);
                        return reason;
                    }
                }
                else {
                    return [];
                }
            }
            catch (reason) {
                SpCore.showErrorLog(reason);
                return reason;
            }
        });
    }
    retrieveExternalForCombobox(listName, baseUrl, valueField = 'value', textField = 'text', top = 0) {
        return __awaiter(this, void 0, void 0, function* () {
            try {
                const fields = [valueField, textField];
                const listItems = yield this.retrieveExternal(listName, fields, baseUrl, top);
                const comboBox = [];
                for (const item of listItems) {
                    comboBox.push({
                        text: item[textField],
                        value: item[valueField].toString()
                    });
                }
                return comboBox;
            }
            catch (reason) {
                SpCore.showErrorLog(reason);
                return reason;
            }
        });
    }
    retrieveSingle(listItemId, listName, fieldsArray = [], baseUrl) {
        return __awaiter(this, void 0, void 0, function* () {
            const fields = this.fieldsToStringArray(fieldsArray);
            const base = baseUrl == null ? SpCore.baseUrl : baseUrl;
            try {
                return yield sp.configure(SpCore.config, base).web.lists.getByTitle(listName).items.getById(listItemId).select(fields).get();
            }
            catch (reason) {
                const error = SpCore.onError(reason);
                SpCore.showErrorLog(reason);
                throw new Error(error.code.toString());
            }
        });
    }
    recycle(listItemId, listName, baseUrl) {
        const base = baseUrl == null ? SpCore.baseUrl : baseUrl;
        return new Promise((resolve, reject) => {
            sp.configure(SpCore.config, base).web.lists.getByTitle(listName).items.getById(listItemId).recycle().then(() => {
                resolve({ code: 200, description: 'Success!', message: 'The item has been moved to recycle bin.' });
            })
                .catch(reason => {
                const error = SpCore.showErrorLog(reason);
                reject(error);
            });
        });
    }
    recycleDuplicated(listName, field, baseUrl) {
        return __awaiter(this, void 0, void 0, function* () {
            const base = baseUrl == null ? SpCore.baseUrl : baseUrl;
            try {
                const data = yield this.retrieve(listName, ['Id', field], base);
                const duplicates = yield ArrayRemove.notDuplicatedByKey(data, field, field, true);
                for (const duplicate of duplicates) {
                    const items = data.filter(x => x[field] === duplicate);
                    for (let i = 1; i < items.length; i++) {
                        yield this.recycle(items[i].Id, listName, base);
                    }
                }
                return duplicates;
            }
            catch (reason) {
                SpCore.showErrorLog(reason);
                return [];
            }
        });
    }
    delete(listItemId, listName, baseUrl) {
        const base = baseUrl == null ? SpCore.baseUrl : baseUrl;
        return new Promise((resolve, reject) => {
            sp.configure(SpCore.config, base).web.lists.getByTitle(listName).items.getById(listItemId).delete().then(() => {
                resolve({ code: 200, description: 'Success!', message: 'The item has been deleted successfully.' });
            })
                .catch(reason => {
                const error = SpCore.showErrorLog(reason);
                reject(error);
            });
        });
    }
    deleteDuplicated(listName, field, baseUrl) {
        return __awaiter(this, void 0, void 0, function* () {
            const base = baseUrl == null ? SpCore.baseUrl : baseUrl;
            try {
                const data = yield this.retrieve(listName, ['Id', field], base);
                const duplicates = yield ArrayRemove.notDuplicatedByKey(data, field, field, true);
                for (const duplicate of duplicates) {
                    const items = data.filter(x => x[field] === duplicate);
                    for (let i = 1; i < items.length; i++) {
                        yield this.delete(items[i].Id, listName, base);
                    }
                }
                return duplicates;
            }
            catch (reason) {
                SpCore.showErrorLog(reason);
                return [];
            }
        });
    }
    add(listName, data, baseUrl) {
        return __awaiter(this, void 0, void 0, function* () {
            const base = baseUrl == null ? SpCore.baseUrl : baseUrl;
            try {
                return yield sp.configure(SpCore.config, base).web.lists.getByTitle(listName).items.add(data);
            }
            catch (reason) {
                SpCore.showErrorLog(reason);
                throw new Error(reason);
            }
        });
    }
    update(listItemId, listName, data, baseUrl) {
        const base = baseUrl == null ? SpCore.baseUrl : baseUrl;
        return new Promise((resolve, reject) => {
            sp.configure(SpCore.config, base).web.lists.getByTitle(listName).items.getById(listItemId).update(data)
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
    add(listItemId, listName, attachments, baseUrl) {
        const base = baseUrl == null ? SpCore.baseUrl : baseUrl;
        return new Promise((resolve, reject) => {
            sp.configure(SpCore.config, base).web.lists.getByTitle(listName).items.getById(listItemId).attachmentFiles.addMultiple(attachments)
                .then(() => {
                resolve({ code: 200, description: 'Success!', message: 'The attachments has been added successfully.' });
            })
                .catch(reason => {
                SpCore.showErrorLog(reason);
                reject([]);
            });
        });
    }
    delete(listItemId, listName, attachments, baseUrl) {
        const base = baseUrl == null ? SpCore.baseUrl : baseUrl;
        return new Promise((resolve, reject) => {
            sp.configure(SpCore.config, base).web.lists.getByTitle(listName).items.getById(listItemId).attachmentFiles.deleteMultiple(...attachments)
                .then(() => {
                resolve({ code: 200, description: 'Success!', message: 'The attachments has been deleted successfully.' });
            })
                .catch(reason => {
                SpCore.showErrorLog(reason);
                reject([]);
            });
        });
    }
    retrieve(listItemId, listName, baseUrl) {
        const base = baseUrl == null ? SpCore.baseUrl : baseUrl;
        return new Promise((resolve, reject) => {
            sp.configure(SpCore.config, base).web.lists.getByTitle(listName).items.getById(listItemId).attachmentFiles.get()
                .then((attachments) => {
                resolve(attachments);
            })
                .catch(reason => {
                SpCore.showErrorLog(reason);
                reject([]);
            });
        });
    }
    retrieveBlob(listItemId, listName, fileName, baseUrl) {
        return __awaiter(this, void 0, void 0, function* () {
            const base = baseUrl == null ? SpCore.baseUrl : baseUrl;
            try {
                const item = sp.configure(SpCore.config, base).web.lists.getByTitle(listName).items.getById(listItemId);
                return yield item.attachmentFiles.getByName(fileName).getBlob();
            }
            catch (reason) {
                SpCore.showErrorLog(reason);
                return null;
            }
        });
    }
    retrieveMultipleBlobSameItem(listItemId, listName, fileNames, baseUrl) {
        return __awaiter(this, void 0, void 0, function* () {
            const base = baseUrl == null ? SpCore.baseUrl : baseUrl;
            const blobs = [];
            try {
                for (const fileName of fileNames) {
                    const item = sp.configure(SpCore.config, base).web.lists.getByTitle(listName).items.getById(listItemId);
                    blobs.push({
                        fileName,
                        file: yield item.attachmentFiles.getByName(fileName).getBlob(),
                        icon: AttachmentIcon.get(fileName)
                    });
                }
                return blobs;
            }
            catch (reason) {
                SpCore.showErrorLog(reason);
                return [];
            }
        });
    }
    retrieveMultipleBlob(data, baseUrl) {
        return __awaiter(this, void 0, void 0, function* () {
            const base = baseUrl == null ? SpCore.baseUrl : baseUrl;
            const blobs = [];
            try {
                for (const content of data) {
                    const item = sp.configure(SpCore.config, base).web.lists.getByTitle(content.listName).items.getById(content.id);
                    blobs.push({
                        fileName: content.fileName,
                        file: yield item.attachmentFiles.getByName(content.fileName).getBlob(),
                        icon: AttachmentIcon.get(content.fileName)
                    });
                }
                return blobs;
            }
            catch (reason) {
                SpCore.showErrorLog(reason);
                return [];
            }
        });
    }
    retrieveForBinding(listItemId, listName, baseUrl) {
        return __awaiter(this, void 0, void 0, function* () {
            let base = baseUrl == null ? SpCore.baseUrl : baseUrl;
            if (!base.includes('http')) {
                base = 'http://' + base;
            }
            const attachmentList = yield this.retrieve(listItemId, listName, base);
            const attachments = [];
            for (const attachment of attachmentList) {
                const host = base.split('://')[0] + '://' + base.split('://')[1].split('/')[0];
                const serverRelativeUrl = attachment.ServerRelativeUrl;
                const url = encodeURI(host + serverRelativeUrl);
                const file = {
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
        });
    }
    retrieveTxtContent(listItemId, listName, fileName = 'attachment', baseUrl) {
        const base = baseUrl == null ? SpCore.baseUrl : baseUrl;
        return new Promise((resolve, reject) => {
            sp.configure(SpCore.config, base).web.lists.getByTitle(listName).items.getById(listItemId).attachmentFiles.getByName(`${fileName}.txt`).getText()
                .then((image) => {
                resolve(image);
            })
                .catch(reason => {
                SpCore.showErrorLog(reason);
                reject('');
            });
        });
    }
    setTxtContent(listItemId, listName, fileName = 'attachment', content, baseUrl) {
        const base = baseUrl == null ? SpCore.baseUrl : baseUrl;
        return new Promise((resolve, reject) => {
            sp.configure(SpCore.config, base).web.lists.getByTitle(listName).items.getById(listItemId).attachmentFiles.getByName(`${fileName}.txt`).setContent(content)
                .then((attachmentFile) => { resolve(attachmentFile); })
                .catch(reason => {
                const error = SpCore.showErrorLog(reason);
                reject(error);
            });
        });
    }
    addTxtContent(listItemId, listName, fileName = 'attachment', content, baseUrl) {
        const base = baseUrl == null ? SpCore.baseUrl : baseUrl;
        return new Promise((resolve, reject) => {
            sp.configure(SpCore.config, base).web.lists.getByTitle(listName).items.getById(listItemId).attachmentFiles.add(`${fileName}.txt`, content)
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
    retrieveSearch(listName, fieldsArray = [], filter = '', orderBy = 'ID', ascending = true, baseUrl) {
        const base = baseUrl == null ? SpCore.baseUrl : baseUrl;
        const fields = fieldsArray.toString().replace('[', '').replace(']', '');
        return new Promise((resolve, reject) => __awaiter(this, void 0, void 0, function* () {
            sp.configure(SpCore.config, base).web.lists.getByTitle(listName).items
                .orderBy(orderBy, ascending)
                .select(fields)
                .filter(filter)
                .getAll()
                .then((result) => {
                resolve(result);
            })
                .catch(reason => {
                SpCore.showErrorLog(reason);
                reject(reason);
            });
        }));
    }
    retrieveSearchLimited(listName, fieldsArray = [], maxItems = 100, filter = '', orderBy = 'ID', ascending = true, baseUrl) {
        const base = baseUrl == null ? SpCore.baseUrl : baseUrl;
        const fields = fieldsArray.toString().replace('[', '').replace(']', '');
        return new Promise((resolve, reject) => __awaiter(this, void 0, void 0, function* () {
            sp.configure(SpCore.config, base).web.lists.getByTitle(listName).items
                .orderBy(orderBy, ascending)
                .top(maxItems)
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