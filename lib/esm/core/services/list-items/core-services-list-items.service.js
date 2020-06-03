var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
var __read = (this && this.__read) || function (o, n) {
    var m = typeof Symbol === "function" && o[Symbol.iterator];
    if (!m) return o;
    var i = m.call(o), r, ar = [], e;
    try {
        while ((n === void 0 || n-- > 0) && !(r = i.next()).done) ar.push(r.value);
    }
    catch (error) { e = { error: error }; }
    finally {
        try {
            if (r && !r.done && (m = i["return"])) m.call(i);
        }
        finally { if (e) throw e.error; }
    }
    return ar;
};
var __spread = (this && this.__spread) || function () {
    for (var ar = [], i = 0; i < arguments.length; i++) ar = ar.concat(__read(arguments[i]));
    return ar;
};
var __values = (this && this.__values) || function(o) {
    var s = typeof Symbol === "function" && Symbol.iterator, m = s && o[s], i = 0;
    if (m) return m.call(o);
    if (o && typeof o.length === "number") return {
        next: function () {
            if (o && i >= o.length) o = void 0;
            return { value: o && o[i++], done: !o };
        }
    };
    throw new TypeError(s ? "Object is not iterable." : "Symbol.iterator is not defined.");
};
import pnp from 'sp-pnp-js';
import { SpCore } from '../setup/core-services-setup.service';
import { AttachmentIcon } from 'gl-w-frontend';
var Core = /** @class */ (function () {
    function Core() {
    }
    /**
     * Retrieve a search list based on fields, filter and ordering
     * @param listName
     * @param fieldsArray (optional)
     */
    Core.prototype.retrieve = function (listName, fieldsArray) {
        var _this = this;
        if (fieldsArray === void 0) { fieldsArray = []; }
        var fields = fieldsArray.toString().replace('[', '').replace(']', '');
        return new Promise(function (resolve, reject) { return __awaiter(_this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                pnp.sp.web.lists.getByTitle(listName).items
                    .select(fields)
                    .get()
                    .then(function (result) {
                    resolve(result);
                })
                    .catch(function (reason) {
                    SpCore.showErrorLog(reason);
                    reject(reason);
                });
                return [2 /*return*/];
            });
        }); });
    };
    /**
     * Move items to recycle bin. The user will be able to restore the information.
     * @param listItemId
     * @param listName
     */
    Core.prototype.recycle = function (listItemId, listName) {
        return new Promise(function (resolve, reject) {
            pnp.sp.web.lists.getByTitle(listName).items.getById(listItemId).recycle().then(function () {
                resolve({ code: 200, description: 'Success!', message: 'The item has been moved to recycle bin.' });
            })
                .catch(function (reason) {
                var error = SpCore.showErrorLog(reason);
                reject(error);
            });
        });
    };
    /**
     * Delete items permanently. The user will not be able to restore the information.
     * @param listItemId
     * @param listName
     */
    Core.prototype.delete = function (listItemId, listName) {
        return new Promise(function (resolve, reject) {
            pnp.sp.web.lists.getByTitle(listName).items.getById(listItemId).delete().then(function () {
                resolve({ code: 200, description: 'Success!', message: 'The item has been deleted successfully.' });
            })
                .catch(function (reason) {
                var error = SpCore.showErrorLog(reason);
                reject(error);
            });
        });
    };
    /**
     * Adds a new item to the collection
     * @param listName
     * @param data
     */
    Core.prototype.add = function (listName, data) {
        return new Promise(function (resolve, reject) {
            pnp.sp.web.lists.getByTitle(listName).items.add(data)
                .then(function (iar) { resolve(iar); })
                .catch(function (reason) {
                var error = SpCore.showErrorLog(reason);
                reject(error);
            });
        });
    };
    /**
     * Update a new item to the collection
     * @param listItemId
     * @param listName
     * @param data
     */
    Core.prototype.update = function (listItemId, listName, data) {
        return new Promise(function (resolve, reject) {
            pnp.sp.web.lists.getByTitle(listName).items.getById(listItemId).update(data)
                .then(function (iar) { resolve(iar); })
                .catch(function (reason) {
                var error = SpCore.showErrorLog(reason);
                reject(error);
            });
        });
    };
    return Core;
}());
export var ListItemsCore = new Core();
var Attachment = /** @class */ (function () {
    function Attachment() {
    }
    /**
     * Before using this method, you need defining the base url on setup (SpCore).
     * Adds multiple new attachment to the collection. Not supported for batching.
     * @param listItemId
     * @param listName
     * @param attachments
     */
    Attachment.prototype.add = function (listItemId, listName, attachments) {
        return new Promise(function (resolve, reject) {
            pnp.sp.web.lists.getByTitle(listName).items.getById(listItemId).attachmentFiles.addMultiple(attachments)
                .then(function () {
                resolve({ code: 200, description: 'Success!', message: 'The attachments has been added successfully.' });
            })
                .catch(function (reason) {
                SpCore.showErrorLog(reason);
                reject([]);
            });
        });
    };
    /**
     * Before using this method, you need defining the base url on setup (SpCore).
     * Delete multiple attachments from the collection. Not supported for batching.
     * @param listItemId
     * @param listName
     * @param attachments
     */
    Attachment.prototype.delete = function (listItemId, listName, attachments) {
        return new Promise(function (resolve, reject) {
            var _a;
            (_a = pnp.sp.web.lists.getByTitle(listName).items.getById(listItemId).attachmentFiles).deleteMultiple.apply(_a, __spread(attachments)).then(function () {
                resolve({ code: 200, description: 'Success!', message: 'The attachments has been deleted successfully.' });
            })
                .catch(function (reason) {
                SpCore.showErrorLog(reason);
                reject([]);
            });
        });
    };
    /**
     * Before using this method, you need defining the base url on setup (SpCore).
     * @param listItemId
     * @param listName
     */
    Attachment.prototype.retrieve = function (listItemId, listName) {
        return new Promise(function (resolve, reject) {
            pnp.sp.web.lists.getByTitle(listName).items.getById(listItemId).attachmentFiles.get()
                .then(function (attachments) {
                resolve(attachments);
            })
                .catch(function (reason) {
                SpCore.showErrorLog(reason);
                reject([]);
            });
        });
    };
    /**
     * Before using this method, you need defining the base url on setup (SpCore).
     * This method is intended to retrieve all attachments inside a list item and prepare the list to use directly to the user interface (binding)
     * @param listItemId
     * @param listName
     */
    Attachment.prototype.retrieveForBinding = function (listItemId, listName) {
        return __awaiter(this, void 0, void 0, function () {
            var attachmentList, attachments, attachmentList_1, attachmentList_1_1, attachment, host, serverRelativeUrl, url, file;
            var e_1, _a;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0: return [4 /*yield*/, this.retrieve(listItemId, listName)];
                    case 1:
                        attachmentList = _b.sent();
                        attachments = [];
                        try {
                            for (attachmentList_1 = __values(attachmentList), attachmentList_1_1 = attachmentList_1.next(); !attachmentList_1_1.done; attachmentList_1_1 = attachmentList_1.next()) {
                                attachment = attachmentList_1_1.value;
                                host = attachment.__metadata.uri.split('/s/')[0];
                                serverRelativeUrl = attachment.ServerRelativeUrl;
                                url = encodeURI(host + serverRelativeUrl);
                                file = {
                                    id: attachmentList.length,
                                    new: false,
                                    remove: false,
                                    name: attachment.FileName,
                                    url: url,
                                    icon: AttachmentIcon.get(attachment.FileName)
                                };
                                attachments.push(file);
                            }
                        }
                        catch (e_1_1) { e_1 = { error: e_1_1 }; }
                        finally {
                            try {
                                if (attachmentList_1_1 && !attachmentList_1_1.done && (_a = attachmentList_1.return)) _a.call(attachmentList_1);
                            }
                            finally { if (e_1) throw e_1.error; }
                        }
                        return [2 /*return*/, attachments];
                }
            });
        });
    };
    /**
     * Before using this method, you need defining the base url on setup (SpCore).
     * @param listItemId
     * @param listName
     * @param fileName Without extension, e.g. "attachment"
     */
    Attachment.prototype.retrieveTxtContent = function (listItemId, listName, fileName) {
        if (fileName === void 0) { fileName = 'attachment'; }
        return new Promise(function (resolve, reject) {
            pnp.sp.web.lists.getByTitle(listName).items.getById(listItemId).attachmentFiles.getByName(fileName + ".txt").getText()
                .then(function (image) {
                resolve(image);
            })
                .catch(function (reason) {
                SpCore.showErrorLog(reason);
                reject('');
            });
        });
    };
    /**
     * Sets the content of a file. Not supported for batching
     * @param listItemId
     * @param listName
     * @param fileName Without extension, e.g. "attachment"
     * @param content Content to be added to the text
     */
    Attachment.prototype.setTxtContent = function (listItemId, listName, fileName, content) {
        if (fileName === void 0) { fileName = 'attachment'; }
        return new Promise(function (resolve, reject) {
            pnp.sp.web.lists.getByTitle(listName).items.getById(listItemId).attachmentFiles.getByName(fileName + ".txt").setContent(content)
                .then(function (attachmentFile) { resolve(attachmentFile); })
                .catch(function (reason) {
                var error = SpCore.showErrorLog(reason);
                reject(error);
            });
        });
    };
    /**
     * Adds a new attachment to the collection. Not supported for batching.
     * @param listItemId
     * @param listName
     * @param fileName Without extension, e.g. "attachment"
     * @param content Content to be added to the text
     */
    Attachment.prototype.addTxtContent = function (listItemId, listName, fileName, content) {
        if (fileName === void 0) { fileName = 'attachment'; }
        return new Promise(function (resolve, reject) {
            pnp.sp.web.lists.getByTitle(listName).items.getById(listItemId).attachmentFiles.add(fileName + ".txt", content)
                .then(function (attachmentFile) { resolve(attachmentFile); })
                .catch(function (reason) {
                var error = SpCore.showErrorLog(reason);
                reject(error);
            });
        });
    };
    return Attachment;
}());
export var ListItemsAttachment = new Attachment();
var Search = /** @class */ (function () {
    function Search() {
    }
    /**
     * Retrieve a search list based on fields, filter and ordering
     * @param listName
     * @param fieldsArray (optional)
     * @param filter (optional)
     * @param orderBy (optional)
     * @param ascending (optional)
     */
    Search.prototype.retrieveSearch = function (listName, fieldsArray, filter, orderBy, ascending) {
        var _this = this;
        if (fieldsArray === void 0) { fieldsArray = []; }
        if (filter === void 0) { filter = ''; }
        if (orderBy === void 0) { orderBy = 'ID'; }
        if (ascending === void 0) { ascending = true; }
        var fields = fieldsArray.toString().replace('[', '').replace(']', '');
        return new Promise(function (resolve, reject) { return __awaiter(_this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                pnp.sp.web.lists.getByTitle(listName).items
                    .orderBy(orderBy, ascending)
                    .select(fields)
                    .filter(filter)
                    .get()
                    .then(function (result) {
                    resolve(result);
                })
                    .catch(function (reason) {
                    SpCore.showErrorLog(reason);
                    reject(reason);
                });
                return [2 /*return*/];
            });
        }); });
    };
    return Search;
}());
export var ListItemsSearch = new Search();
//# sourceMappingURL=core-services-list-items.service.js.map