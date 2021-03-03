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
import '@pnp/sp/fields';
import '@pnp/sp/lists';
import '@pnp/sp/webs';
import '@pnp/sp/views';
import { CalendarType, DateTimeFieldFormatType, DateTimeFieldFriendlyFormatType, FieldTypes } from '@pnp/sp/fields';
class Core {
    fieldsToStringArray(fields) {
        return fields.toString().replace('[', '').replace(']', '');
    }
    exists(listName, baseUrl) {
        return __awaiter(this, void 0, void 0, function* () {
            const base = baseUrl == null ? SpCore.baseUrl : baseUrl;
            try {
                const lists = yield this.retrieve(listName, base);
                return lists.filter(x => x.Title === listName).length > 0;
            }
            catch (reason) {
                const error = SpCore.onError(reason);
                SpCore.showErrorLog(reason);
                throw new Error(error.code.toString());
            }
        });
    }
    retrieveSingle(listName, baseUrl) {
        return __awaiter(this, void 0, void 0, function* () {
            const base = baseUrl == null ? SpCore.baseUrl : baseUrl;
            try {
                return yield sp.configure(SpCore.config, base).web.lists.getByTitle(listName).get();
            }
            catch (reason) {
                const error = SpCore.onError(reason);
                SpCore.showErrorLog(reason);
                throw new Error(error.code.toString());
            }
        });
    }
    retrieve(listName, baseUrl) {
        return __awaiter(this, void 0, void 0, function* () {
            const base = baseUrl == null ? SpCore.baseUrl : baseUrl;
            try {
                return yield sp.configure(SpCore.config, base).web.lists.get();
            }
            catch (reason) {
                const error = SpCore.onError(reason);
                SpCore.showErrorLog(reason);
                throw new Error(error.code.toString());
            }
        });
    }
    recycle(listName, baseUrl) {
        return __awaiter(this, void 0, void 0, function* () {
            const base = baseUrl == null ? SpCore.baseUrl : baseUrl;
            try {
                const exists = yield this.exists(listName, base);
                if (exists) {
                    yield sp.configure(SpCore.config, base).web.lists.getByTitle(listName).recycle();
                    return {
                        code: 200,
                        description: 'Success!',
                        message: 'The current list has been recycled.'
                    };
                }
                else {
                    return {
                        code: 503,
                        description: 'Internal Error!',
                        message: 'The current list doesn\'t exist.'
                    };
                }
            }
            catch (reason) {
                SpCore.showErrorLog(reason);
                return {
                    code: 500,
                    description: 'Internal Error!',
                    message: reason.message
                };
            }
        });
    }
    recreate(listName, baseUrl, fields = [], titleRequired = true, properties) {
        return __awaiter(this, void 0, void 0, function* () {
            const base = baseUrl == null ? SpCore.baseUrl : baseUrl;
            const listProperties = properties != null ? properties : {
                AllowContentTypes: true,
                BaseTemplate: 100,
                BaseType: 0,
                ContentTypesEnabled: false,
                EnableAttachments: true,
                DocumentTemplateUrl: undefined,
                EnableVersioning: true,
                Description: ''
            };
            try {
                yield this.recycle(listName, baseUrl);
                yield sp.configure(SpCore.config, base).web.lists.add(listName, listProperties.Description, listProperties.BaseTemplate, false, listProperties);
                yield sp.configure(SpCore.config, base).web.lists.getByTitle(listName).fields.getByTitle('Title').update({
                    Required: titleRequired,
                    __metadata: { type: 'SP.FieldText' }
                });
                yield this.addFields(listName, fields, base);
                return {
                    code: 200,
                    description: 'Success!',
                    message: 'The current list has been recycled.'
                };
            }
            catch (reason) {
                SpCore.showErrorLog(reason);
                return {
                    code: 500,
                    description: 'Internal Error!',
                    message: reason.message
                };
            }
        });
    }
    addFields(listName, fields, baseUrl) {
        return __awaiter(this, void 0, void 0, function* () {
            const base = baseUrl == null ? SpCore.baseUrl : baseUrl;
            try {
                for (const [i, field] of fields.entries()) {
                    const title = field.Title == null ? `Column${i + 1}` : field.Title;
                    const required = field.Required == null ? false : field.Required;
                    const indexed = field.Indexed == null ? false : field.Indexed;
                    const defaultBooleanValue = field.DefaultValue == null ? '0' : field.DefaultValue;
                    switch (field.FieldTypeKind) {
                        case FieldTypes.Text:
                            yield sp.configure(SpCore.config, base).web.lists.getByTitle(listName).fields.addText(title, undefined, {
                                Required: required,
                                Indexed: indexed
                            });
                            break;
                        case FieldTypes.Note:
                            yield sp.configure(SpCore.config, base).web.lists.getByTitle(listName).fields.addMultilineText(title, 6, false, false, false, false, {
                                Required: required,
                                Indexed: indexed
                            });
                            break;
                        case FieldTypes.Boolean:
                            yield sp.configure(SpCore.config, base).web.lists.getByTitle(listName).fields.addBoolean(title, {
                                Required: required,
                                Indexed: indexed,
                                DefaultValue: defaultBooleanValue
                            });
                            break;
                        case FieldTypes.Number:
                            yield sp.configure(SpCore.config, base).web.lists.getByTitle(listName).fields.addNumber(title, undefined, undefined, {
                                Required: required,
                                Indexed: indexed,
                                DefaultValue: defaultBooleanValue
                            });
                            break;
                        case FieldTypes.DateTime:
                            yield sp.configure(SpCore.config, base).web.lists.getByTitle(listName).fields.addDateTime(title, DateTimeFieldFormatType.DateOnly, CalendarType.Gregorian, DateTimeFieldFriendlyFormatType.Disabled, {
                                Required: required,
                                Indexed: indexed,
                                DefaultValue: defaultBooleanValue
                            });
                            break;
                    }
                    yield sp.configure(SpCore.config, base).web.lists.getByTitle(listName).defaultView.fields.add(title);
                }
                return {
                    code: 200,
                    description: 'Success!',
                    message: 'The current list has been recycled.'
                };
            }
            catch (reason) {
                SpCore.showErrorLog(reason);
                return {
                    code: 500,
                    description: 'Internal Error!',
                    message: reason.message
                };
            }
        });
    }
}
export const ListsCore = new Core();
//# sourceMappingURL=core-services-lists.service.js.map