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
import { sp } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-groups/web";
import '@pnp/sp/sputilities';
import { ArraySort } from 'gl-w-frontend';
class Core {
    currentUser(baseUrl) {
        return __awaiter(this, void 0, void 0, function* () {
            const base = baseUrl == null ? SpCore.baseUrl : baseUrl;
            try {
                return yield sp.configure(SpCore.config, base).web.currentUser.get();
            }
            catch (reason) {
                const error = SpCore.onError(reason);
                throw new Error(`Error code: ${error.code}.\nError message: ${error.message}.\nError description: ${error.description}`);
            }
        });
    }
    userData(baseUrl, id) {
        return __awaiter(this, void 0, void 0, function* () {
            const base = baseUrl == null ? SpCore.baseUrl : baseUrl;
            try {
                return yield sp.configure(SpCore.config, base).web.siteUsers.getById(id).get();
            }
            catch (reason) {
                const error = SpCore.onError(reason);
                throw new Error(`Error code: ${error.code}.\nError message: ${error.message}.\nError description: ${error.description}`);
            }
        });
    }
    usersList(base, hasEmail = true) {
        return __awaiter(this, void 0, void 0, function* () {
            let users = yield sp.configure(SpCore.config, base).web.siteUsers.get();
            if (hasEmail) {
                users = users.filter(x => x.Email.length > 0);
                users = yield ArraySort.byKey(users, 'Title', true);
            }
            return users;
        });
    }
    usersListCombobox(base, valueField, textField) {
        return __awaiter(this, void 0, void 0, function* () {
            const users = yield this.usersList(base, true);
            const list = [];
            for (const user of users) {
                list.push({
                    value: user[valueField].toString(),
                    text: user[textField].toString()
                });
            }
            return list;
        });
    }
}
export const SpUserCore = new Core();
class Email {
    send(baseUrl, subject, to, body) {
        return __awaiter(this, void 0, void 0, function* () {
            const base = baseUrl == null ? SpCore.baseUrl : baseUrl;
            const emailProps = {
                AdditionalHeaders: {
                    'content-type': 'text/html'
                },
                Subject: subject,
                To: to,
                Body: body
            };
            try {
                yield sp.configure(SpCore.config, base).utility.sendEmail(emailProps);
                return true;
            }
            catch (reason) {
                const error = SpCore.onError(reason);
                throw new Error(`Error code: ${error.code}.\nError message: ${error.message}.\nError description: ${error.description}`);
            }
        });
    }
}
export const SpUserEmail = new Email();
class Permissions {
    isAdmin(baseUrl) {
        return __awaiter(this, void 0, void 0, function* () {
            const base = baseUrl == null ? SpCore.baseUrl : baseUrl;
            const currentUser = yield SpUserCore.currentUser(base);
            return currentUser.IsSiteAdmin;
        });
    }
    isInGroup(groupName, userEmail, baseUrl) {
        return __awaiter(this, void 0, void 0, function* () {
            const base = baseUrl == null ? SpCore.baseUrl : baseUrl;
            try {
                const user = yield sp.configure(SpCore.config, base).web.siteGroups.getByName(groupName).users.getByEmail(userEmail).get();
                return user.Email != null && user.Email.length > 0;
            }
            catch (reason) {
                const error = SpCore.onError(reason);
                throw new Error(`Error code: ${error.code}.\nError message: ${error.message}.\nError description: ${error.description}`);
            }
        });
    }
    isCurrentUserInGroup(groupName, baseUrl) {
        return __awaiter(this, void 0, void 0, function* () {
            const base = baseUrl == null ? SpCore.baseUrl : baseUrl;
            const currentUser = yield SpUserCore.currentUser(base);
            return yield this.isInGroup(groupName, currentUser.Email, base);
        });
    }
}
export const SpUserPermissions = new Permissions();
//# sourceMappingURL=core-services-user.service.js.map