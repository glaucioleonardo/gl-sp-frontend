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
}
export const SpUserCore = new Core();
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