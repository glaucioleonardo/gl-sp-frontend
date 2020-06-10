var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
import { default as pnp } from 'sp-pnp-js';
import { SpCore } from '../setup/core-services-setup.service';
class Core {
    currentUser() {
        return new Promise((resolve, reject) => {
            pnp.sp.web.currentUser.get()
                .then((user) => {
                resolve(user);
            })
                .catch(reason => {
                reject(SpCore.onError(reason));
            });
        });
    }
}
export const SpUserCore = new Core();
class Permissions {
    isAdmin() {
        return __awaiter(this, void 0, void 0, function* () {
            const currentUser = yield SpUserCore.currentUser();
            return currentUser.IsSiteAdmin;
        });
    }
    isInGroup(groupName, userEmail) {
        return new Promise((resolve, reject) => {
            pnp.sp.web.siteGroups.getByName(groupName).users.get()
                .then((users) => {
                const user = users.filter(x => x.Email === userEmail);
                resolve(user.length > 0);
            })
                .catch(reason => {
                reject(reason);
            });
        });
    }
    isCurrentUserInGroup(groupName) {
        return __awaiter(this, void 0, void 0, function* () {
            const currentUser = yield SpUserCore.currentUser();
            return yield this.isInGroup(groupName, currentUser.Email);
        });
    }
}
export const SpUserPermissions = new Permissions();
//# sourceMappingURL=core-services-user.service.js.map