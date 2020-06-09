import { default as pnp } from 'sp-pnp-js';
import { SpCore } from '../setup/core-services-setup.service';
import { ISpCoreResult, ISpCurrentUser } from '../setup/core-services-setup.interface';

class Core {
  currentUser(): PromiseLike<ISpCurrentUser> {
    return new Promise((resolve, reject) => {
      pnp.sp.web.currentUser.get()
        .then((user: ISpCurrentUser) => {
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
  async isAdmin(): Promise<boolean> {
    const currentUser: ISpCurrentUser = await SpUserCore.currentUser();
    return currentUser.IsSiteAdmin;
  }

  isInGroup(groupName: string, userEmail: string): Promise<boolean | ISpCoreResult> {
    return new Promise((resolve, reject) => {
      pnp.sp.web.siteGroups.getByName(groupName).users.get()
        .then((users: ISpCurrentUser[]) => {
          const user: ISpCurrentUser[] = users.filter(x => x.Email === userEmail);
          resolve(user.length > 0);
        })
        .catch(reason => {
          reject(reason);
        });
    });
  }

  async isCurrentUserInGroup(groupName: string): Promise<any> {
    const currentUser: ISpCurrentUser = await SpUserCore.currentUser();
    return await this.isInGroup(groupName, currentUser.Email);
  }
}

export const SpUserPermissions = new Permissions();
