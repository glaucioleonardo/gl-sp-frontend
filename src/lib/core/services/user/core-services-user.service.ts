import { default as pnp } from 'sp-pnp-js';
import { SpCore } from '../setup/core-services-setup.service';
import { ISpCoreResult, ISpCurrentUser } from '../setup/core-services-setup.interface';

class Core {
  async currentUser(baseUrl: string): Promise<ISpCurrentUser> {
    const base = baseUrl == null ? SpCore.baseUrl : baseUrl;

    try {
      return await pnp.sp.configure(SpCore.config, base).web.currentUser.get() as ISpCurrentUser;
    } catch (reason) {
      const error = SpCore.onError(reason);
      throw new Error (`Error code: ${error.code}.\nError message: ${error.message}.\nError description: ${error.description}`);
    }
  }
}

export const SpUserCore = new Core();

class Permissions {
  async isAdmin(baseUrl: string): Promise<boolean> {
    const base = baseUrl == null ? SpCore.baseUrl : baseUrl;
    const currentUser: ISpCurrentUser = await SpUserCore.currentUser(base);
    return currentUser.IsSiteAdmin;
  }

  async isInGroup(groupName: string, userEmail: string, baseUrl: string): Promise<boolean | ISpCoreResult> {
    const base = baseUrl == null ? SpCore.baseUrl : baseUrl;

    try {
      const user: ISpCurrentUser = await pnp.sp.configure(SpCore.config, base).web.siteGroups.getByName(groupName).users.getByEmail(userEmail).get();
      return user.Email != null && user.Email.length > 0;
    } catch (reason) {
      const error = SpCore.onError(reason);
      throw new Error (`Error code: ${error.code}.\nError message: ${error.message}.\nError description: ${error.description}`);
    }
  }

  async isCurrentUserInGroup(groupName: string, baseUrl: string): Promise<any> {
    const base = baseUrl == null ? SpCore.baseUrl : baseUrl;

    const currentUser: ISpCurrentUser = await SpUserCore.currentUser(base);
    return await this.isInGroup(groupName, currentUser.Email, base);
  }
}

export const SpUserPermissions = new Permissions();
