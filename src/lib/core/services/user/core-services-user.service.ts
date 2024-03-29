import { SpCore } from '../setup/core-services-setup.service';
import { ISpCoreResult, ISpCurrentUser } from '../setup/core-services-setup.interface';
import { sp } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-groups/web";
import '@pnp/sp/sputilities';
import { ISiteUserInfo } from '@pnp/sp/site-users/types';
import { IEmailProperties } from '@pnp/sp/sputilities';
import { IUserList, TUserListField } from './core-services-user.interface';
import { ArraySort } from 'gl-w-array-frontend';
import { IComboBoxData } from 'gl-w-combobox-frontend';

class Core {
  async currentUser(baseUrl: string): Promise<ISpCurrentUser> {
    const base = baseUrl == null ? SpCore.baseUrl : baseUrl;

    try {
      return await sp.configure(SpCore.config, base).web.currentUser.get() as ISpCurrentUser;
    } catch (reason) {
      const error = SpCore.onError(reason);
      throw new Error (`Error code: ${error.code}.\nError message: ${error.message}.\nError description: ${error.description}`);
    }
  }

  async userData(baseUrl: string, id: number): Promise<ISiteUserInfo> {
    const base = baseUrl == null ? SpCore.baseUrl : baseUrl;

    try {
      return await sp.configure(SpCore.config, base).web.siteUsers.getById(id).get();
    } catch (reason) {
      const error = SpCore.onError(reason);
      throw new Error (`Error code: ${error.code}.\nError message: ${error.message}.\nError description: ${error.description}`);
    }
  }

  async usersList(base: string, hasEmail: boolean = true): Promise<IUserList[]> {
    let users = await sp.configure(SpCore.config, base).web.siteUsers.get() as IUserList[];
    if (hasEmail) {
      users = users.filter(x => x.Email.length > 0);
      users = await ArraySort.byKey(users, 'Title', true);
    }

    return users;
  }
  async usersListCombobox(base: string, valueField: TUserListField, textField: TUserListField): Promise<IComboBoxData[]> {
    const users = await this.usersList(base, true);
    const list: IComboBoxData[] = [];

    for (const user of users) {
      list.push({
        value: user[valueField].toString(),
        text: user[textField].toString()
      });
    }

    return list;
  }
}
export const SpUserCore = new Core();

class Email {
  async send(baseUrl: string, subject: string, to: string[], body: string): Promise<boolean> {
    const base = baseUrl == null ? SpCore.baseUrl : baseUrl;

    const emailProps: IEmailProperties = {
      AdditionalHeaders: {
        'content-type': 'text/html'
      },
      Subject: subject,
      To: to,
      Body: body
    };

    try {
      await sp.configure(SpCore.config, base).utility.sendEmail(emailProps);
      return true;
    } catch (reason) {
      const error = SpCore.onError(reason);
      throw new Error (`Error code: ${error.code}.\nError message: ${error.message}.\nError description: ${error.description}`);
    }
  }
}
export const SpUserEmail = new Email();

class Permissions {
  async isAdmin(baseUrl: string): Promise<boolean> {
    const base = baseUrl == null ? SpCore.baseUrl : baseUrl;
    const currentUser: ISpCurrentUser = await SpUserCore.currentUser(base);
    return currentUser.IsSiteAdmin;
  }

  async isInGroup(groupName: string, userEmail: string, baseUrl: string): Promise<boolean | ISpCoreResult> {
    const base = baseUrl == null ? SpCore.baseUrl : baseUrl;

    try {
      const user: ISpCurrentUser[] = await sp.configure(SpCore.config, base).web.siteGroups.getByName(groupName).users.get();
      return user.filter(x => x.Email === userEmail).length > 0;
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
