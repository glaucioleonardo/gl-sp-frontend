import { ISpCoreResult, ISpCurrentUser } from '../setup/core-services-setup.interface';
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-groups/web";
import '@pnp/sp/sputilities';
import { ISiteUserInfo } from '@pnp/sp/site-users/types';
declare class Core {
    currentUser(baseUrl: string): Promise<ISpCurrentUser>;
    userData(baseUrl: string, id: number): Promise<ISiteUserInfo>;
}
export declare const SpUserCore: Core;
declare class Email {
    send(baseUrl: string, subject: string, to: string[], body: string): Promise<boolean>;
}
export declare const SpUserEmail: Email;
declare class Permissions {
    isAdmin(baseUrl: string): Promise<boolean>;
    isInGroup(groupName: string, userEmail: string, baseUrl: string): Promise<boolean | ISpCoreResult>;
    isCurrentUserInGroup(groupName: string, baseUrl: string): Promise<any>;
}
export declare const SpUserPermissions: Permissions;
export {};
