import { ISpCoreResult, ISpCurrentUser } from '../setup/core-services-setup.interface';
declare class Core {
    currentUser(baseUrl: string): Promise<ISpCurrentUser>;
}
export declare const SpUserCore: Core;
declare class Permissions {
    isAdmin(baseUrl: string): Promise<boolean>;
    isInGroup(groupName: string, userEmail: string, baseUrl: string): Promise<boolean | ISpCoreResult>;
    isCurrentUserInGroup(groupName: string, baseUrl: string): Promise<any>;
}
export declare const SpUserPermissions: Permissions;
export {};
