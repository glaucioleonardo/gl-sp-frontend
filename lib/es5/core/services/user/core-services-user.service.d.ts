import { ISpCoreResult, ISpCurrentUser } from '../setup/core-services-setup.interface';
declare class Core {
    currentUser(): PromiseLike<ISpCurrentUser>;
}
export declare const SpUserCore: Core;
declare class Permissions {
    isAdmin(): Promise<boolean>;
    isInGroup(groupName: string, userEmail: string): Promise<boolean | ISpCoreResult>;
    isCurrentUserInGroup(groupName: string): Promise<any>;
}
export declare const SpUserPermissions: Permissions;
export {};
