export interface IMetadata {
    id: string;
    uri: string;
    type: string;
}
export interface ISpCurrentUserUserId {
    __metadata: ISpCurrentUserUserIdMetadata;
}
export interface ISpCurrentUserUserIdMetadata {
    type: string;
    NameId: string;
    NameIdIssuer: string;
}
export interface IDeferred {
    uri: string;
}
export interface IGroups {
    __deferred: IDeferred;
}
export interface ISpCoreResult {
    code: number;
    message?: string | null;
    description: string | null;
}
export interface ISpCurrentUser {
    __metadata: IMetadata;
    Groups: IGroups;
    Id: number;
    IsHiddenInUI: boolean;
    LoginName: string;
    Title: string;
    PrincipalType: number;
    Email: string;
    IsSiteAdmin: boolean;
    UserId: ISpCurrentUserUserId;
}
