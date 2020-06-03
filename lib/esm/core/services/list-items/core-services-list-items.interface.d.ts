import { IDeferred, IMetadata } from '../../..';
import { AttachmentFile } from 'sp-pnp-js/lib/sharepoint/attachmentfiles';
export interface IAttachmentServerData {
    __metadata: IAttachmentMetadata;
    FileName: string;
    ServerRelativeUrl: string;
}
interface IAttachmentMetadata {
    id: string;
    uri: string;
    type: string;
}
export interface IAttachmentData {
    id: number;
    new: boolean;
    remove: boolean;
    name: string;
    file?: File;
    url?: string;
    icon?: string;
}
export interface IListDatabaseResults {
    __metadata: IMetadata;
    AttachmentFiles: IDeferred;
    FirstUniqueAncestorSecurableObject: IDeferred;
    RoleAssignments: IDeferred;
    ContentTypes: IDeferred;
    DefaultView: IDeferred;
    EventReceivers: IDeferred;
    Fields: IDeferred;
    Forms: IDeferred;
    InformationRightsManagementSettings: IDeferred;
    Items: IDeferred;
    ParentWeb: IDeferred;
    RootFolder: IDeferred;
    UserCustomActions: IDeferred;
    Views: IDeferred;
    WorkflowAssociations: IDeferred;
    AllowContentTypes: boolean;
    BaseTemplate: number;
    BaseType: number;
    ContentTypesEnabled: boolean;
    Created: string;
    DefaultContentApprovalWorkflowId: string;
    Description: string;
    Direction: string;
    DocumentTemplateUrl: any;
    DraftVersionVisibility: number;
    EnableAttachments: boolean;
    EnableFolderCreation: boolean;
    EnableMinorVersions: boolean;
    EnableModeration: boolean;
    EnableVersioning: boolean;
    EntityTypeName: string;
    ForceCheckout: boolean;
    HasExternalDataSource: boolean;
    Hidden: boolean;
    Id: number;
    ImageUrl: string;
    IrmEnabled: boolean;
    IrmExpire: boolean;
    IrmReject: boolean;
    IsApplicationList: boolean;
    IsCatalog: boolean;
    IsPrivate: boolean;
    ItemCount: number;
    LastItemDeletedDate: string;
    LastItemModifiedDate: string;
    ListItemEntityTypeFullName: string;
    MajorVersionLimit: number;
    MajorWithMinorVersionsLimit: number;
    MultipleDataList: boolean;
    NoCrawl: boolean;
    ParentWebUrl: string;
    ServerTemplateCanCreateFolders: boolean;
    TemplateFeatureId: string;
    Title: string;
    Attachments: boolean;
}
export interface IAttachmentFileInfo {
    name: string;
    content: string | Blob | ArrayBuffer;
}
export interface IAttachmentFileAddResult {
    file: AttachmentFile;
    data: any;
}
export {};
