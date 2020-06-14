import { AttachmentFileAddResult, AttachmentFileInfo, ItemAddResult, TypedHash } from 'sp-pnp-js';
import { IAttachmentData, IAttachmentServerData, IListDatabaseResults } from './core-services-list-items.interface';
import { AttachmentFile } from 'sp-pnp-js/lib/sharepoint/attachmentfiles';
declare class Core {
    fieldsToStringArray(fields: string[]): string;
    retrieve(listName: string, fieldsArray?: string[], baseUrl?: string): Promise<any[]>;
    retrieveSingle(listItemId: number, listName: string, fieldsArray?: string[], baseUrl?: string): Promise<any>;
    recycle(listItemId: number, listName: string, baseUrl?: string): Promise<unknown>;
    delete(listItemId: number, listName: string, baseUrl?: string): Promise<unknown>;
    add(listName: string, data: TypedHash<any>, baseUrl?: string): Promise<ItemAddResult>;
    update(listItemId: number, listName: string, data: TypedHash<any>, baseUrl?: string): Promise<ItemAddResult>;
}
export declare const ListItemsCore: Core;
declare class Attachment {
    add(listItemId: number, listName: string, attachments: AttachmentFileInfo[], baseUrl?: string): Promise<any>;
    delete(listItemId: number, listName: string, attachments: string[], baseUrl?: string): Promise<any>;
    retrieve(listItemId: number, listName: string, baseUrl?: string): Promise<IAttachmentServerData[]>;
    retrieveForBinding(listItemId: number, listName: string, baseUrl?: string): Promise<IAttachmentData[]>;
    retrieveTxtContent(listItemId: number, listName: string, fileName?: string, baseUrl?: string): Promise<string>;
    setTxtContent(listItemId: number, listName: string, fileName: string, content: string, baseUrl?: string): Promise<AttachmentFile>;
    addTxtContent(listItemId: number, listName: string, fileName: string, content: string, baseUrl?: string): Promise<AttachmentFileAddResult>;
}
export declare const ListItemsAttachment: Attachment;
declare class Search {
    retrieveSearch(listName: string, fieldsArray?: string[], filter?: string, orderBy?: string, ascending?: boolean, baseUrl?: string): Promise<IListDatabaseResults[]>;
}
export declare const ListItemsSearch: Search;
export {};
