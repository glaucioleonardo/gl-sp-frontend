import { IAttachment, IAttachmentAddResult, IAttachmentFileInfo, IAttachmentInfo } from '@pnp/sp/attachments';
import { IAttachmentData, IListDatabaseResults, ItemAddResult } from './core-services-list-items.interface';
import "@pnp/sp/attachments";
import { ITypedHash } from '@pnp/common';
declare class Core {
    fieldsToStringArray(fields: string[]): string;
    retrieve(listName: string, fieldsArray?: string[], baseUrl?: string): Promise<any[]>;
    retrieveSingle(listItemId: number, listName: string, fieldsArray?: string[], baseUrl?: string): Promise<any>;
    recycle(listItemId: number, listName: string, baseUrl?: string): Promise<unknown>;
    delete(listItemId: number, listName: string, baseUrl?: string): Promise<unknown>;
    add(listName: string, data: ITypedHash<any>, baseUrl?: string): Promise<ItemAddResult>;
    update(listItemId: number, listName: string, data: ITypedHash<any>, baseUrl?: string): Promise<ItemAddResult>;
}
export declare const ListItemsCore: Core;
declare class Attachment {
    add(listItemId: number, listName: string, attachments: IAttachmentFileInfo[], baseUrl?: string): Promise<any>;
    delete(listItemId: number, listName: string, attachments: string[], baseUrl?: string): Promise<any>;
    retrieve(listItemId: number, listName: string, baseUrl?: string): Promise<IAttachmentInfo[]>;
    retrieveForBinding(listItemId: number, listName: string, baseUrl?: string): Promise<IAttachmentData[]>;
    retrieveTxtContent(listItemId: number, listName: string, fileName?: string, baseUrl?: string): Promise<string>;
    setTxtContent(listItemId: number, listName: string, fileName: string | undefined, content: string, baseUrl?: string): Promise<IAttachment>;
    addTxtContent(listItemId: number, listName: string, fileName: string | undefined, content: string, baseUrl?: string): Promise<IAttachmentAddResult>;
}
export declare const ListItemsAttachment: Attachment;
declare class Search {
    retrieveSearch(listName: string, fieldsArray?: string[], filter?: string, orderBy?: string, ascending?: boolean, baseUrl?: string): Promise<IListDatabaseResults[]>;
    retrieveSearchLimited(listName: string, fieldsArray?: string[], maxItems?: number, filter?: string, orderBy?: string, ascending?: boolean, baseUrl?: string): Promise<IListDatabaseResults[]>;
}
export declare const ListItemsSearch: Search;
export {};
