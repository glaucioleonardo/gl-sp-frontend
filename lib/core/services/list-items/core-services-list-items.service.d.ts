import { IAttachment, IAttachmentAddResult, IAttachmentFileInfo, IAttachmentInfo } from '@pnp/sp/attachments/index';
import { IAttachmentBlob, IAttachmentData, IAttachmentMultipleBlobs, IListDatabaseResults, ItemAddResult } from './core-services-list-items.interface';
import "@pnp/sp/attachments";
import { ITypedHash } from '@pnp/common';
import { ISpCoreResult } from '../setup/core-services-setup.interface';
import { IComboBoxData } from 'gl-w-combobox-frontend';
declare class Core {
    fieldsToStringArray(fields: string[]): string;
    retrieve(listName: string, fieldsArray?: string[], baseUrl?: string): Promise<any[]>;
    retrieveExternal(listName: string, fieldsArray: string[] | undefined, baseUrl: string, top?: number): Promise<any[]>;
    retrieveExternalSingle(id: number, listName: string, fieldsArray: string[] | undefined, baseUrl: string): Promise<any[]>;
    retrieveForCombobox(listName: string, baseUrl: string, valueField?: string, textField?: string): Promise<IComboBoxData[]>;
    retrieveExternalForCombobox(listName: string, baseUrl: string, valueField?: string, textField?: string, top?: number): Promise<IComboBoxData[]>;
    retrieveSingle(listItemId: number, listName: string, fieldsArray?: string[], baseUrl?: string): Promise<any>;
    recycle(listItemId: number, listName: string, baseUrl?: string): Promise<unknown>;
    recycleDuplicated(listName: string, field: string, baseUrl?: string): Promise<any[]>;
    delete(listItemId: number, listName: string, baseUrl?: string): Promise<unknown>;
    deleteDuplicated(listName: string, field: string, baseUrl?: string): Promise<any[]>;
    add(listName: string, data: ITypedHash<any>, baseUrl?: string): Promise<ItemAddResult | ISpCoreResult>;
    update(listItemId: number, listName: string, data: ITypedHash<any>, baseUrl?: string): Promise<ItemAddResult>;
}
export declare const ListItemsCore: Core;
declare class Attachment {
    add(listItemId: number, listName: string, attachments: IAttachmentFileInfo[], baseUrl?: string): Promise<any>;
    delete(listItemId: number, listName: string, attachments: string[], baseUrl?: string): Promise<any>;
    retrieve(listItemId: number, listName: string, baseUrl?: string): Promise<IAttachmentInfo[]>;
    retrieveBlob(listItemId: number, listName: string, fileName: string, baseUrl?: string): Promise<Blob | null>;
    retrieveMultipleBlobSameItem(listItemId: number, listName: string, fileNames: string[], baseUrl?: string): Promise<IAttachmentBlob[]>;
    retrieveMultipleBlob(data: IAttachmentMultipleBlobs[], baseUrl?: string): Promise<IAttachmentBlob[]>;
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
