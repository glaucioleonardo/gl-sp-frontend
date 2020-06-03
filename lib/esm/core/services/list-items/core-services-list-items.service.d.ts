import { AttachmentFileAddResult, AttachmentFileInfo, ItemAddResult, TypedHash } from 'sp-pnp-js';
import { IAttachmentData, IAttachmentServerData, IListDatabaseResults } from './core-services-list-items.interface';
import { AttachmentFile } from 'sp-pnp-js/lib/sharepoint/attachmentfiles';
declare class Core {
    /**
     * Retrieve a search list based on fields, filter and ordering
     * @param listName
     * @param fieldsArray (optional)
     */
    retrieve(listName: string, fieldsArray?: string[]): Promise<any[]>;
    /**
     * Move items to recycle bin. The user will be able to restore the information.
     * @param listItemId
     * @param listName
     */
    recycle(listItemId: number, listName: string): Promise<unknown>;
    /**
     * Delete items permanently. The user will not be able to restore the information.
     * @param listItemId
     * @param listName
     */
    delete(listItemId: number, listName: string): Promise<unknown>;
    /**
     * Adds a new item to the collection
     * @param listName
     * @param data
     */
    add(listName: string, data: TypedHash<any>): Promise<ItemAddResult>;
    /**
     * Update a new item to the collection
     * @param listItemId
     * @param listName
     * @param data
     */
    update(listItemId: number, listName: string, data: TypedHash<any>): Promise<ItemAddResult>;
}
export declare const ListItemsCore: Core;
declare class Attachment {
    /**
     * Before using this method, you need defining the base url on setup (SpCore).
     * Adds multiple new attachment to the collection. Not supported for batching.
     * @param listItemId
     * @param listName
     * @param attachments
     */
    add(listItemId: number, listName: string, attachments: AttachmentFileInfo[]): Promise<any>;
    /**
     * Before using this method, you need defining the base url on setup (SpCore).
     * Delete multiple attachments from the collection. Not supported for batching.
     * @param listItemId
     * @param listName
     * @param attachments
     */
    delete(listItemId: number, listName: string, attachments: string[]): Promise<any>;
    /**
     * Before using this method, you need defining the base url on setup (SpCore).
     * @param listItemId
     * @param listName
     */
    retrieve(listItemId: number, listName: string): Promise<IAttachmentServerData[]>;
    /**
     * Before using this method, you need defining the base url on setup (SpCore).
     * This method is intended to retrieve all attachments inside a list item and prepare the list to use directly to the user interface (binding)
     * @param listItemId
     * @param listName
     */
    retrieveForBinding(listItemId: number, listName: string): Promise<IAttachmentData[]>;
    /**
     * Before using this method, you need defining the base url on setup (SpCore).
     * @param listItemId
     * @param listName
     * @param fileName Without extension, e.g. "attachment"
     */
    retrieveTxtContent(listItemId: number, listName: string, fileName?: string): Promise<string>;
    /**
     * Sets the content of a file. Not supported for batching
     * @param listItemId
     * @param listName
     * @param fileName Without extension, e.g. "attachment"
     * @param content Content to be added to the text
     */
    setTxtContent(listItemId: number, listName: string, fileName: string | undefined, content: string): Promise<AttachmentFile>;
    /**
     * Adds a new attachment to the collection. Not supported for batching.
     * @param listItemId
     * @param listName
     * @param fileName Without extension, e.g. "attachment"
     * @param content Content to be added to the text
     */
    addTxtContent(listItemId: number, listName: string, fileName: string | undefined, content: string): Promise<AttachmentFileAddResult>;
}
export declare const ListItemsAttachment: Attachment;
declare class Search {
    /**
     * Retrieve a search list based on fields, filter and ordering
     * @param listName
     * @param fieldsArray (optional)
     * @param filter (optional)
     * @param orderBy (optional)
     * @param ascending (optional)
     */
    retrieveSearch(listName: string, fieldsArray?: string[], filter?: string, orderBy?: string, ascending?: boolean): Promise<IListDatabaseResults[]>;
}
export declare const ListItemsSearch: Search;
export {};
