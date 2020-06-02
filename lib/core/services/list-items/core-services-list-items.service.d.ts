import { IAttachmentData, IAttachmentServerData } from './core-services-list-items.interface';
declare class Attachment {
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
}
export declare const ListItemsAttachment: Attachment;
export {};
