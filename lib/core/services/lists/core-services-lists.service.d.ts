import { IListInfo } from '@pnp/sp/presets/core';
import '@pnp/sp/fields';
import '@pnp/sp/lists';
import '@pnp/sp/webs';
import '@pnp/sp/views';
import { ISpCoreResult } from '../setup/core-services-setup.interface';
import { IListProperties } from './core-services-lists.interface';
import { IFieldInfo } from '@pnp/sp/fields';
declare class Core {
    fieldsToStringArray(fields: string[]): string;
    exists(listName: string, baseUrl?: string): Promise<boolean>;
    retrieveSingle(listName: string, baseUrl?: string): Promise<IListInfo>;
    retrieve(listName: string, baseUrl?: string): Promise<IListInfo[]>;
    recycle(listName: string, baseUrl?: string): Promise<ISpCoreResult>;
    recreate(listName: string, baseUrl?: string, fields?: Partial<IFieldInfo>[], titleRequired?: boolean, properties?: IListProperties): Promise<ISpCoreResult>;
    addFields(listName: string, fields: Partial<IFieldInfo>[], baseUrl?: string): Promise<ISpCoreResult>;
    rename(listName: string, name: string, overwrite?: boolean, baseUrl?: string): Promise<ISpCoreResult>;
}
export declare const ListsCore: Core;
export {};
