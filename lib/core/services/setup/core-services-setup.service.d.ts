import { ISpCoreResult } from './core-services-setup.interface';
import { IConfigOptions } from '@pnp/common';
declare class Core {
    private _baseUrl;
    private readonly _jsonHeader;
    private readonly _config;
    get config(): IConfigOptions;
    get jsonHeader(): string;
    get baseUrl(): string;
    set baseUrl(value: string);
    fetchHeader(): Promise<any>;
    getDigest(url: string): Promise<any>;
    setup(url?: string): Promise<ISpCoreResult>;
    private setupDigest;
    onError(reason: ISpCoreResult): ISpCoreResult;
    showErrorLog(reason: any): string;
}
export declare const SpCore: Core;
export {};
