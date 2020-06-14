import { ConfigOptions } from 'sp-pnp-js';
import { ISpCoreResult } from './core-services-setup.interface';
declare class Core {
    get config(): ConfigOptions;
    get jsonHeader(): string;
    get baseUrl(): string;
    set baseUrl(value: string);
    private _baseUrl;
    private readonly _jsonHeader;
    private readonly _config;
    setup(url?: string): PromiseLike<ISpCoreResult>;
    onError(reason: ISpCoreResult): ISpCoreResult;
    showErrorLog(reason: any): void;
}
export declare const SpCore: Core;
export {};
