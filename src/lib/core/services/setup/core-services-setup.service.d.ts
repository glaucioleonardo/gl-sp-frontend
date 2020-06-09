import { ISpCoreResult } from './core-services-setup.interface';
declare class Core {
    get baseUrl(): string;
    set baseUrl(value: string);
    private readonly _jsonHeader;
    private _baseUrl;
    setup(url?: string): PromiseLike<ISpCoreResult>;
    onError(reason: ISpCoreResult): ISpCoreResult;
    showErrorLog(reason: any): void;
}
export declare const SpCore: Core;
export {};
