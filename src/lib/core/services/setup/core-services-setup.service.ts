import { ConfigOptions, default as pnp } from 'sp-pnp-js';
import { ISpCoreResult } from './core-services-setup.interface';

class Core {
  // @ts-ignore
  get config(): ConfigOptions { return this._config; }
  // @ts-ignore
  get jsonHeader(): string { return this._jsonHeader; }
  // @ts-ignore
  get baseUrl(): string { return this._baseUrl; }
  // @ts-ignore
  set baseUrl(value: string) { this._baseUrl = value; }

  private _baseUrl = '';
  private readonly _jsonHeader = 'application/json;odata=verbose';
  private readonly _config: ConfigOptions = {
    headers: { Accept: this._jsonHeader }
  }

  setup(url?: string): PromiseLike<ISpCoreResult> {
    return new Promise((resolve, reject) => {
      if (this._baseUrl.trim().length === 0 && url == null) {
        reject(this.onError({ code: 405, description: null, message: null }));
      } else {
        try {
          pnp.setup({
            sp: {
              headers: {
                Accept: this._jsonHeader,
              },
              baseUrl: url == null ? this._baseUrl : url,
            }
          });

          this.baseUrl = url == null ? this._baseUrl : url;

          resolve({
            code: 200,
            description: 'The setup has been configured correctly.',
          });
        } catch (reason) {
          reject(this.onError(reason));
        }
      }
    });
  }

  onError(reason: ISpCoreResult): ISpCoreResult {
    if (reason.code != null && reason.code === 405) {
      return {
        code: 405,
        message: 'This request do not allow baseUrl property as null. You need informing it before continuing.',
        description: 'Method Not Allowed',
      };
    } else if (reason.message != null && reason.message === 'Unexpected token < in JSON at position 0') {
      return {
        code: 522,
        message: 'Unexpected token < in JSON at position 0.',
        description: 'Connection Timed Out',
      };
    } else if (reason.message != null && reason.message === 'Error making HttpClient request in queryable: [404] Not Found') {
      return {
        code: 404,
        message: 'Error making HttpClient request in queryable: [404] Not Found',
        description: 'Server not found',
      };
    } else {
      return {
        code: 500,
        // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment
        message: reason.message == null ? null : reason.message,
        description: 'Internal Server Error',
      };
    }
  }

  showErrorLog(reason: any) {
    const error = SpCore.onError(reason);
    console.error(`Error code: ${error.code}`);
    if (error.description != null) { console.error(`Error description: ${error.description}`); }
    if (error.message) { console.error(`Error message: ${error.message}`); }
  }
}

export const SpCore = new Core();
