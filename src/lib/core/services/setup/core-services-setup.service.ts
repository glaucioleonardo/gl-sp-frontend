import { sp } from "@pnp/sp";
import { ISpCoreResult } from './core-services-setup.interface';
import { IConfigOptions } from '@pnp/common';
import { IContextInfo } from '@pnp/sp/sites';

class Core {
  private _baseUrl = '';
  private readonly _jsonHeader = 'application/json;odata=verbose';
  private readonly _config: IConfigOptions = {
    headers: { accept: this._jsonHeader }
  }

  get config(): IConfigOptions { return this._config; }
  get jsonHeader(): string { return this._jsonHeader; }
  get baseUrl(): string { return this._baseUrl; }
  set baseUrl(value: string) { this._baseUrl = value; }

  async fetchHeader() {
    const digest = await this.getDigest(this._baseUrl)

    return {
      headers: {
        Accept: 'application/json;odata=verbose',
        'X-RequestDigest': digest
      },
      credentials: 'include'
    } as any;
  }

  async getDigest(url: string): Promise<any> {
    const context: IContextInfo = await sp.configure(this.config, url).site.getContextInfo();
    return  context.FormDigestValue;
  }

  setup(url?: string): Promise<ISpCoreResult> {
    return new Promise(async (resolve, reject) => {
      if (this._baseUrl.trim().length === 0 && url == null) {
        reject(this.onError({ code: 405, description: null, message: null }));
      } else {
        this._baseUrl = url as string;

        try {
          sp.setup({
            sp: {
              headers: {
                Accept: this._jsonHeader,
                // 'X-RequestDigest': digest
              },
              baseUrl: url == null ? this._baseUrl : url
            }
          })

          this.baseUrl = url == null ? this._baseUrl : url;

          const base: string = url == null ? this._baseUrl : url;
          await this.setupDigest(base);

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
  private setupDigest(url: string): Promise<ISpCoreResult> {
    return new Promise(async (resolve, reject) => {
      if (this._baseUrl.trim().length === 0 && url == null) {
        reject(this.onError({ code: 405, description: null, message: null }));
      } else {
        this._baseUrl = url;
        const digest = await this.getDigest(url as string);

        try {
          sp.setup({
            sp: {
              headers: {
                Accept: this._jsonHeader,
                'X-RequestDigest': digest
              },
              baseUrl: url == null ? this._baseUrl : url
            }
          })

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

  showErrorLog(reason: any): string {
    const error = SpCore.onError(reason);
    let errorMessage: string = `Error code: ${error.code}`;

    console.error(errorMessage);

    if (error.description != null) {
      errorMessage = `Error description: ${error.description}`;
      console.error(errorMessage);
      return errorMessage;
    }

    if (error.message) {
      errorMessage = `Error message: ${error.message}`;
      console.error(errorMessage);
      return errorMessage;
    }

    return errorMessage;
  }
}

export const SpCore = new Core();
