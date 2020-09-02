import { sp } from "@pnp/sp";
class Core {
    constructor() {
        this._baseUrl = '';
        this._jsonHeader = 'application/json;odata=verbose';
        this._config = {
            headers: { accept: this._jsonHeader }
        };
    }
    get config() { return this._config; }
    get jsonHeader() { return this._jsonHeader; }
    get baseUrl() { return this._baseUrl; }
    set baseUrl(value) { this._baseUrl = value; }
    setup(url) {
        return new Promise((resolve, reject) => {
            if (this._baseUrl.trim().length === 0 && url == null) {
                reject(this.onError({ code: 405, description: null, message: null }));
            }
            else {
                try {
                    sp.setup({
                        sp: {
                            headers: {
                                Accept: this._jsonHeader
                            },
                            baseUrl: url == null ? this._baseUrl : url
                        }
                    });
                    this.baseUrl = url == null ? this._baseUrl : url;
                    resolve({
                        code: 200,
                        description: 'The setup has been configured correctly.',
                    });
                }
                catch (reason) {
                    reject(this.onError(reason));
                }
            }
        });
    }
    onError(reason) {
        if (reason.code != null && reason.code === 405) {
            return {
                code: 405,
                message: 'This request do not allow baseUrl property as null. You need informing it before continuing.',
                description: 'Method Not Allowed',
            };
        }
        else if (reason.message != null && reason.message === 'Unexpected token < in JSON at position 0') {
            return {
                code: 522,
                message: 'Unexpected token < in JSON at position 0.',
                description: 'Connection Timed Out',
            };
        }
        else if (reason.message != null && reason.message === 'Error making HttpClient request in queryable: [404] Not Found') {
            return {
                code: 404,
                message: 'Error making HttpClient request in queryable: [404] Not Found',
                description: 'Server not found',
            };
        }
        else {
            return {
                code: 500,
                message: reason.message == null ? null : reason.message,
                description: 'Internal Server Error',
            };
        }
    }
    showErrorLog(reason) {
        const error = SpCore.onError(reason);
        console.error(`Error code: ${error.code}`);
        if (error.description != null) {
            console.error(`Error description: ${error.description}`);
        }
        if (error.message) {
            console.error(`Error message: ${error.message}`);
        }
    }
}
export const SpCore = new Core();
//# sourceMappingURL=core-services-setup.service.js.map