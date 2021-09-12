var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
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
    fetchHeader() {
        return __awaiter(this, void 0, void 0, function* () {
            const digest = yield this.getDigest(this._baseUrl);
            return {
                headers: {
                    Accept: 'application/json;odata=verbose',
                    'X-RequestDigest': digest
                },
                credentials: 'include'
            };
        });
    }
    getDigest(url) {
        return __awaiter(this, void 0, void 0, function* () {
            const context = yield sp.configure(this.config, url).site.getContextInfo();
            return context.FormDigestValue;
        });
    }
    setup(url) {
        return new Promise((resolve, reject) => __awaiter(this, void 0, void 0, function* () {
            if (this._baseUrl.trim().length === 0 && url == null) {
                reject(this.onError({ code: 405, description: null, message: null }));
            }
            else {
                this._baseUrl = url;
                try {
                    sp.setup({
                        sp: {
                            headers: {
                                Accept: this._jsonHeader,
                            },
                            baseUrl: url == null ? this._baseUrl : url
                        }
                    });
                    this.baseUrl = url == null ? this._baseUrl : url;
                    const base = url == null ? this._baseUrl : url;
                    yield this.setupDigest(base);
                    resolve({
                        code: 200,
                        description: 'The setup has been configured correctly.',
                    });
                }
                catch (reason) {
                    reject(this.onError(reason));
                }
            }
        }));
    }
    setupDigest(url) {
        return new Promise((resolve, reject) => __awaiter(this, void 0, void 0, function* () {
            if (this._baseUrl.trim().length === 0 && url == null) {
                reject(this.onError({ code: 405, description: null, message: null }));
            }
            else {
                this._baseUrl = url;
                const digest = yield this.getDigest(url);
                try {
                    sp.setup({
                        sp: {
                            headers: {
                                Accept: this._jsonHeader,
                                'X-RequestDigest': digest
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
        }));
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
        let errorMessage = `Error code: ${error.code}`;
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
//# sourceMappingURL=core-services-setup.service.js.map