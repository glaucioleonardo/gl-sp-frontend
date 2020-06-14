"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var sp_pnp_js_1 = require("sp-pnp-js");
var Core = (function () {
    function Core() {
        this._baseUrl = '';
        this._jsonHeader = 'application/json;odata=verbose';
        this._config = {
            headers: { Accept: this._jsonHeader }
        };
    }
    Object.defineProperty(Core.prototype, "config", {
        get: function () { return this._config; },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Core.prototype, "jsonHeader", {
        get: function () { return this._jsonHeader; },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Core.prototype, "baseUrl", {
        get: function () { return this._baseUrl; },
        set: function (value) { this._baseUrl = value; },
        enumerable: true,
        configurable: true
    });
    Core.prototype.setup = function (url) {
        var _this = this;
        return new Promise(function (resolve, reject) {
            if (_this._baseUrl.trim().length === 0 && url == null) {
                reject(_this.onError({ code: 405, description: null, message: null }));
            }
            else {
                try {
                    sp_pnp_js_1.default.setup({
                        sp: {
                            headers: {
                                Accept: _this._jsonHeader,
                            },
                            baseUrl: url == null ? _this._baseUrl : url,
                        }
                    });
                    _this.baseUrl = url == null ? _this._baseUrl : url;
                    resolve({
                        code: 200,
                        description: 'The setup has been configured correctly.',
                    });
                }
                catch (reason) {
                    reject(_this.onError(reason));
                }
            }
        });
    };
    Core.prototype.onError = function (reason) {
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
    };
    Core.prototype.showErrorLog = function (reason) {
        var error = exports.SpCore.onError(reason);
        console.error("Error code: " + error.code);
        if (error.description != null) {
            console.error("Error description: " + error.description);
        }
        if (error.message) {
            console.error("Error message: " + error.message);
        }
    };
    return Core;
}());
exports.SpCore = new Core();
//# sourceMappingURL=core-services-setup.service.js.map