"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const package_json_1 = __importDefault(require("../../../package.json"));
const fs_extra_1 = require("fs-extra");
const path_1 = require("path");
const xlsx_1 = require("xlsx");
/**
 * @zh 如果希望兼容 3.3 之前的版本可以使用下方的代码
 * @en You can add the code below if you want compatibility with versions prior to 3.3
 */
// Editor.Panel.define = Editor.Panel.define || function(options: any) { return options }
module.exports = Editor.Panel.define({
    listeners: {
        show() { },
        hide() { },
    },
    template: (0, fs_extra_1.readFileSync)((0, path_1.join)(__dirname, "../../../static/template/default/index.html"), "utf-8"),
    style: (0, fs_extra_1.readFileSync)((0, path_1.join)(__dirname, "../../../static/style/default/index.css"), "utf-8"),
    $: {
        excelFile: "#excelAsset",
        fileName: "#fileName",
        out: "#out",
        sheetName: "#sheetName",
        submit: "#submit",
        range: "#range",
        exclude: "#exclude",
        blankRow: "#blankRow",
        blankCell: "#blankCell",
        useHeader: "#useHeader",
    },
    methods: {
        async loadFormData() {
            const data = await Editor.Profile.getConfig(package_json_1.default.name, "excelToJsonData");
            if (data) {
                const inputFile = this.$.excelFile;
                const fileName = this.$.fileName;
                const outputFile = this.$.out;
                const sheetName = this.$.sheetName;
                const exclude = this.$.exclude;
                const range = this.$.range;
                const blankRow = this.$.blankRow;
                const blankCell = this.$.blankCell;
                const useHeader = this.$.useHeader;
                inputFile.value = data.inputFile;
                fileName.value = data.fileName;
                outputFile.value = data.outputFile;
                sheetName.value = data.sheetName;
                exclude.value = data.exclude;
                range.value = data.range;
                blankRow.value = data.blankRow;
                blankCell.value = data.blankCell;
                useHeader.value = data.useHeader;
            }
            else {
                setTimeout(() => {
                    //@ts-ignore
                    this.$.out.value = "project://assets";
                }, 100);
            }
        },
        async convertToJson(url) {
            var _a, _b, _c, _d, _e, _f, _g, _h, _j, _k;
            const fileName = this.$.fileName;
            const outputFile = this.$.out;
            const sheetName = this.$.sheetName;
            const exclude = this.$.exclude;
            const range = this.$.range;
            const blankRow = this.$.blankRow;
            const blankCell = this.$.blankCell;
            const useHeader = this.$.useHeader;
            const data = {
                inputFile: url,
                fileName: fileName.value,
                outputFile: outputFile.value,
                sheetName: sheetName.value,
                blankRow: blankRow.value,
                blankCell: blankCell.value,
                exclude: exclude.value,
                range: range.value,
                useHeader: useHeader.value,
            };
            const excludeSheet = data.exclude.split(/\s*,\s*/);
            const getRow = (sheet) => {
                return xlsx_1.utils.sheet_to_json(sheet, {
                    raw: true,
                    rawNumbers: true,
                    defval: !!data.blankCell ? null : undefined,
                    blankrows: !!data.blankRow,
                    range: !!data.range ? data.range : undefined,
                    header: !!data.useHeader ? undefined : 1,
                });
            };
            if (outputFile.getAttribute("invalid")) {
                await Editor.Dialog.warn("Warning", {
                    detail: "Output path invalid",
                });
                outputFile.focus();
                (_a = this.$.submit) === null || _a === void 0 ? void 0 : _a.removeAttribute("disabled");
                (_b = this.$.submit) === null || _b === void 0 ? void 0 : _b.removeAttribute("loading");
                return;
            }
            if (!data.fileName) {
                await Editor.Dialog.warn("Warning", {
                    detail: "Output name is required",
                });
                fileName.focus();
                (_c = this.$.submit) === null || _c === void 0 ? void 0 : _c.removeAttribute("disabled");
                (_d = this.$.submit) === null || _d === void 0 ? void 0 : _d.removeAttribute("loading");
                return;
            }
            if (excludeSheet.length > 1 &&
                excludeSheet.includes(data.sheetName)) {
                await Editor.Dialog.warn("Warning", {
                    detail: `Sheet ${sheetName} in exclude sheet`,
                });
                sheetName.focus();
                (_e = this.$.submit) === null || _e === void 0 ? void 0 : _e.removeAttribute("disabled");
                (_f = this.$.submit) === null || _f === void 0 ? void 0 : _f.removeAttribute("loading");
                return;
            }
            try {
                const workBook = (0, xlsx_1.readFile)(url, {
                    type: "binary",
                });
                let result = {};
                if (data.sheetName) {
                    const row = getRow(workBook.Sheets[data.sheetName]);
                    if (row.length > 0)
                        result = row;
                }
                else {
                    workBook.SheetNames.forEach((name) => {
                        if (excludeSheet.includes(name))
                            return;
                        const row = getRow(workBook.Sheets[name]);
                        if (row.length > 0)
                            result[name] = row;
                    });
                }
                const output = data.outputFile.replace("project://", "db://");
                await Editor.Message.request("asset-db", "create-asset", output + `/${data.fileName}.json`, JSON.stringify(result));
                (_g = this.$.submit) === null || _g === void 0 ? void 0 : _g.removeAttribute("disabled");
                (_h = this.$.submit) === null || _h === void 0 ? void 0 : _h.removeAttribute("loading");
                Editor.Profile.setConfig(package_json_1.default.name, "excelToJsonData", data);
            }
            catch (e) {
                await Editor.Dialog.error("Error", { detail: e.message });
                (_j = this.$.submit) === null || _j === void 0 ? void 0 : _j.removeAttribute("disabled");
                (_k = this.$.submit) === null || _k === void 0 ? void 0 : _k.removeAttribute("loading");
            }
        },
    },
    ready() {
        var _a;
        this.loadFormData();
        (_a = this.$.submit) === null || _a === void 0 ? void 0 : _a.addEventListener("confirm", (event) => {
            var _a, _b;
            const inputFile = this.$.excelFile;
            (_a = this.$.submit) === null || _a === void 0 ? void 0 : _a.setAttribute("disabled", "true");
            (_b = this.$.submit) === null || _b === void 0 ? void 0 : _b.setAttribute("loading", "true");
            Editor.Message.send("excel-extension", "convertToJson", inputFile.value);
        });
    },
    beforeClose() { },
    close() { },
});
