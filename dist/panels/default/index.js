"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
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
        out: "#out",
        progress: "#progress",
        submit: "#submit",
    },
    methods: {
        convertToJson(url) {
            const outputFile = this.$.out;
            const workBook = (0, xlsx_1.readFile)(url, { type: "binary" });
            const result = {};
            workBook.SheetNames.forEach((sheetName) => {
                const row = xlsx_1.utils.sheet_to_json(workBook.Sheets[sheetName], {
                    raw: true,
                    rawNumbers: true,
                });
                if (row.length > 0)
                    result[sheetName] = row;
            });
            console.log(result);
            (0, fs_extra_1.writeFile)(outputFile.value + "/data.json", JSON.stringify(result));
        },
    },
    ready() {
        var _a;
        (_a = this.$.submit) === null || _a === void 0 ? void 0 : _a.addEventListener("confirm", (event) => {
            const inputFile = this.$.excelFile;
            Editor.Message.send("excel-extension", "convertToJson", inputFile.value);
        });
    },
    beforeClose() { },
    close() { },
});