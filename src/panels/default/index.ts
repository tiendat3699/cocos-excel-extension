import packageJSON from "../../../package.json";
import { readFileSync } from "fs-extra";
import { join } from "path";
import { readFile, utils, WorkSheet } from "xlsx";

/**
 * @zh 如果希望兼容 3.3 之前的版本可以使用下方的代码
 * @en You can add the code below if you want compatibility with versions prior to 3.3
 */
// Editor.Panel.define = Editor.Panel.define || function(options: any) { return options }
module.exports = Editor.Panel.define({
    listeners: {
        show() {},
        hide() {},
    },
    template: readFileSync(
        join(__dirname, "../../../static/template/default/index.html"),
        "utf-8"
    ),
    style: readFileSync(
        join(__dirname, "../../../static/style/default/index.css"),
        "utf-8"
    ),
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
            const data = await Editor.Profile.getConfig(
                packageJSON.name,
                "excelToJsonData"
            );
            if (data) {
                const inputFile = this.$.excelFile as HTMLInputElement;
                const fileName = this.$.fileName as HTMLInputElement;
                const outputFile = this.$.out as HTMLInputElement;
                const sheetName = this.$.sheetName as HTMLInputElement;
                const exclude = this.$.exclude as HTMLInputElement;
                const range = this.$.range as HTMLInputElement;
                const blankRow = this.$.blankRow as HTMLInputElement;
                const blankCell = this.$.blankCell as HTMLInputElement;
                const useHeader = this.$.useHeader as HTMLInputElement;

                inputFile.value = data.inputFile;
                fileName.value = data.fileName;
                outputFile.value = data.outputFile;
                sheetName.value = data.sheetName;
                exclude.value = data.exclude;
                range.value = data.range;
                blankRow.value = data.blankRow;
                blankCell.value = data.blankCell;
                useHeader.value = data.useHeader;
            } else {
                setTimeout(() => {
                    //@ts-ignore
                    this.$.out.value = "project://assets";
                }, 100);
            }
        },

        async convertToJson(url: string) {
            const fileName = this.$.fileName as HTMLInputElement;
            const outputFile = this.$.out as HTMLInputElement;
            const sheetName = this.$.sheetName as HTMLInputElement;
            const exclude = this.$.exclude as HTMLInputElement;
            const range = this.$.range as HTMLInputElement;
            const blankRow = this.$.blankRow as HTMLInputElement;
            const blankCell = this.$.blankCell as HTMLInputElement;
            const useHeader = this.$.useHeader as HTMLInputElement;

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

            const getRow = (sheet: WorkSheet) => {
                return utils.sheet_to_json(sheet, {
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
                this.$.submit?.removeAttribute("disabled");
                this.$.submit?.removeAttribute("loading");
                return;
            }

            if (!data.fileName) {
                await Editor.Dialog.warn("Warning", {
                    detail: "Output name is required",
                });
                fileName.focus();
                this.$.submit?.removeAttribute("disabled");
                this.$.submit?.removeAttribute("loading");
                return;
            }

            if (
                excludeSheet.length > 1 &&
                excludeSheet.includes(data.sheetName)
            ) {
                await Editor.Dialog.warn("Warning", {
                    detail: `Sheet ${sheetName} in exclude sheet`,
                });
                sheetName.focus();
                this.$.submit?.removeAttribute("disabled");
                this.$.submit?.removeAttribute("loading");
                return;
            }

            try {
                const workBook = readFile(url, {
                    type: "binary",
                });
                let result: any = {};
                if (data.sheetName) {
                    const row = getRow(workBook.Sheets[data.sheetName]);
                    if (row.length > 0) result = row;
                } else {
                    workBook.SheetNames.forEach((name) => {
                        if (excludeSheet.includes(name)) return;
                        const row = getRow(workBook.Sheets[name]);
                        if (row.length > 0) result[name] = row;
                    });
                }

                const output = data.outputFile.replace("project://", "db://");
                await Editor.Message.request(
                    "asset-db",
                    "create-asset",
                    output + `/${data.fileName}.json`,
                    JSON.stringify(result)
                );
                this.$.submit?.removeAttribute("disabled");
                this.$.submit?.removeAttribute("loading");
                Editor.Profile.setConfig(
                    packageJSON.name,
                    "excelToJsonData",
                    data
                );
            } catch (e: any) {
                await Editor.Dialog.error("Error", { detail: e.message });
                this.$.submit?.removeAttribute("disabled");
                this.$.submit?.removeAttribute("loading");
            }
        },
    },
    ready() {
        this.loadFormData();
        this.$.submit?.addEventListener("confirm", (event) => {
            const inputFile = this.$.excelFile as HTMLInputElement;
            this.$.submit?.setAttribute("disabled", "true");
            this.$.submit?.setAttribute("loading", "true");
            Editor.Message.send(
                "excel-extension",
                "convertToJson",
                inputFile.value
            );
        });
    },
    beforeClose() {},
    close() {},
});
