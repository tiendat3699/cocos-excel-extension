import packageJSON from "../../../package.json";
import { readFileSync, writeFile } from "fs-extra";
import { join } from "path";
import { readFile, utils } from "xlsx";

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
        blankRow: "#blankRow",
        blankCell: "#blankCell",
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
                const blankRow = this.$.blankRow as HTMLInputElement;
                const blankCell = this.$.blankCell as HTMLInputElement;

                inputFile.value = data.inputFile;
                fileName.value = data.fileName;
                outputFile.value = data.outputFile;
                sheetName.value = data.sheetName;
                blankRow.value = data.blankRow;
                blankCell.value = data.blankCell;
            }
        },

        convertToJson(url: string) {
            const fileName = this.$.fileName as HTMLInputElement;
            const outputFile = this.$.out as HTMLInputElement;
            const sheetName = this.$.sheetName as HTMLInputElement;
            const blankRow = this.$.blankRow as HTMLInputElement;
            const blankCell = this.$.blankCell as HTMLInputElement;
            const data = {
                inputFile: url,
                fileName: fileName.value,
                outputFile: outputFile.value,
                sheetName: sheetName.value,
                blankRow: blankRow.value,
                blankCell: blankCell.value,
            };

            const workBook = readFile(url, { type: "binary" });
            let result: any = {};
            if (data.sheetName) {
                const row = utils.sheet_to_json(
                    workBook.Sheets[data.sheetName],
                    {
                        raw: true,
                        rawNumbers: true,
                        defval: !!data.blankCell ? null : undefined,
                        blankrows: !!data.blankRow,
                    }
                );
                if (row.length > 0) result = row;
            } else {
                workBook.SheetNames.forEach((name) => {
                    const row = utils.sheet_to_json(workBook.Sheets[name], {
                        raw: true,
                        rawNumbers: true,
                        defval: !!data.blankCell ? null : undefined,
                        blankrows: !!data.blankRow,
                    });
                    if (row.length > 0) result[name] = row;
                });
            }

            writeFile(
                data.outputFile + `/${data.fileName}.json`,
                JSON.stringify(result)
            ).then(() => {
                this.$.submit?.removeAttribute("disabled");
                this.$.submit?.removeAttribute("loading");
                Editor.Profile.setConfig(
                    packageJSON.name,
                    "excelToJsonData",
                    data
                );
            });
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
