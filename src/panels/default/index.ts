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
        out: "#out",
        progress: "#progress",
        submit: "#submit",
    },
    methods: {
        convertToJson(url: string) {
            const outputFile = this.$.out as HTMLInputElement;
            const workBook = readFile(url, { type: "binary" });
            const result: { [key: string]: any } = {};
            workBook.SheetNames.forEach((sheetName) => {
                const row = utils.sheet_to_json(workBook.Sheets[sheetName], {
                    raw: true,
                    rawNumbers: true,
                });
                if (row.length > 0) result[sheetName] = row;
            });
            console.log(result);
            writeFile(outputFile.value + "/data.json", JSON.stringify(result));
        },
    },
    ready() {
        this.$.submit?.addEventListener("confirm", (event) => {
            const inputFile = this.$.excelFile as HTMLInputElement;
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
