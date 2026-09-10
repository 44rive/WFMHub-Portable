const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const root = path.resolve(__dirname, "../../..");
const output = path.join(root, "build", "office-script-test");
const writer = fs.readFileSync(path.join(output, "WFM_Write_Absence.js"), "utf8");
const tests = fs.readFileSync(path.join(output, "WFM_Write_Absence.test.js"), "utf8");
vm.runInThisContext(`${writer}\n${tests}`, {filename: "WFM_Write_Absence.combined.test.js"});
