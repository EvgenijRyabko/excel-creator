import * as excel from "exceljs";
import { writeFile, mkdir, access } from "fs/promises";
import data from "./data/wrong_passport_and_birthday.json" assert { type: "json" };

const isDirExist = async (path) => {
  try {
    await access("path");
    return true;
  } catch (e) {
    return false;
  }
};

const prepareData = (separator, columns, data) => {
  const columnIndex = Object.keys(data[0]).findIndex((el) => el === separator);
  const columnName = Object.keys(data[0]).find((el) => el === separator);

  columns.splice(columnIndex, 1);

  const newData = {};

  for (const el of data) {
    const newKey = el[columnName];

    delete el[columnName];

    if (!Object.keys(newData).find((key) => key === newKey)) {
      newData[newKey] = [];

      newData[newKey].push(el);
    } else {
      newData[newKey].push(el);
    }
  }

  return newData;
};

const generateExcel = async (fileName, columns, data) => {
  const workbook = new excel.default.Workbook();
  const worksheet = workbook.addWorksheet("new");

  columns.forEach((column, colIndex) => {
    const cell = worksheet.getCell(1, colIndex + 1);
    cell.value = column;
    // изменить ширину столбцов
    worksheet.getColumn(colIndex + 1).width = 16;
  });

  let rowIndex = 2;

  for (const student of data) {
    let colIndex = 1;

    for (const key in student) {
      worksheet.getCell(rowIndex, colIndex).value = student[key];

      colIndex++;
    }

    rowIndex++;
  }

  const buffer = await workbook.xlsx.writeBuffer();

  const dirExist = await isDirExist("./storage");

  if (!dirExist) await mkdir("./storage", { recursive: true });

  await writeFile(`./storage/${fileName}.xlsx`, Buffer.from(buffer));
};

// generateExcel(
//   "Студенты с неправильной датой рождения или паспортом",
//   data.columns,
//   data.data
// );

const newData = prepareData("faculty", data.columns, data.data);

for (const key in newData) {
  generateExcel(key, data.columns, newData[key]);
}
