import xlsx from "node-xlsx";
const dataRaw = xlsx.parse("tkb.xls");
import * as moment from "moment";
import { google } from "googleapis";
import { getAuth } from "./src/utils/getAuth";
const { data } = dataRaw[0];
const dataFinal = [];
data.forEach((s, index) => {
  if (s[0] && s[1] && s[3] && s[10] && s[0] !== "Thứ") {
    dataFinal.push(s);
    // console.log(index)
  }
});
// console.log(dataFinal)
const colorIds = [
  "1",
  "2",
  "3",
  "4",
  "5",
  "6",
  "7",
  "8",
  "9",
  "10",
  "11",
  "2",
  "3",
  "4",
];

// function convertColorId(s: Final) {
//     switch(s.fullName) {
//         case ''
//     }
// }
const mon = [];
const dataFinalTest = [...dataFinal];
let colorIndex = -1;
while (dataFinalTest.length > 0) {
  if (mon.findIndex((s) => s.name === dataFinalTest[0][4]) === -1) {
    colorIndex++;
    mon.push({ name: dataFinalTest[0][4], colorId: colorIds[colorIndex] });
  }
  dataFinalTest.splice(0, 1);
}
function convertTimePart(s: string) {
  switch (s) {
    case "1,2,3":
      return { start_time: "07:00:00", end_time: "09:25:00" };
    case "4,5,6":
      return { start_time: "09:35:00", end_time: "12:00:00" };
    case "7,8,9":
      return { start_time: "12:30:00", end_time: "14:55:00" };
    case "10,11,12":
      return { start_time: "15:05:00", end_time: "17:30:00" };
    case "13,14,15,16":
      return { start_time: "18:00:00", end_time: "21:15:00" };
    case "1,2,3,4,5,6":
      return { start_time: "07:00:00", end_time: "12:00:00" };
    case "7,8,9,10,11,12":
      return { start_time: "12:30:00", end_time: "17:30:00" };
    default:
      return { start_time: "07:00:00", end_time: "09:25:00" };
  }
}
function convertDayOfWeek(s: string): number {
  switch (s) {
    case "2":
      return 1;
    case "3":
      return 2;
    case "4":
      return 3;
    case "5":
      return 4;
    case "6":
      return 5;
    case "7":
      return 6;
    case "CN":
      return 7;
    default:
      return 1;
  }
}
interface DataTest {
  dayOfWeek: string;
  idSubject: string;
  name: string;
  fullName: string;
  gv: string;
  part: string;
  location: string;
  time: string;
  colorId: string;
}

interface Final {
  summary: string;
  location: string;
  colorId: string;
  start: string;
  end: string;
  description: string;
}

const full: Final[] = [];
const dataTest: DataTest[] = [];
dataFinal.forEach((element: any[]) => {
  const monIndex = mon.findIndex((s) => s.name === element[4]);
  let colorId = mon[monIndex].colorId;
  dataTest.push({
    dayOfWeek: element[0],
    idSubject: element[1],
    name: element[3],
    fullName: element[4],
    gv: element[7],
    part: element[8],
    location: element[9],
    time: element[10],
    colorId,
  });
});
dataTest.forEach((d) => {
  const {
    time,
    dayOfWeek,
    idSubject,
    name,
    fullName,
    gv,
    part,
    location,
    colorId,
  } = d;
  const timeArr = time.split("-");
  const start = moment(timeArr[0], "DD/MM/YYYY");
  const end = moment(timeArr[1], "DD/MM/YYYY");
  const { start_time, end_time } = convertTimePart(part);
  const startTemp = start.clone();
  startTemp.set("day", convertDayOfWeek(dayOfWeek));
  while (startTemp.isSameOrBefore(end)) {
    const s = moment(
      startTemp.format("DD/MM/YYYY") + " " + start_time,
      "DD/MM/YYYY hh:mm:ss"
    ).format();
    const e = moment(
      startTemp.format("DD/MM/YYYY") + " " + end_time,
      "DD/MM/YYYY hh:mm:ss"
    ).format();
    // console.log(startTemp.format("DD/MM/YYYY"));
    full.push({
      summary: fullName,
      location,
      start: s,
      end: e,
      colorId,
      description: idSubject + "\n" + gv,
    });
    startTemp.set("day", startTemp.day() + 7);
  }
});
