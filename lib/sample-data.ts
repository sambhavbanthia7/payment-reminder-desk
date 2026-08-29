import type { SheetTable } from "./reminders"

const duesHeaders = ["Sno", "Account Name", "A/C Code", "RepName", "Credit Days", "Dated", "Vrno", "Vrtype", "Dr/Cr", "New Entry Amt", "Adj Amt", "Bal", "Day"]

const invoices: Array<[string, string, number | null, string, string, number | string, number]> = [
  ["ABC TRADERS >> Manali", "A-320", 30, "14/06/26", "G-613-2627", 42961, 68],
  ["ABC TRADERS >> Manali", "A-320", 30, "07/07/26", "G-1204-2627", "18,430.00", 45],
  ["ABC TRADERS >> Manali", "A-320", 30, "22/07/26", "G-2627-2954", "82,506.00", 30],
  ["ABC TRADERS >> Manali", "A-320", 30, "23/07/26", "G-2627-2988", 1687, 29],
  ["ABC TRADERS >> Manali", "A-320", 30, "25/07/26", "G-2627-3067", "1,05,250.00", 27],
  ["ABC TRADERS >> Manali", "A-320", 30, "27/07/26", "GB-2627-783", 25039.5, 25],
  ["ABC TRADERS >> Manali", "A-320", 30, "29/07/26", "G-2627-3172", "9,984.00", 23],
  ["ABC TRADERS >> Manali", "A-320", 30, "30/07/26", "G-2627-3196", "2,17,155.00", 22],
  ["ABHILASH >> Manali", "A4398", 15, "06/08/26", "G-2627-3801", "16,360.00", 15],
  ["ABHILASH >> Manali", "A4398", 15, "07/08/26", "GB-2627-951", "23,600.00", 14],
  ["ABHILASH >> Manali", "A4398", 15, "11/08/26", "G-2627-3906", "<p style=\"color:red\">88,410.00</p>", 10],
  ["ABHILASH >> Manali", "A4398", 15, "09/08/26", "G-2627-3853", 0, 12],
  ["ABHILASH >> Manali", "A4398", 15, "10/08/26", "G-2627-3860", "-2,340.00", 11],
  ["ABHILASH >> Manali", "A4398", 15, "18/08/26", "G-2627-3964", "1,47,987.00", 3],
  ["ABHILASH >> Manali", "A4398", 15, "01/08/26", "G-2627-2701", "7,720.00", 20],
  ["ABNESH DHADWAL >> Manali", "A3833", 30, "24/07/26", "G-2627-3927", "1,69,519.00", 28],
  ["ABNESH DHADWAL >> Manali", "A3833", 30, "28/07/26", "G-2627-3930", "33,301.00", 24],
  ["ABNESH DHADWAL >> Manali", "A3833", 30, "10/07/26", "G-2526-9194", "1,852.00", 42],
  ["FUTURE ENERGIES- HARIPUR >> Manali", "F-11", 7, "14/08/26", "G-2627-3951", "3,107.00", 7],
  ["FUTURE ENERGIES- HARIPUR >> Manali", "F-11", 7, "18/08/26", "GB-2627-986", "40,877.00", 3],
  ["FUTURE ENERGIES- HARIPUR >> Manali", "F-11", 7, "12/08/26", "G-2627-3770", "11,126.00", 9],
  ["Keshav thakur >> Manali", "K-450", 0, "21/08/26", "GB-2627-987", "14,424.00", 0],
  ["Keshav thakur >> Manali", "K-450", 0, "16/08/26", "G-2627-3747", "6,638.00", 5],
  ["ADITYA ENTERPRISES >> Manali", "A4202", null, "16/08/26", "G-2627-3884", "41,950.00", 5],
  ["ADITYA ENTERPRISES >> Manali", "A4202", null, "12/07/26", "G-2627-2732", "1,83,286.00", 40],
  ["UNKNOWN TRADERS >> Kullu", "X-999", 30, "25/07/26", "G-2627-3919", "56,745.00", 27],
  ["UNKNOWN TRADERS >> Kullu", "X-999", 30, "28/07/26", "G-2627-3942", "23,394.00", 24],
  ["UNKNOWN TRADERS >> Kullu", "X-999", 30, "02/06/26", "G-2526-6352", "1,592.00", 80],
  ["ZENITH SUPPLIES >> Kullu", "Z-777", -1, "21/08/26", "G-2627-3999", "5,000.00", 0],
  ["ZENITH SUPPLIES >> Kullu", "Z-777", -1, "18/08/26", "G-2627-4001", "12,300.00", 3],
  ["SUNRISE SANITARY >> Kullu", "A4383", 300, "30/10/25", "G-AO-4618-2425", "1,29,715.00", 295],
  ["SUNRISE SANITARY >> Kullu", "A4383", 300, "27/10/25", "G-AO-4626-2425", "59,334.90", 298],
  ["SUNRISE SANITARY >> Kullu", "A4383", 300, "05/10/25", "G-AO-4627-2425", "2,692.00", 320],
]

export const sampleDues: SheetTable = {
  fileName: "Test_Dues_Analysis.xlsx",
  sheetName: "Worksheet",
  headers: duesHeaders,
  rows: invoices.map(([name, code, credit, dated, vrno, balance, day], index) => ({
    Sno: index + 1,
    "Account Name": name,
    "A/C Code": code,
    RepName: "Rishabh Dhiman, A.S.M., Himachal - 14RD",
    "Credit Days": credit,
    Dated: dated,
    Vrno: vrno,
    Vrtype: "GST INVOICE",
    "Dr/Cr": " Dr",
    "New Entry Amt": balance,
    "Adj Amt": 0,
    Bal: balance,
    Day: day,
  })),
}

const masterHeaders = ["Sno", "Party Code", "Name", "Mobile", "Email", "City", "State", "Credit Days", "Active"]

export const sampleMaster: SheetTable = {
  fileName: "Sample_Master_Data.xlsx",
  sheetName: "Customer List",
  headers: masterHeaders,
  rows: [
    [1, "A-320", "ABC TRADERS", 0, "rishabhdhiman5712@gmail.com", "Manali", "Himachal Pradesh", 30, "yes"],
    [2, "F-11", "FUTURE ENERGIES- HARIPUR", 0, null, "Manali", "Himachal Pradesh", 0, "yes"],
    [3, "A4398", "ABHILASH", 0, "rishabhdhiman5712@gmail.com", "Manali", "Himachal Pradesh", 0, "yes"],
    [4, "A3833", "ABNESH DHADWAL", null, "rishabhdhiman5712/gmail.com", "Manali", "Himachal Pradesh", 0, "yes"],
    [5, "K-450", "Keshav thakur", 0, null, "Manali", "Himachal Pradesh", 0, "yes"],
    [6, "A4202", "ADITYA ENTERPRISES", 0, null, "Manali", "Himachal Pradesh", 0, "yes"],
  ].map((values) => Object.fromEntries(masterHeaders.map((header, index) => [header, values[index]]))),
}
