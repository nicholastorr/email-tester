const reader = require('xlsx')

const file = reader.readFile('./test.xlsx')

let data = []


//read sheets and console log condionally
/*const sheets = file.SheetNames

for (let i = 0; i < sheets.length; i++) {
    const temp = reader.utils.sheet_to_json(file.Sheets[file.SheetNames[i]])
        temp.forEach((res) => {
            if (res.branch === 'CSE') {
                console.log(res)
            }
        })
}*/


//data to append new sheet
let studentData = [{
        student: 'Nikhil',
        age: '33',
        branch: 'ISE',
        marks: 70
    },
    {
        student: 'Amitha',
        age: '24',
        branch: 'EC',
        marks: 80
    },
]

//change student data into json format of xcel sheet
const ws = reader.utils.json_to_sheet(studentData)

//append new sheet to xcel file
reader.utils.book_append_sheet(file,ws,"Sheet3")
  
// Writing to our file
reader.writeFile(file,'./test.xlsx')


//append new data to Sheet 1 of excel file
var wb = file.Sheets["Sheet1"]

reader.utils.sheet_add_aoa(wb, [
    ['Nikhil', 33, 'ISE', 70],
    ['Amitha', 24, 'EC', 80]
  ], {origin: -1})

//


const sheets = file.SheetNames
//loop through sheet and make sure everything works
for (let i = 0; i < sheets.length; i++) {
    const temp = reader.utils.sheet_to_json(file.Sheets[file.SheetNames[i]])
        temp.forEach((res) => {
                data.push(res)
        })
}

console.log(data)