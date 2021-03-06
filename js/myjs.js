let myHeaders = []; //this array will contain all the headers from the table. gets the data from the function BindTableHeader()
let allValues = [];
function removeSymbols(theString){
    let output = theString.replace(
        /^(?:')(.*)(?:')$/,
        "$1"
      );
      return output;
};

function extractValues(theData, theString){
    let values = [];
    for(i=1; i<theData.length;i++){
        let element = theData[i];
        let valueName = element[theString];
        values.push(valueName);
    }

    console.log(values +" array with the key " + theString);
    return values;
};

function createChart(theData){
    
const ctx = document.getElementById('myChart').getContext('2d');

/*this is my code*/
//let subjects = [...theData[0]];

//console.log(subjects);

/**
 * myArray is an object
 *
 *all =  is the key name that extract from theData.
 *key = is the position and becomes a key value in the new object.
 *example of the object {0: '3', 1: '22', 2: '20', 3: '0', 4: '42', 5: '89'}
 *myArray[0] = 3
 */
let myArray = Object.fromEntries(Object
    .entries(theData)
    .map(([key, { all }]) => [key, all])
);

for(let i=1 ; i < myHeaders.length ; i++){
    let value = myHeaders[i];
    let values = [];
    for(let i=0; i<theData.length;i++){
        let element = theData[i];
        let valueName = element[value];
        values.push(valueName);
       
    }
    console.log(values +" array with the key " + value);
    //myChart.data.labels.push[value];
    //console.log(myChart.data.labels);
    allValues.push(values);
};
console.log(allValues)

drawChart(myHeaders[0], allValues[0]);
//This is not working
/*
 for(let i=0 ; i < myHeaders.length ; i++){
    let myN = myHeaders[i];
    let cleanValue = removeSymbols(myN);
    let myArray = Object.fromEntries(Object
        .entries(theData)
        .map((key, { cleanValue }) => [key, cleanValue])
    );

    console.log(myArray);
    console.log(myN);
 };



//console.log(myHeaders);
*/

/*----------------*/


/**
 * This function draw the chart with the values send in the object values and use the headers from the array headers
 *
 *
 */
function drawChart(headers, values){

const myChart = new Chart(ctx, {
    type: 'bar',
    data: {
        labels: myHeaders,
        datasets: [{
            label: headers,
            data: values,
            backgroundColor: [
                'rgba(255, 99, 132, 0.2)',
                'rgba(54, 162, 235, 0.2)',
                'rgba(255, 206, 86, 0.2)',
                'rgba(75, 192, 192, 0.2)',
                'rgba(153, 102, 255, 0.2)',
                'rgba(255, 159, 64, 0.2)'
            ],
            borderColor: [
                'rgba(255, 99, 132, 1)',
                'rgba(54, 162, 235, 1)',
                'rgba(255, 206, 86, 1)',
                'rgba(75, 192, 192, 1)',
                'rgba(153, 102, 255, 1)',
                'rgba(255, 159, 64, 1)'
            ],
            borderWidth: 1
        }]
    },
    options: {
        scales: {
            y: {
                beginAtZero: true
            }
        }
    }
});
};


};


/*this code is from Midhun T P (developer) shared in the page https://www.c-sharpcorner.com/article/reading-a-excel-file-using-html5-jquery/ */

function ExportToTable() {  
    var regex = /^([a-zA-Z0-9\s_\\.\-:])+(.xlsx|.xls)$/;  
    /*Checks whether the file is a valid excel file*/  
    if (regex.test($("#excelfile").val().toLowerCase())) {  
        var xlsxflag = false; /*Flag for checking whether excel is .xls format or .xlsx format*/  
        if ($("#excelfile").val().toLowerCase().indexOf(".xlsx") > 0) {  
            xlsxflag = true;  
        }  
        /*Checks whether the browser supports HTML5*/  
        if (typeof (FileReader) != "undefined") {  
            var reader = new FileReader();  
            reader.onload = function (e) {  
                var data = e.target.result;  
                /*Converts the excel data in to object*/  
                if (xlsxflag) {  
                    var workbook = XLSX.read(data, { type: 'binary' });  
                }  
                else {  
                    var workbook = XLS.read(data, { type: 'binary' });  
                }  
                /*Gets all the sheetnames of excel in to a variable*/  
                var sheet_name_list = workbook.SheetNames;  
 
                var cnt = 0; /*This is used for restricting the script to consider only first sheet of excel*/  
                sheet_name_list.forEach(function (y) { /*Iterate through all sheets*/  
                    /*Convert the cell value to Json*/  
                    if (xlsxflag) {  
                        var exceljson = XLSX.utils.sheet_to_json(workbook.Sheets[y]);  
                    }  
                    else {  
                        var exceljson = XLS.utils.sheet_to_row_object_array(workbook.Sheets[y]);  
                    }  
                    if (exceljson.length > 0 && cnt == 0) {  
                        BindTable(exceljson, '#exceltable');  
                        cnt++;  
                        createChart(exceljson);
                    }  
                });  
                $('#exceltable').show();  
            }  
            if (xlsxflag) {/*If excel file is .xlsx extension than creates a Array Buffer from excel*/  
                reader.readAsArrayBuffer($("#excelfile")[0].files[0]);  
            }  
            else {  
                reader.readAsBinaryString($("#excelfile")[0].files[0]);  
            }  
        }  
        else {  
            alert("Sorry! Your browser does not support HTML5!");  
        }  
    }  
    else {  
        alert("Please upload a valid Excel file!");  
    }  
}  


function BindTable(jsondata, tableid) {/*Function used to convert the JSON array to Html Table*/  
    var columns = BindTableHeader(jsondata, tableid); /*Gets all the column headings of Excel*/  
    for (var i = 0; i < jsondata.length; i++) {  
        var row$ = $('<tr/>');  
        for (var colIndex = 0; colIndex < columns.length; colIndex++) {  
            var cellValue = jsondata[i][columns[colIndex]];  
            if (cellValue == null)  
                cellValue = "";  
            row$.append($('<td/>').html(cellValue));

        }  
        $(tableid).append(row$);  
    }  
}  
function BindTableHeader(jsondata, tableid) {/*Function used to get all column names from JSON and bind the html table header*/  
    var columnSet = [];  
    var headerTr$ = $('<tr/>');  
    for (var i = 0; i < jsondata.length; i++) {  
        var rowHash = jsondata[i];  
        for (var key in rowHash) {  
            if (rowHash.hasOwnProperty(key)) {  
                if ($.inArray(key, columnSet) == -1) {/*Adding each unique column names to a variable array*/  
                    columnSet.push(key);  
                    headerTr$.append($('<th/>').html(key));  
                    myHeaders.push(key);//this is my code it push the headers into one array named myHeaders
                }  
            }  
        }  
    }  
    $(tableid).append(headerTr$);  
    return columnSet;  
}  


/* this is how it looks the json: 1 array with 1 object for each row and keys for every column
0: {time: '10', all: '3', more: '4', tall: '5', call: '55', ???}
1: {time: '12', all: '22', more: '85', tall: '85', call: '24', ???}
2: {time: '14', all: '20', more: '20', tall: '365', call: '45', ???}
3: {time: '16', all: '0', more: '42', tall: '2220', call: '42', ???}
4: {time: '18', all: '42', more: '43', tall: '436', call: '3456', ???}
5: {time: '20', all: '89', more: '789', tall: '97', call: '8', ???}
*/