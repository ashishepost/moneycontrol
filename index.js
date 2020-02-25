// Libraries Imports
const axios = require("axios");
const cheerio = require("cheerio");
const fs = require('fs');
var parse = require('html-dom-parser');
var DomParser = require('dom-parser');
var parser = new DomParser();
const voca = require('voca');
var requestSync = require('sync-request');
const createCsvWriter = require('csv-writer').createObjectCsvWriter;
const convertCsvToXlsx = require('@aternus/csv-to-xlsx');
const moment = require('moment');

// Global Usefull Variable Definations
const messageBlock = '=============';
const baseURL = "https://www.moneycontrol.com/";

const rootURL = baseURL + "stocks/marketinfo/marketcap/bse/index.html";
var records = [];
var CSVFilePath = './latest-Data.csv';
var hrstart = process.hrtime();



const fetchData = async() => {
    terminalPrinter("Start Fetching Root Page");
    const result = await axios.get(rootURL);
    terminalPrinter("Stop Fetching Root Page");
    return cheerio.load(result.data);
};
const parseHTML = async(selector) => {
    const cheerio$ = await fetchData();
    const selectorData = cheerio$(selector).html();
    // console.log(selectorData);
    return selectorData;
}

// Main Code Start Here

// const postJobButton = $('.top > .action-post-job').text();
// console.log(postJobButton) // Logs 'Post a Job'

// const cheerio$ = await fetchData();
// let parsedData = parseHTML('.tbldata14 bdrtpg');


(async() => {
    let parsedData = await parseHTML('.tbldata14');
    // terminalPrinter("Got the HTML Table");
    let DOMObject = parseHTMLTableToJSON('<table>' + parsedData + '</table>');
    // console.log(DOMObject);
    let allRows = DOMObject.getElementsByTagName("tr");
    // console.log(allRows[0].innerHTML);

    terminalPrinter('Starting Scraping');
    // allRows.length
    for (let index = 1; index < allRows.length; index++) {
        // for (let index = 1; index < 2; index++) {
        // console.log(allRows[index].outerHTML);
        const tempDOM = parseHTMLTableToJSON(allRows[index].outerHTML); // All Rows
        let allColums = tempDOM.getElementsByTagName("td");
        let companyURL = tempDOM.getElementsByTagName("a");
        // Any particular column for any Single Row & URL from Attributues
        // console.log(baseURL + companyURL[0].attributes[0].value); 
        splittedString = voca.split(companyURL[0].attributes[0].value, '/');
        ratioURL = baseURL + 'financials/' + splittedString[splittedString.length - 2] + '/ratios/' + splittedString[splittedString.length - 1];

        let ratioURLHTML = URLRequester(ratioURL);
        let ratioData = getInterestCover(ratioURLHTML)
        // records.push({ company: allColums[0].firstChild.textContent, lastprice: allColums[1].firstChild.textContent, interestcover: ratioData, ratioURL: ratioURL }, );
        
        let netProfitMarginData = getNetProfitMargin(ratioURLHTML);
      
        // console.log("hello");
        
        // console.log(netProfitMarginData);
        // let allYearsSeparatedData = netProfitMarginData.split(",");
        if (netProfitMarginData === undefined) {
            netProfitMarginData = ",";
            
        }
        if ( netProfitMarginData.includes(".")) {
            let allYearsSeparatedData = netProfitMarginData.split("::");
            // console.log(netProfitMarginData);
            // console.log(allYearsSeparatedData[2]);
            // process.exit(0);
            
            records.push({ company: allColums[0].firstChild.textContent, lastprice: allColums[1].firstChild.textContent, ratio: ratioData, 2019: allYearsSeparatedData[0], 2018: allYearsSeparatedData[1], 2017: allYearsSeparatedData[2], 2016: allYearsSeparatedData[2], 2015: allYearsSeparatedData[4], ratioURL: ratioURL });
            // console.log({ company: allColums[0].firstChild.textContent, lastprice: allColums[1].firstChild.textContent, ratio: ratioData, '2019': allYearsSeparatedData[0]), '2018': allYearsSeparatedData[1]), '2017': allYearsSeparatedData[2]), '2016': allYearsSeparatedData[3]), '2015': allYearsSeparatedData[4]), ratioURL: ratioURL });
            
          
        }else{
            records.push({ company: allColums[0].firstChild.textContent, lastprice: allColums[1].firstChild.textContent, ratio: ratioData, ratioURL: ratioURL } );

        }
       
        
        // For One Record While Checking
        // break;
        console.log("record " + index);
        // console.log(allColums[0].firstChild.textContent);
        // console.log(allColums[1].firstChild.textContent);
        // console.log(ratioData);

    }
    terminalPrinter('Stoping Scraping')

    // Write CSV File
    writeCSV(CSVFilePath);


})();




// All Synchronous Definations
function terminalPrinter(message) {
    console.log(messageBlock + message + messageBlock);
}

function parseHTMLTableToJSON(HTMLSource) {
    // console.log(HTMLSource);
    const DOMObject = parser.parseFromString(HTMLSource);
    return DOMObject;
}

function stringFormatter(string) {
    string = string.replace(/\\n/g, '');
    string = string.replace(/\\t/g, '');
    return string;
}

function URLRequester(URL) {
    let res = requestSync('GET', URL);
    fs.writeFileSync('output/last.html', res.getBody().toString());
    return res.getBody().toString();
}



function writeFile(filePath, data) {
    fs.writeFile(filePath, data, (err) => {
        // throws an error, you could also catch it here
        if (err) throw err;

        // success case, the file was saved
        console.log('File Saved');
    });
}

function getInterestCover(data) {
    const tempDOM = parseHTMLTableToJSON(data); // All Rows
    let allColums = tempDOM.getElementsByTagName("td");
    for (let index = 0; index < allColums.length; index++) {
        const element = allColums[index];
        // console.log(element.innerHTML);

        if (element.innerHTML == "Interest Cover") {
            if (allColums[index + 1].innerHTML) {
                return allColums[index + 1].innerHTML
            } else {
                return ",";
            };
        }
    }
    // return ',';
}

function getNetProfitMargin(data) {
    const tempDOM = parseHTMLTableToJSON(data); // All Rows
    
    // fs.writeFileSync("/usr/local/google/home/tiwashish/projects/personal/projects/moneycontrol/scrapper/source-code/temp/allrows.txt", JSON.stringify(tempDOM));
    // process.exit(0);

    let allColums = tempDOM.getElementsByTagName("td");
    for (let index = 0; index < allColums.length; index++) {
        const element = allColums[index];
        // console.log(element.innerHTML);
        // fs.writeFileSync("/usr/local/google/home/tiwashish/projects/personal/projects/moneycontrol/scrapper/source-code/temp/allrows.txt", JSON.stringify(tempDOM));
        // process.exit(0);

        if (element.innerHTML == "Net Profit Margin(%)" || element.innerHTML == "Net Profit Margin") {
            if (allColums[index + 1].innerHTML) {
                // console.log(allColums[index + 1].innerHTML + '::' + allColums[index + 2].innerHTML + '::' + allColums[index + 3].innerHTML + '::' + allColums[index + 4].innerHTML + '::' + allColums[index + 5].innerHTML);
                
                return allColums[index + 1].innerHTML + '::' + allColums[index + 2].innerHTML + '::' + allColums[index + 3].innerHTML + '::' + allColums[index + 4].innerHTML + '::' + allColums[index + 5].innerHTML;
            } else {
                return ",";
            };
        }
    }
}





function writeCSV(destination) {
    const csvWriter = createCsvWriter({
        path: destination,
        header: [
            { id: 'company', title: 'Company' },
            { id: 'lastprice', title: 'Price' },
            { id: 'ratio', title: 'Ratio' },
            { id: 2019, title: 'Year 2019' },
            { id: 2018, title: 'Year 2018' },
            { id: 2017, title: 'Year 2017' },
            { id: 2016, title: 'Year 2016' },
            { id: 2015, title: 'Year 2015' },
            { id: 'ratioURL', title: 'Ratio URL' }
        ]
    });
    // console.log(records);

    csvWriter.writeRecords(records) // returns a promise
        .then(() => {

            // Write Excel File
            writeXLSX(CSVFilePath, './' + moment().format('DD-MM-YYYY-h-mmA') + '.xlsx');
            console.log('XLSX File is Written');

            // Delete CSV File
            fs.unlinkSync(CSVFilePath);

            // Time Section
            hrend = process.hrtime(hrstart)
            console.info('Execution time (Minutes:Seconds): %dm %ds', hrend[0] / 60, hrend[0])
        });
}

function writeXLSX(source, destination) {
    // console.log(source, destination);
    convertCsvToXlsx(source, destination);

}
