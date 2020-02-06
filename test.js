const convertCsvToXlsx = require('@aternus/csv-to-xlsx');
const moment = require('moment');

var hrstart = process.hrtime()

convertCsvToXlsx('./latest-Data.csv', './' + moment().format('DD-MM-YYYY-h-mmA') + '.xlsx');

setTimeout(() => {
    hrend = process.hrtime(hrstart)

    console.info('Execution time (hr): %dm %dms', hrend[0] / 60)
}, 60000)


console.log(moment().format('DD-MM-YYYY-h:mmA'));