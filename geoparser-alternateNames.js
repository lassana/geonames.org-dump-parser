var filename = 'alternateNames.txt'
var outputFilename = filename + '-' + new Date().toISOString() + '.xlsx'

var Excel = require('exceljs')

var options = { filename: outputFilename }
var workbook = new Excel.stream.xlsx.WorkbookWriter(options)

workbook.creator = 'Me'
workbook.lastModifiedBy = 'Me'
workbook.created = new Date()
workbook.modified = new Date()

var worksheet = workbook.addWorksheet(filename)

worksheet.columns = [
    { width: 16, header: 'alternateNameId', key: 'alternateNameId' },
    { width: 16, header: 'geonameid',       key: 'geonameid' },
    { width: 12, header: 'isolanguage',     key: 'isolanguage' },
    { width: 25, header: 'alternate name',  key: 'alternate name' },
    { width: 16, header: 'isPreferredName', key: 'isPreferredName' },
    { width: 14, header: 'isShortName',     key: 'isShortName' },
    { width: 14, header: 'isColloquial',    key: 'isColloquial' },
    { width: 14, header: 'isHistoric',      key: 'isHistoric' },
]
worksheet.addRow([
    'the id of this alternate name, int',
    'geonameId referring to id in table \'geoname\', int',
    'iso 639 language code 2- or 3-characters; 4-characters \'post\' for postal codes and \'iata\',\'icao\' and faac for airport codes, fr_1793 for French Revolution names,  abbr for abbreviation, link for a website, varchar(7)',
    'alternate name or name variant, varchar(200)',
    '\'1\', if this alternate name is an official/preferred name',
    '\'1\', if this is a short name like \'California\' for \'State of California\'',
    '\'1\', if this alternate name is a colloquial or slang term',
    '\'1\', if this alternate name is historic and was used in the past'
]).commit()

worksheet.addRow().commit()

console.log('Let\'s parse \"' + filename + '\"...')

var count = 0
var startTime = Date.now()

var fs = require('fs')
require('readline')
    .createInterface( { input: fs.createReadStream(filename)})
    .on('close', function() {
        workbook.commit()
        console.log('Done! Successfully writen to ' + outputFilename)
    })
    .on('line', function (line) {
        var items = line.split('\t')
        
        if ( items.length != 8 ) {
            console.log('Invalid line in' + filename + '!')
            console.log(line)
            process.exit(1)
        }

        if ( items[2].length == 2 || items[2].length == 3 ) { // item is ISO lang code
            worksheet.addRow(items).commit()
            if ( ++count % 100000 == 0 ) {
                console.log('Processed ' + count + ' rows in ' + (Date.now()-startTime)/1000 + ' seconds')
            }
        }
    })

Array.prototype.contains = function(element){
    return this.indexOf(element) > -1
}