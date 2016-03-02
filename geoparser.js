var Excel = require('exceljs')

var outputFilename = 'geonames.org' + '-' + new Date().toISOString() + '.xlsx'

var geonamesIds = {}

var options = { filename: outputFilename }
var workbook = new Excel.stream.xlsx.WorkbookWriter(options)
workbook.creator = 'Me'
workbook.lastModifiedBy = 'Me'
workbook.created = new Date()
workbook.modified = new Date()

parseCities(function() {
    parseAlternateNames(function() {
        workbook.commit()
        console.log('Done! Successfully writen to ' + outputFilename)
    })
})



function parseCities(callback) {
    var filename = 'cities15000.txt'

    var worksheetCities = workbook.addWorksheet(filename)
    var worksheetAlternateNames = workbook.addWorksheet(filename + '-alternatenames')

    addHeader(worksheetCities, [
        /*  0 */ [ 10, 'geonameid',         'integer id of record in geonames database' ],
        /*  1 */ [ 20, 'name',              'name of geographical point (utf8) varchar(200)' ],
        /*  2 */ [ 20, 'asciiname',         'name of geographical point in plain ascii characters, varchar(200)' ],
        /*  3    [ 30, 'alternatenames',    'alternatenames, comma separated, ascii names automatically transliterated, convenience attribute from alternatename table, varchar(10000)' ],
        /*  4 */ [ 10, 'latitude',          'latitude in decimal degrees (wgs84)' ],
        /*  5 */ [ 10, 'longitude',         'longitude in decimal degrees (wgs84)' ],
        /*  6 */ [ 5,  'feature class',     'see http://www.geonames.org/export/codes.html, char(1)' ],
        /*  7 */ [ 5,  'feature code',      'see http://www.geonames.org/export/codes.html, varchar(10)' ],
        /*  8 */ [ 5,  'country code',      'ISO-3166 2-letter country code, 2 characters' ],
        /*  9 */ [ 5,  'cc2',               'alternate country codes, comma separated, ISO-3166 2-letter country code, 200 characters' ],
        /* 10 */ [ 5,  'admin1 code',       'fipscode (subject to change to iso code), see exceptions below, see file admin1Codes.txt for display names of this code; varchar(20)' ],
        /* 11 */ [ 15, 'admin2 code',       'code for the second administrative division, a county in the US, see file admin2Codes.txt; varchar(80) ' ],
        /* 12 */ [ 15, 'admin3 code',       'code for third level administrative division, varchar(20)' ],
        /* 13 */ [ 15, 'admin4 code',       'code for fourth level administrative division, varchar(20)' ],
        /* 14 */ [ 10, 'population',        'bigint (8 byte int) ' ],
        /* 15 */ [ 5,  'elevation',         'in meters, integer' ],
        /* 16 */ [ 5,  'dem',               'digital elevation model, srtm3 or gtopo30, average elevation of 3\'\'x3\'\' (ca 90mx90m) or 30\'\'x30\'\' (ca 900mx900m) area in meters, integer. srtm processed by cgiar/ciat.' ],
        /* 17 */ [ 20, 'timezone',          'the timezone id (see file timeZone.txt) varchar(40)' ],
        /* 18 */ [ 15, 'modification date', 'date of last modification in yyyy-MM-dd format' ]
    ])

    addHeader(worksheetAlternateNames, [
        /*  0 */ [ 10, 'geonameid',         'integer id of record in geonames database' ],
        /*  1 */ [ 40, 'alternatenames',    'alternatenames, comma separated, ascii names automatically transliterated, convenience attribute from alternatename table, varchar(10000)' ]
    ])
    console.log('Let\'s parse \"' + filename + '\"...')

    var count = 0
    var startTime = Date.now()

    var fs = require('fs')
    require('readline')
        .createInterface( { input: fs.createReadStream(filename)})
        .on('close', function() {
            console.log('Finish parsing ' + filename)
            console.log('geonamesIds.length = ' + geonamesIds.length)
            callback()
        })
        .on('line', function (line) {
            var items = line.split('\t')

            if ( items.length != 19 ) {
                console.log('Invalid line in' + filename + '!')
                console.log(line)
                process.exit(1)
            }

            geonamesIds[items[0]] = null

            var alternates = items[3].split(',')
            for (var i=0, len = alternates.length; i<len; i++) {
                worksheetAlternateNames.addRow([items[0], alternates[i]])
            }

            items.splice(3,1)
            worksheetCities.addRow(items).commit()
            if ( ++count % 100000 == 0 ) {
                console.log('Processed ' + count + ' rows in ' + (Date.now()-startTime)/1000 + ' seconds')
            }
        })
}

function parseAlternateNames(callback) {
    var filename = 'alternateNames.txt'
    var outputFilename = filename + '-' + new Date().toISOString() + '.xlsx'
    var worksheet = workbook.addWorksheet(filename)
    var ignoredLanguages = { 'post' : null, 'iata' : null, 'icao' : null, 'faac' : null, 'fr_1793' : null,  'abbr' : null, 'link' : null }

    addHeader(worksheet, [
        /* 0 */ [ 16, 'alternateNameId', 'the id of this alternate name, int' ],
        /* 1 */ [ 16, 'geonameid',       'geonameId referring to id in table \'geoname\', int' ],
        /* 2 */ [ 12, 'isolanguage',     'iso 639 language code 2- or 3-characters; 4-characters \'post\' for postal codes and \'iata\',\'icao\' and faac for airport codes, fr_1793 for French Revolution names,  abbr for abbreviation, link for a website, varchar(7)' ],
        /* 3 */ [ 25, 'alternate name',  'alternate name or name variant, varchar(200)' ],
        /* 4 */ [ 16, 'isPreferredName', '\'1\', if this alternate name is an official/preferred name' ],
        /* 5 */ [ 14, 'isShortName',     '\'1\', if this is a short name like \'California\' for \'State of California\'' ],
        /* 6 */ [ 14, 'isColloquial',    '\'1\', if this alternate name is a colloquial or slang term' ],
        /* 7 */ [ 14, 'isHistoric',      '\'1\', if this alternate name is historic and was used in the past' ],
    ])
    worksheet.addRow().commit()

    console.log('Let\'s parse \"' + filename + '\"...')

    var count = 0
    var startTime = Date.now()

    var fs = require('fs')
    require('readline')
        .createInterface( { input: fs.createReadStream(filename)})
        .on('close', function() {
            callback()
        })
        .on('line', function (line) {
            var items = line.split('\t')

            if ( items.length != 8 ) {
                console.log('Invalid line in' + filename + '!')
                console.log(line)
                process.exit(1)
            }

            //var isValidLanguageCode = (items[2].length == 2 || items[2].length == 3)
            var isValidLanguageCode = ! (items[2].toString() in ignoredLanguages)

            if ( isValidLanguageCode
                && items[1].toString() in geonamesIds
                /* && geonamesIds.contains(items[1]) */ ) {
                worksheet.addRow(items).commit()
                if ( ++count % 10000 == 0 ) {
                    console.log('Processed ' + count + ' rows in ' + (Date.now()-startTime)/1000 + ' seconds')
                }
            }
        })
}

Array.prototype.contains = function(element){
    return this.indexOf(element) > -1
}

function addHeader(worksheet, meta) {
    var headerRow = []
    var subHeaderRow = []
    for (var i=0, len=meta.length; i<len; i++) {
        headerRow.push({ header: meta[i][1], key: meta[i][1], width: meta[i][0] })
        subHeaderRow.push(meta[i][2])
    }
    worksheet.columns = headerRow
    worksheet.addRow(subHeaderRow).commit()
    worksheet.addRow().commit()
}
