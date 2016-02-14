var wsCitiesName = "Cities"
var XLSX = require('xlsx')
var wb = {}
wb.Sheets = {}
wb.Props = {}
wb.SSF = {}
wb.SheetNames = []
var wsCities = {}
var range = {s: {c:0, r:0}, e: {c:0, r:0 }}

var header = [
    [ 10, 'geonameid',          'integer id of record in geonames database' ],
    [ 20, 'name',               'name of geographical point (utf8) varchar(200)' ],
    [ 20, 'asciiname',          'name of geographical point in plain ascii characters, varchar(200)' ],
    [ 30, 'alternatenames',     'alternatenames, comma separated, ascii names automatically transliterated, convenience attribute from alternatename table, varchar(10000)' ],
    [ 10, 'latitude',           'latitude in decimal degrees (wgs84)' ],
    [ 10, 'longitude',          'longitude in decimal degrees (wgs84)' ],
    [ 5,  'feature class',      'see http://www.geonames.org/export/codes.html, char(1)' ],
    [ 5,  'feature code',       'see http://www.geonames.org/export/codes.html, varchar(10)' ],
    [ 5,  'country code',       'ISO-3166 2-letter country code, 2 characters' ],
    [ 5,  'cc2',                'alternate country codes, comma separated, ISO-3166 2-letter country code, 200 characters' ],
    [ 5,  'admin1 code',        'fipscode (subject to change to iso code), see exceptions below, see file admin1Codes.txt for display names of this code; varchar(20)' ],
    [ 5,  'admin2 code',        'code for the second administrative division, a county in the US, see file admin2Codes.txt; varchar(80) ' ],
    [ 5,  'admin3 code',        'code for third level administrative division, varchar(20)' ],
    [ 5,  'admin4 code',        'code for fourth level administrative division, varchar(20)' ],
    [ 10, 'population',         'bigint (8 byte int) ' ],
    [ 5,  'elevation',          'in meters, integer' ],
    [ 5,  'dem',                'digital elevation model, srtm3 or gtopo30, average elevation of 3\'\'x3\'\' (ca 90mx90m) or 30\'\'x30\'\' (ca 900mx900m) area in meters, integer. srtm processed by cgiar/ciat.' ],
    [ 10, 'timezone',           'the timezone id (see file timeZone.txt) varchar(40)' ],
    [ 10, 'modification date',  'date of last modification in yyyy-MM-dd format' ]
]

for (var i=0, len = header.length; i<len; i++) {
    wsCities[XLSX.utils.encode_cell({c:i,r:0})] = { v: header[i][1], t: 's' }
    wsCities[XLSX.utils.encode_cell({c:i,r:1})] = { v: header[i][2], t: 's' }
}
range.e.c = header.length
range.e.r = header[0].length

var fileName = process.argv[2]
console.log('Let\' process \"' + fileName + '\"...')
var fs = require('fs')
require('readline')
    .createInterface( { input: fs.createReadStream(fileName)})
    .on('line', function (line) {
        var items = line.split('\t')
        if ( items.length != range.e.c ) {
            console.log('Invalid line!')
            console.log(line)
            process.exit(1)
        }
        for (var i=0, len=items.length; i<len; i++) {
            var cellRef = XLSX.utils.encode_cell({c:i,r:range.e.r})
            wsCities[cellRef] = { v: items[i], t: 's' }
        }
        range.e.r += 1
    })
    .on('close', function () {
        console.log('Parsing finished, now writing XLSX output file...')
        wb.SheetNames.push(wsCitiesName)
        wb.Sheets[wsCitiesName] = wsCities
        wsCities['!ref'] = XLSX.utils.encode_range(range)
        wsCities['!cols'] = []
        for (var i=0, len=header.length; i<len; i++) {
            wsCities['!cols'][i] = {wch:header[i][0]}
        }
        var outputFilename = fileName + '-' + new Date().toISOString() +'.xlsx'
        XLSX.writeFile(wb, outputFilename)
        console.log('Successfully writen to ' + outputFilename)
    })

