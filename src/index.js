const xlsx = require('node-xlsx')
const dbf = require('dbf')
const fs = require('fs')

function item(codbar, qtdemb1, embala1 ) {
    this.codbar = codbar
    this.qtdemb1 = qtdemb1
    this.embala1 = embala1
}

const config = fs.readFileSync(__dirname + '/config.json', 'utf-8')

const plan = xlsx.parse(JSON.parse(config).path.input + '/dados.xls')

const lista = plan[0].data.map( function(e, i) {
    if (i !== 0) {
        return new item(e[1], 1, e[6] + e[7])
    }
}).filter(e => e !== undefined)

//console.log(lista)

let buf = dbf.structure(lista)

fs.writeFileSync(JSON.parse(config).path.output + '/dados.dbf', toBuffer(buf.buffer))

function toBuffer(ab) {
    let buffer = new Buffer(ab.byteLength)
    let view = new Uint8Array(ab)
    for (var i = 0; i < buffer.length; ++i) {
        buffer[i] = view[i]
    }
    return buffer
}



