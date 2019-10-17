const xlsx = require('node-xlsx')
const dbf = require('dbf')
const readline = require('readline')
const fs = require('fs')

const now = new Date
const monName = new Array ('01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12')

function item(codbar, qtdemb1, embala1 ) {
    this.codbar = codbar
    this.qtdemb1 = qtdemb1
    this.embala1 = embala1
}

const leitor = readline.createInterface({
    input: process.stdin,
    output: process.stdout
})

const config = fs.readFileSync(__dirname + '/config.json', 'utf-8')


const plan = xlsx.parse(JSON.parse(config).path.input + '/dados.xls')

const lista = plan[0].data.map( function(e, i) {
    if (i !== 0) {
        return new item(e[1], 1, e[6] + e[7])
    }
}).filter(e => e !== undefined)

//console.log(lista)

let buf = dbf.structure(lista)

leitor.question("Para qual unidade é: ", (unidade) => {
    fs.writeFileSync(JSON.parse(config).path.output + 
        '/tp' +unidade + '01'+ now.getFullYear() + monName[now.getMonth() ] + 
        now.getDate() + '00.dbf', toBuffer(buf.buffer))
    leitor.close()
})

fs.rename(JSON.parse(config).path.input + '/dados.xls' , JSON.parse(config).path.input + '/dados_old.xls', (e) => { })

function toBuffer(ab) {
    let buffer = new Buffer.alloc(ab.byteLength)
    let view = new Uint8Array(ab)
    for (var i = 0; i < buffer.length; ++i) {
        buffer[i] = view[i]
    }
    return buffer
}