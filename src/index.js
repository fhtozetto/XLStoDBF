const xlsx = require('node-xlsx')
const dbf = require('dbf')
const readline = require('readline')
const fs = require('fs')

const now = new Date
const dateXX = new Array (null, '01', '02', '03', '04', '05', '06', '07', '08', '09', '10', 
                            '11', '12', '13', '14', '15', '16', '17', '18', '19', '20', 
                            '21', '22', '23', '24', '25', '26', '27', '28', '29', '30', '31')
const monXX = new Array ('01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12')

function item(codbar, qtdemb1, embala1 ) {
    this.codbar = codbar
    this.qtdemb1 = qtdemb1
    this.embala1 = embala1
}

const leitor = readline.createInterface({
    input: process.stdin,
    output: process.stdout
})

const config = fs.readFileSync(__dirname +'/config.json', 'utf-8')

try {
    const plan = xlsx.parse(JSON.parse(config).path.input +'/dados.xls')

    const lista = plan[0].data.map( function(e, i) {
        if (i !== 0) {
            return new item(e[1], e[6] + e[7], 1)
        }
    }).filter(e => e !== undefined)

    let buf = dbf.structure(lista)
    try {
        leitor.question("Para qual unidade é: ", (unidade) => {
            fs.readdir(JSON.parse(config).path.backup +'/'+ unidade, (err) => {
                if (err) {
                    console.log('pasta dessa Unidade não encontrada: '+ err.path)
                } else {
                    fs.writeFileSync(JSON.parse(config).path.output 
                        +'/tp'+ unidade +'099'
                        + now.getFullYear() 
                        + monXX[now.getMonth()] 
                        + dateXX[now.getDate()]
                        +'00.dbf', 
                        toBuffer(buf.buffer))

                    const readableStream = fs.createReadStream(JSON.parse(config).path.input +'/dados.xls') 
                    
                    let writableStream = fs.createWriteStream(JSON.parse(config).path.backup 
                        +'/'+ unidade 
                        +'/'+ unidade 
                        +'_'+ now.getFullYear() 
                        +'-'+ monXX[now.getMonth()] 
                        +'-'+ dateXX[now.getDate()]
                        +'-'+ now.getHours() +'h'+ now.getMinutes() +'m'+ now.getSeconds() +'s.xls')
        
                    readableStream.pipe(writableStream)
                    fs.rename(JSON.parse(config).path.input +'/dados.xls' , JSON.parse(config).path.input +'/dados_old.xls', (e) => { })

                    leitor.close()
                }
            })
            
        })
        
    } catch(e) {
        console.log(`Caminho "${JSON.parse(config).path.output}" inacessivel`)
    }

} catch(e) {
    console.log(`Arquivo "Dados.xls" NAO encontrado em "${JSON.parse(config).path.input}"`)
}

function toBuffer(ab) {
    let buffer = new Buffer.alloc(ab.byteLength)
    let view = new Uint8Array(ab)
    for (let i = 0; i < buffer.length; ++i) {
        buffer[i] = view[i]
    }
    return buffer
}