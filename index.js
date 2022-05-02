const sheets = require('./sheets/index.js');

//Id dos arquivos
const SHEET_ID = '1gSx6ASI6dxspN6fMzVM7puY3zaBhCjmcccDAyM3AlxM';
const SLIDE_ID = '1dxTMrDV9rjSp4FMNM6Ic0M3zDB7w3hQV6p9SesketX4';

//Buscar dados planilha
async function buscarSheet(){

    let dados = await sheets.buscandoDadosGoogleSheet(SHEET_ID);   

    console.log("\nInformações extraídas da planilha: ")
    for(d in dados){
        console.log(`\nNome completo: ${dados[d].Nome} ${dados[d].Sobrenome} - Email: ${dados[d].Email}`)
    }

}

//Alterar informações slide
async function updateSlide(){
    
}

buscarSheet();
