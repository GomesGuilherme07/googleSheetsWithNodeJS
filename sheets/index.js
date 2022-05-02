const {GoogleSpreadsheet} = require ('google-spreadsheet');
const credentials = require ("../credentials/credentials.json");

process.env['NODE_TLS_REJECT_UNAUTHORIZED'] = 0;

async function acessarGoogleSheet(googleSheetId){
    
    try{
        //Iniciando a planilha --> Passando o ID do documento como parâmetro
        const documento = new GoogleSpreadsheet (googleSheetId);

        // Autenticação
        await documento.useServiceAccountAuth(credentials);

        // Carregando propriedades do documento
        await documento.loadInfo(); 

        console.log('Acesso ao Google Sheet realizado');

        return documento;

    }catch(error){        
        console.log('error', error)
    }

}

async function buscandoDadosGoogleSheet(googleSheetId){

    try{

    //Instanciando objeto referente a planilha
    documento = await acessarGoogleSheet(googleSheetId);    

    //Selecionando a aba da planilha
    const sheet = documento.sheetsByIndex[0];    

    //Buscando os dados da aba
    const registros = await sheet.getRows();
    console.log('Registros da planilha carregados');

    //Carregando dados
    await sheet.loadCells('A1:C11'); 

    //Acessando celula com base no index
    // const nome = sheet.getCell(1, 0); 
    // const sobrenome = sheet.getCell(1, 1);
    // const email = sheet.getCell(1, 2);

    let dados = [];

    for(r in registros){
        dados.push(
            {
                "Nome": registros[r]._rawData[0],
                "Sobrenome": registros[r]._rawData[1],
                "Email": registros[r]._rawData[2]
            }
        )
    }

    return dados;

    }catch(e){
        console.log("Error - ", e)
    }

}

async function comandosGoogleSheet(){

    try{

        //Instanciando objeto referente a planilha
        documento = await acessarGoogleSheet();

        //Selecionando a aba da planilha
        const sheet = documento.sheetsByIndex[0];

        //Buscando os dados da aba
        const registros = await sheet.getRows();

        //Atualizando titulo
        await documento.updateProperties({title: "Titulo renomeado"});    
        
        // Criando uma nova aba
        const newSheet = await documento.addSheet({ title: 'Nova aba' }); 

        // Deletando aba
        await newSheet.delete();

        //Adicionando novos dados em uma linha
        const newRow = await sheet.addRow({ Nome: 'Larry Page', Email: 'larry@google.com' });

        // Deletando uma linha
        await newRow[1].delete(); 

        //Atualizando valor de uma celula
        registros[1].Email = 'sergey@abc.xyz';
        await registros[1].save();

        //Exibindo dados de um celula
        console.log(registros[1].Email);

        //Carregando os dados de um range
        await sheet.loadCells('A1:B11'); 
        console.log(sheet.cellStats);
        
        //Acessando celula com base no index
        const a1 = sheet.getCell(0, 1); 
        console.log(a1.value);
        console.log(a1.formula);
        console.log(a1.formattedValue);

        //Alterando formatação do texto
        a1.textFormat = {bold: true};

        //Adicionando uma nota dentro da celula
        a1.note = "Adicionando uma nota"

        //Salvando alterações
        await sheet.saveUpdatedCells();       

        console.log(registros);

    } catch(error){
        console.log('error', error)
    }   

}

module.exports = {
    buscandoDadosGoogleSheet: buscandoDadosGoogleSheet
}


