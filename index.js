const {GoogleSpreadsheet} = require ('google-spreadsheet');
const credentials = require ("./credentials.json");

//ID da planila --> Localizado na URL do documento
const googleSheetId = '1gSx6ASI6dxspN6fMzVM7puY3zaBhCjmcccDAyM3AlxM';


async function acessarGoogleSheet(){
    
    try{
        //Iniciando a planilha --> Passando o ID do documento como parâmetro
        const documento = new GoogleSpreadsheet (googleSheetId);

        // Autenticação
        await documento.useServiceAccountAuth(credentials);

        // Carregando propriedades do documento
        await documento.loadInfo(); 

        return documento;

    }catch(error){        
        console.log('error', error)
    }

}

async function manipulandoDadosGoogleSheet(){

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
        registros[1].Email = 'sergey@abc.xyz'; // update a value
        await registros[1].save(); // save updates

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

    } catch(error){
        console.log('error', error)
    }   

}

manipulandoDadosGoogleSheet();

