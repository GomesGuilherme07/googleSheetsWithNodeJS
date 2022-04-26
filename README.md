# Utilizando NodeJS para acessar planilha (Google Sheets)

### Configurando o projeto

1. Com a conta do Google logado, criar um novo projeto no [Google Cloud]( https://console.cloud.google.com/apis/dashboard) e logo após ativar clicar em "ATIVAR APIS E SERVIÇOS". Devemos selecionar o serviço Google Sheets API e Google Drive API e ativá-los.

2. Selecione o botão de “Criar credenciais” e siga pela opção "conta de serviço". Após criar, acesse a credencial e procure o campo "Chaves", nele clique em "Adicionar chave", escolha o formato JSON e baixe o arquivo. Esse arquivo possui as credenciais que serão utilizadas para acessar a planilha.

3. Inicie um novo projeto NodeJS e instale as dependecias:
```sh
npm install google-spreadsheet@3.2.0
```
```sh
npm install nodemon@2.0.15
```
4. Cole o arquivo JSON com as credenciais no mesmo diretório do projeto.

5. Agora será possível utilizar as credenciais e acessar a planilha

### Acessando a planilha

1. No arquivo index.js faça os seguintes imports:
~~~javascript
const {GoogleSpreadsheet} = require ('google-spreadsheet');
const credentials = require ("./credentials.json");
~~~
2. Crie uma constante que irá receber o ID da planilha, esse ID é encontrado na URL do arquivo
~~~javascript
const googleSheetId = '<id da planilha>';
~~~
3. Crie uma função com o código abaixo:
~~~javascript
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
~~~
4. O retorno dessa função será o próprio documento do Google Sheets, através da variável **documento** é possível acessar os dados da planilha

### Manipulando dados 

1. Alguns comandos para manipular a planilha 
~~~javascript
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
        const newRow = await sheet.addRow({ Nome: 'teste', Email: 'teste@google.com' });

        // Deletando uma linha
        await newRow[1].delete(); 

        //Atualizando valor de uma celula
        registros[1].Email = 'teste@abc.xyz'; // update a value
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
~~~
