# Utilizando NodeJS para acessar planilha Google Sheets e Google Slides

## Projeto que utiliza NodeJS para manipular dados dentro de uma planilha e atualizar informações de uma apresentação .pptx

### Configurando o projeto

1. Com a conta do Google logada, criar um novo projeto no [Google Cloud]( https://console.cloud.google.com/apis/dashboard) e logo após ativar clicar em "ATIVAR APIS E SERVIÇOS". Devemos selecionar os serviços: Google Sheets API, Google Slides API, Google Drive API e ativá-los.

![image](https://user-images.githubusercontent.com/23075005/165371718-0fa0d989-9838-48b3-a2bf-c9f73fabbba7.png)

![image](https://user-images.githubusercontent.com/23075005/165371782-8ae88ed4-ba65-45c2-b2a1-9d5f74dd3148.png)

2. Selecione o botão de “Criar credenciais” e siga pela opção "conta de serviço". Após criar, acesse a credencial e procure o campo "Chaves", nele clique em "Adicionar chave", escolha o formato JSON e baixe o arquivo. Esse arquivo possui as credenciais que serão utilizadas para acessar a planilha.

![image](https://user-images.githubusercontent.com/23075005/165371815-a8998254-5ead-4db4-b1d3-db1f16ed7a1a.png)

![image](https://user-images.githubusercontent.com/23075005/165371854-960395e7-629d-46a9-87a1-60f0561aac67.png)

3. Dentro do arquivo JSON existe um campo chamado "client_email" com um email gerado para acessar a planilha. Copie o endereço de e-mail e compartilhe o acesso da planilha e da pasta onde o arquivo está para esse e-mail.

4. Inicie um novo projeto NodeJS e instale as dependecias:
```sh
npm install google-spreadsheet@3.2.0
```
```sh
npm install google-slides@2.1.1
```
```sh
npm install nodemailer@6.7.5
```
```sh
npm install nodemon@2.0.15
```
5. Cole o arquivo JSON com as credenciais no mesmo diretório do projeto.

6. Agora será possível utilizar as credenciais e acessar a planilha

## Google Sheets

### Acessando a planilha

1. No arquivo index.js faça os imports abaixo (o caminho "./credentials.json" é o local do arquivo JSON com as credenciais, altere para o caminho correto do seu diretório)
~~~javascript
import { GoogleSpreadsheet } from 'google-spreadsheet';

import { createRequire } from "module";
const require = createRequire(import.meta.url);
const credentials = require('../credentials/credentials.json');
~~~

2. Crie uma constante que irá receber o ID da planilha, esse ID é encontrado na URL do arquivo. Substitua <id da planilha> pelo seu ID.
~~~javascript
const googleSheetId = '<id da planilha>';
~~~
3. Crie uma função com o código abaixo, ela será responsável por fazer a autenticação e buscar a planilha.
~~~javascript
async function acessarGoogleSheet(googleSheetId){
    
    try{
        //Iniciando a planilha --> Passando o ID do documento como parâmetro
        const documento = new GoogleSpreadsheet (googleSheetId);

        // Autenticação
        await documento.useServiceAccountAuth(credentials);

        // Carregando propriedades do documento
        await documento.loadInfo(); 

        console.log('\nAcesso ao Google Sheet realizado');

        return documento;

    }catch(error){        
        console.log('error', error)
    }

}
~~~
4. O retorno dessa função será o próprio documento do Google Sheets, através da variável **documento** é possível acessar os dados da planilha

### Manipulando dados 

1. Alguns comandos para manipular a planilha:
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

## Google Slides

### Acessando a apresentação

1. No arquivo index.js faça os imports abaixo (o caminho "./credentials.json" é o local do arquivo JSON com as credenciais, altere para o caminho correto do seu diretório)

~~~javascript
import { API, TextReplacement, ShapeReplacementWithImage, Presentation } from 'google-slides'
~~~

2. Crie uma constante 'api' que irá ser instanciada passado o arquivo das credenciais
~~~javascript
const api = new API('./credentials/credentials.json')
~~~

3. Crie uma constante que irá receber o ID da apresentação, esse ID é encontrado na URL do arquivo. Substitua <id da planilha> pelo seu ID.

~~~javascript
const presentationId = '<id da planilha>'; 
~~~

### Copiando para um novo arquivo

~~~javascript
//Copiando o template para um novo arquivo
const newPresentation = await api.copyPresentation(presentationId);          
console.log(`\nApresentação copiada --> name: ${newPresentation.name}, id: ${newPresentation.id}`);
~~~

### Compartilhando acesso ao slide

~~~javascript
const emailAddress = '<email que deve receber acesso>'
const role = 'writer'  // owner | organizer | fileOrganizer | writer | commenter | reader
const type = 'user'  // user | group | domain | anyone
const sendNotificationEmails = false 
api.sharePresentation(id, emailAddress, role, type, sendNotificationEmails)
.then(() => console.log(`Presentation with ID ${id} successfully shared with ${emailAddress}!`))
~~~

### Substituindo texto dentro do arquivo

1. O TextReplacement recebe como parâmetro o texto que deve ser localizado e o texto que será colocado no lugar. Ex: TextReplacement('Nome', 'Nome Completo'), nesse caso toda correspondência de nome seria substituída para Nome Completo no arquivo.

~~~javascript
api.presentationBatchUpdate(id, [
     new TextReplacement('', ''),
     new TextReplacement('', ''),
     new TextReplacement('', ''),
])
.catch(error => console.log('Error - ', error))
~~~

## Google Slides

### Envio de Email

1. Importar o nodemailer no arquivo:

~~~javascript
import nodemailer from 'nodemailer';
~~~

2. Definir o transportador nodemailer e preencher com as credenciais:

~~~javascript
let transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: {
        user: abc@gmail.com', 
        pass: '1234'
    }
});
~~~

3. Defina uma variável mailOptions. Ela deve conter informações que seu receptor deve saber sobre ele.

~~~javascript
let mailOptions = {
    from: 'abc@gmail.com', 
    to: 'cba@gmail.com',
    subject: 'Nodemailer - Test',
    text: 'Wooohooo it works!!'
};
~~~

4. Para enviar um email é necessário invocar a função sendMail. Se tudo correr bem, você deve receber um e-mail.
~~~javascript
let mailOptions = {
    from: 'abc@gmail.com', 
    to: 'cba@gmail.com',
    subject: 'Test',
    text: 'Test!!'
};
~~~

### Links uteis:
1. [Documentação google-spreadsheet](https://www.npmjs.com/package/google-spreadsheet)
2. [Localizar o arquivo de credenciais](https://www.youtube.com/watch?v=TjYIF45IwjQ)
3. [Documentação google-slides](https://www.npmjs.com/package/google-slides#installation)
4. [Nodemailer](https://github.com/accimeesterlin/nodemailer-examples/tree/master/sendEmailWithGmail)