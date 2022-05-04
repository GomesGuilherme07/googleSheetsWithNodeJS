import { buscandoDadosGoogleSheet } from './sheets/index.js';
import { copySlide } from './slides/index.js';
import { sendMail } from './gmail/send.js';

//Id dos arquivos
const SHEET_ID = '1gSx6ASI6dxspN6fMzVM7puY3zaBhCjmcccDAyM3AlxM';
const SLIDE_ID = '1dxTMrDV9rjSp4FMNM6Ic0M3zDB7w3hQV6p9SesketX4';

//Alterar informações slide
async function run(){

    let dados = await buscandoDadosGoogleSheet(SHEET_ID);
    let presentation = await copySlide(SLIDE_ID, dados);     

    let id = presentation.id;
    let title = `Envio apresentação - ${dados[0].Nome} ${dados[0].Sobrenome}`;
    let textMail = `Segue abaixo link com acesso apresentação \nhttps://docs.google.com/presentation/d/${id}/edit#slide=id.p`

    await sendMail(dados[0].Email, title, textMail);    

}

run();
