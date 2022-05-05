import nodemailer from 'nodemailer';
import {emailCredentials} from '../../credentials/emailCredentials.js'

process.env['NODE_TLS_REJECT_UNAUTHORIZED'] = 0;

export async function sendMail(dados, presentation){


    //Definindo as variáveis do e-mail
    let id = presentation.id;
    let title = `Envio apresentação - ${dados[0].Nome} ${dados[0].Sobrenome}`;
    let textMail = `Segue abaixo link com acesso apresentação \nhttps://docs.google.com/presentation/d/${id}/edit#slide=id.p`

    //Definindo o transportador
    let transporter = nodemailer.createTransport({
        service: 'gmail',
        auth: {
            user: emailCredentials[0].EMAIL, 
            pass: emailCredentials[0].PASSWORD
        }
    });
    
    //Definindo uma variável com as informações do envio
    let mailOptions = {
        from: emailCredentials[0].EMAIL, 
        to: dados[0].Email,
        subject: title,
        text: textMail
    };
    
    //Chamando a função para enviar o e-mail
    transporter.sendMail(mailOptions, (err, data) => {
        if (err) {
            return console.log('\nError occurs - ', err);
        }
        return console.log(`\nEmail sent to ${dados[0].Email} !!!`);
    });

}


