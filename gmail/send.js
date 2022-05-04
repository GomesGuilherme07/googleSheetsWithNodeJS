import nodemailer from 'nodemailer';
import {emailCredentials} from '../credentials/emailCredentials.js'

process.env['NODE_TLS_REJECT_UNAUTHORIZED'] = 0;

export async function sendMail(recipient, title, textMail){

    let transporter = nodemailer.createTransport({
        service: 'gmail',
        auth: {
            user: emailCredentials[0].EMAIL, 
            pass: emailCredentials[0].PASSWORD
        }
    });
    
    let mailOptions = {
        from: emailCredentials[0].EMAIL, 
        to: recipient,
        subject: title,
        text: textMail
    };
    
    transporter.sendMail(mailOptions, (err, data) => {
        if (err) {
            return console.log('\nError occurs - ', err);
        }
        return console.log(`\nEmail sent to ${recipient} !!!`);
    });

}


