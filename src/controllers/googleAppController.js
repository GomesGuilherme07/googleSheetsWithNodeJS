import { buscandoDadosGoogleSheet } from '../sheets/index.js';
import { copySlide } from '../slides/index.js';
import { sendMail } from '../gmail/send.js';
import { constId } from '../../credentials/constId.js'

process.env['NODE_TLS_REJECT_UNAUTHORIZED'] = 0;


class GoogleAppController{

    static runGoogleApp = async (req, res) => {

        try{

            let dados = await buscandoDadosGoogleSheet(constId[0].SHEET_ID);
            let presentation = await copySlide(constId[0].SLIDE_ID, dados);

            await sendMail(dados, presentation); 
            res.status(200).send({textResponse: `Dados buscados no Google Sheets, apresentação atualizada e e-mail enviado para o endereço: ${dados[0].Email}`});

        }catch(e){
            console.log('Error - ', e)
            res.status(500).send("Error - ", e);
        }
        
    }
    

}

export default GoogleAppController;
