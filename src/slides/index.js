import { API, TextReplacement, ShapeReplacementWithImage, Presentation } from 'google-slides'

process.env['NODE_TLS_REJECT_UNAUTHORIZED'] = 0;
 
const api = new API('./credentials/credentials.json')

export async function copySlide(presentationId, dados){

     try{

          //Copiando o template para um novo arquivo
          const newPresentation = await api.copyPresentation(presentationId); 
          
          console.log(`\nname: ${newPresentation.name}, id: ${newPresentation.id}`);   

          const newPresentationId = newPresentation.id;          
          
          await shareAccess(newPresentationId, dados);
          await replaceText(newPresentationId, dados);

          return newPresentation;

     }catch(e){
          console.log('\nError - ', e);
     }
}

async function shareAccess(id, dados){

     try{
          
          const emailAddress = dados[0].Email
          const role = 'writer'  // owner | organizer | fileOrganizer | writer | commenter | reader
          const type = 'user'  // user | group | domain | anyone
          const sendNotificationEmails = false 

          //Compartilhando acesso a apresentação gerada
          api.sharePresentation(id, emailAddress, role, type, sendNotificationEmails)
          .then(() => console.log(`\nPresentation with ID ${id} successfully shared with ${emailAddress}!`))
          
     }catch(e){
          console.log('Error - ', e);
     }

}


async function replaceText(id, dados){

     try{
          //Substituindo texto dentro do arquivo
          api.presentationBatchUpdate(id, [
               new TextReplacement('%nome%', dados[0].Nome),
               new TextReplacement('%sobrenome%', dados[0].Sobrenome),
               new TextReplacement('%email%', dados[0].Email)
          ])        

          setTimeout(function(){
               console.log('\nInformações substituidas na apresentação');
           }, 3000); 

     }catch(e){
          console.log('\nError - ', e);
     }

}

