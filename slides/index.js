import { API, TextReplacement, ShapeReplacementWithImage, Presentation } from 'google-slides'

process.env['NODE_TLS_REJECT_UNAUTHORIZED'] = 0;
 
const api = new API('./credentials/credentials.json')
const presentationId = '1dxTMrDV9rjSp4FMNM6Ic0M3zDB7w3hQV6p9SesketX4'; 

//Copiando o template para um novo arquivo
const newPresentation = await api.copyPresentation(presentationId);
console.log(`Apresentação copiada --> name: ${newPresentation.name}, id: ${newPresentation.id}`);
console.log(newPresentation);
const id = newPresentation.id;

//Compartilhando acesso a apresentação gerada
const emailAddress = 'guilherme.gomes@take.net'
const role = 'writer'  // owner | organizer | fileOrganizer | writer | commenter | reader
const type = 'user'  // user | group | domain | anyone
const sendNotificationEmails = false 
api.sharePresentation(id, emailAddress, role, type, sendNotificationEmails)
.then(() => console.log(`Presentation with ID ${id} successfully shared with ${emailAddress}!`))

//Substituindo texto dentro do arquivo
api.presentationBatchUpdate(id, [
     new TextReplacement('%nome%', 'Guilherme'),
     new TextReplacement('%sobrenome%', 'Gomes'),
     new TextReplacement('%email%', 'guilherme.gomes@take.net'),
])
.catch(error => console.log('Error - ', error))

