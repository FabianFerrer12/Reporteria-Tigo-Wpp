const qrcode = require('qrcode-terminal');
const express = require('express');

const { Client, LocalAuth, MessageMedia } = require('whatsapp-web.js');
const { initReporte } = require('./genreportimage');

const app = express();

app.use(express.urlencoded({ extended: true }));

const sendWhitApi = async (req, res) => {
    const { to, msg } = req.body;
    await initReporte()

    let from = '120363114662627150@g.us';

    sendMessage(from, 'Reporte');
    sendMedia(from, 'aprovisionamiento.png');
    sendMedia(from, 'aprovisionamientobsc.png');
    sendMedia(from, 'aseguramiento.png');
    sendMedia(from, 'aseguramientobsc.png');
    sendMedia(from, 'aseguramientorepinf.png');

    res.send({status: 'Enviado'})
}

app.post('/sendreport', sendWhitApi);

const client = new Client({
    puppeteer: {
        headless: false,
        executablePath: 'C:/Program Files/Google/Chrome/Application/chrome.exe',
    },
    authStrategy: new LocalAuth({
        dataPath: './wwebjs_auth_local'
    })
});

client.on('qr', (qr) => {
    console.log('QR RECEIVED', qr);
    qrcode.generate(qr, {small: true});
});

/* ----------------------------- VALIDA SI YA ESTA LISTO EL WHATSAPP ---------------------------- */
client.on('ready', async () => {
    console.log('Cliente esta listo!!');
});

/* ------------------------------- ESCUCHA LOS MENSAJES RECIBIDOS ------------------------------- */
client.on('message', (msg) => {
    const {from, to, body} = msg;

    // Este bug lo reporto Lucas Aldeco Brescia para evitar que se publiquen estados
    if (from === 'status@broadcast') {
        return
    }

    // console.log(`Message received from ${msg.from}: ${msg.body}`);
    /* console.log(from);
    console.log(to);
    console.log(body); */

    //sendMessage(from, 'Reporte de aprovisionamiento')
    //sendMedia(from, 'aprovisionamiento.png');
});

const sendMedia = (to, file) => {
    const mediaFile = MessageMedia.fromFilePath(`./images/${file}`);
    client.sendMessage(to, mediaFile);
}

const sendMessage = (to, msg) => {
    client.sendMessage(to, msg);
}

/* --------------------- VALIDA SI HUBO UNA FALLA EN LA AUTENTICACION POR QR -------------------- */
client.on('auth_failure', () => {
    console.log('Error en la autenticacion vuelva a escanear un QR');
});

client.initialize();

app.listen(9000, () => {
    console.log('API LEVANTADO PORT: 9000');
})