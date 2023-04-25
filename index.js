const qrcode = require('qrcode-terminal');
const express = require('express');

const { Client, LocalAuth, MessageMedia } = require('whatsapp-web.js');
const { initReporte } = require('./Reportes/genreportimage');
const {initReportePedidos} =require('./Reportes/PedidosAbiertosHoy');
const { initReportTecnico } = require('./Reportes/porcentaje_tecnicos_en_sitio');
const {initReporteDemorasEnSitio} = require('./Reportes/DemorasEnSitio')
const app = express();

app.use(express.urlencoded({ extended: true }));

const sendWhitApi = async (req, res) => {
    const { to, msg } = req.body;
    await initReporte()

    let from = '120363114662627150@g.us';

    let date = new Date();
    let gethour = date.getHours();

    if(gethour == 7){
        sendMessage(from, 'Reporte inicio operacion nacional');
    }else{
        sendMessage(from, 'Reporte seguimiento operacion nacional');
    }

    
    sendMedia(from, 'Inicio y seguimiento/aprovisionamiento.png');
    sendMedia(from, 'Inicio y seguimiento/aprovisionamientobsc.png');
    sendMedia(from, 'Inicio y seguimiento/aseguramiento.png');
    sendMedia(from, 'Inicio y seguimiento/aseguramientobsc.png');

    res.send({status: 'Enviado'})
}

const sendPedidosHoy = async(req,res)=>{
    const { to, msg } = req.body;
    await initReportePedidos()

    let from = '120363114662627150@g.us';

    sendMessage(from, 'Pedidos en estado abierto con agenda para hoy');
    sendMedia(from, 'PedidosAbiertosHoy/aprovisionamiento.png');
    sendMedia(from, 'PedidosAbiertosHoy/aprovisionamientobsc.png');
    sendMedia(from, 'PedidosAbiertosHoy/aseguramiento.png');
    sendMedia(from, 'PedidosAbiertosHoy/aseguramientobsc.png');

    res.send({status: 'Enviado'})
}

const sendWhitApiPorcent = async (req, res) => {
    const { to, msg } = req.body;
    await initReportTecnico()

    let from = '120363114662627150@g.us';

    sendMessage(from, 'Reporte porcentaje tecnico');
    sendMedia(from, 'porcentaje/aprovisionamiento.png');
    sendMedia(from, 'porcentaje/aprovisionamientobsc.png');
    sendMedia(from, 'porcentaje/aseguramiento.png');
    sendMedia(from, 'porcentaje/aseguramientobsc.png');

    res.send({status: 'Enviado'})
}

const sendDemorasEnSitio = async (req, res) => {
    const { to, msg } = req.body;

    await initReporteDemorasEnSitio()
    let from = '120363114662627150@g.us';
    sendMessage(from, 'Reporte demoras en sitio');
    sendMedia(from, 'DemorasEnSitio/Reporte_Tecnicos_Demora_En_Sito.xlsx');
    res.send({status: 'Enviado'})
}

app.get('/sendreport', sendWhitApi);
app.get('/sendreportPorcentaje', sendWhitApiPorcent);
app.get('/sendPedidosHoy', sendPedidosHoy);
app.get('/sendDemoras', sendDemorasEnSitio);

const client = new Client({
    puppeteer: {
        headless: true
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

app.listen(9001, () => {
    console.log('API LEVANTADO PORT: 9001');
})