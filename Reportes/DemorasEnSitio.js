const ObjectsToCsv = require('objects-to-csv');
const Excel = require('exceljs');
const ExcelJS = require('exceljs');
const mssqlDB = require('../database/conn-mssql');

const getDataClick = async () => {
    try {

        const result = await mssqlDB.dbConnectionMssql(`

        SELECT R.Region,R.EngineerID AS 'Cedula',R.EngineerName AS 'Nombre', r.CallID AS 'Tarea',
(CASE WHEN (R.DiagnosticoFinal = 'OK' OR R.DiagnosticoFinal = 'En progreso') AND(R.RequierePrueba = 0 OR R.RequierePrueba = -1) THEN 0 WHEN (R.DiagnosticoFinal = 'ERROR' OR R.DiagnosticoFinal = 'Timeout') AND (R.RequierePrueba = 0 OR R.RequierePrueba = -1) THEN 1 WHEN R.DiagnosticoFinal IS NULL AND R.RequierePrueba = 0 THEN -3 WHEN R.DiagnosticoFinal IS NULL AND R.RequierePrueba = -1 THEN -2 ELSE -4 END) AS TYD_Categoria2, 
SUM(CASE WHEN r.SerialNo IS NOT NULL THEN 1 ELSE 0 END) AS Cantidad_Equipos_Activados, 
DATEDIFF(second, R.OnSiteDate, GETDATE()) AS Segundos
FROM (
SELECT TK.CallID, TK.UNEPedido, TK.UNEDireccion, TK.UNEMunicipio, EST.Name Estado, SS.Name SS, CAT.Name Proceso, TT.Name TaskType, REG.Name Region, DIS.Name Distrito,
TK.EngineerID, TK.EngineerName,TK.AppointmentStart, TK.AppointmentFinish,
TK.ScheduleDate, ASSG.StartTime, ASSG.FinishTime,TK.OnSiteDate, GETDATE() AS Hora_Actual,TK.Duration,TK.UNETestAndDiagnoseRequired RequierePrueba, 
TK.UNETestAndDiagnoseID PruebaSMNET,
TK.UNEActivateTestAndDiagnoseResult DiagnosticoFinal, TK.UNEActivateTestAndDiagnoseResultDetails DetalleReultado,TK.UNETecnologias,TK.CompletionDate, EQ.SerialNo,
(DATEDIFF(
 MI,
 TK.OnSiteDate, GETDATE()
)
) "TiempoenSitio"
FROM W6TASKS TK
INNER JOIN W6TASKS_REQUIRED_SKILLS1 RQSK ON RQSK.W6Key=TK.W6Key
INNER JOIN W6SKILLS SK ON SK.W6Key=RQSK.SkillKey
INNER JOIN W6TASK_STATUSES EST ON TK.Status=EST.W6Key
INNER JOIN W6UNESOURCESYSTEMS SS ON TK.UNESourceSystem=SS.W6Key
INNER JOIN W6TASKTYPECATEGORY CAT ON TK.TaskTypeCategory=CAT.W6Key
INNER JOIN W6TASK_TYPES TT ON TK.TaskType=TT.W6Key
INNER JOIN W6REGIONS REG ON TK.Region=REG.W6Key
INNER JOIN W6DISTRICTS DIS ON TK.District=DIS.W6Key
INNER JOIN W6ASSIGNMENTS ASSG ON ASSG.Task=TK.W6Key
INNER JOIN W6TASKS_UNEEQUIPMENTSUSED TK_EQ ON TK.W6Key=TK_EQ.W6Key
INNER JOIN W6UNEEQUIPMENTUSED EQ ON TK_EQ.UNEEquipmentUsedKey=EQ.W6Key
WHERE ASSG.StartTime>= CONVERT (DATE, SYSDATETIME()) AND ASSG.StartTime< DATEADD(DAY, 1, CONVERT (DATE, SYSDATETIME())) AND EST.W6Key IN (124135429) AND (CAT.Name IN('Aprovisionamiento', 'Aseguramiento')) AND EST.Name IN ('En Sitio') AND EQ."Type" IN ('Install','Traslado')) AS R
WHERE R.Tiempoensitio >120 AND (R.DiagnosticoFinal = 'OK' OR R.RequierePrueba = 0) AND R.TaskType LIKE ('%Nuevo%')
GROUP BY r.CallID,R.EngineerID,R.EngineerName, R.Region,R.OnSiteDate,R.CompletionDate, R.DiagnosticoFinal, R.RequierePrueba, R.Proceso, R.TiempoenSitio
ORDER BY R.Region, R.EngineerName ASC ;`);

        if (!result) {
            console.log('Error en la conexion de la BD', result);
            return false;
        }

        if (result.recordset.length == 0) {
            console.log('Sin datos para listar', result.recordset.length);
            return false;
        }

        let res = result.recordset;

        return res;
    } catch (error) {
        console.log('ERROR: ', error);
    }
}



// const generarImagen = async (titulo, namefile, data, fecha_full) => {

//     try {

//         const image = fs.readFileSync('./assets/LogoTigoBlanco.png');
//         const base64Image = new Buffer.from(image).toString('base64');
//         const dataURI = 'data:image/jpeg;base64,' + base64Image

//         let contentHtml = `<!DOCTYPE html>
//         <html>
//         <head>
//         <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
//         <style>
//         body {
//         width: 1700px;
//         font-family: arial, sans-serif;
//         font-size: 11px
//         }

//         table {
//         border-collapse: collapse;
//         width: 100%;
//         }

//         td, th {
//         border: 1px solid #dddddd;
//         text-align: left;
//         padding: 5px 5px 5px 5px;
//         }

//         tr:nth-child(even) {
//         background-color: #dddddd;
//         }

//         table td:nth-child(2) {
//         width: 15%;
//         }

//         .text-right {
//         text-align: right
//         }

//         .text-center {
//         text-align: center
//         }
//         </style>
//         </head>
//         <body>

//         <table>
//         <thead>
//             <tr style="background-color:#1F3764; color: white;">
//                 <th colspan="2" style="text-align: center; border-right-color: #1F3764 !important;"><img src="{{imageSource}}" style="width: 50px;"></th>
//                 <th colspan="3" style="text-align: center; vertical-align: bottom; font-size: 20px; border-right-color: #1F3764 !important;">${titulo}</th>
//                 <th colspan="2" style="text-align: right; vertical-align: bottom; font-size: 8px;">Fecha actualización: ${fecha_full}</th>
//             </tr>
//             <tr style="background-color:#1F3764; color: white;">
//                 <th colspan="1" style="text-align: center; vertical-align: bottom; font-size: 16px;">Region</th>
//                 <th colspan="1" style="text-align: center; vertical-align: bottom; font-size: 16px;">Cedula Tecnico</th>
//                 <th colspan="1" style="text-align: center; vertical-align: bottom; font-size: 16px;">Nombre Tecnico</th>
//                 <th colspan="1" style="text-align: center; vertical-align: bottom; font-size: 16px;">Tarea</th>
//                 <th colspan="1" style="text-align: center; vertical-align: bottom; font-size: 16px;">Tiempo en sitio</th>
//                 <th colspan="1" style="text-align: center; vertical-align: bottom; font-size: 16px;">TYD_Categoria2</th>
//                 <th colspan="1" style="text-align: center; vertical-align: bottom; font-size: 16px;">Ingreso EQ</th>
//         </thead>
//             <tbody>`;

//             data.forEach(val => {

//                 let segundosP = val.Segundos
//                 let segundos = (Math.round(segundosP % 0x3C)).toString();
//                 let horas    = (Math.floor(segundosP / 0xE10)).toString();
//                 let minutos  = (Math.floor(segundosP / 0x3C ) % 0x3C).toString();

//                 let eq = ""
//                 if(val.Cantidad_Equipos_Activados > 1){
//                     eq = "SI"
//                 }else{
//                     eq= "NO"
//                 }

//                 let tyd = ""
//                 if(val.TYD_Categoria2 == -3){
//                     tyd = '<i class="fa fa-flag" aria-hidden="true"></i>'
//                 }else if(val.TYD_Categoria2 == 0){
//                     tyd = '<i class="fa fa-check-square" aria-hidden="true"></i>'
//                 }
//                 contentHtml += `<tr>
//                     <td colspan="1" style="text-align: center; vertical-align: bottom; font-size: 17px;">${val.Region}</td>
//                     <td colspan="1" class="text-right" style="text-align: center; vertical-align: bottom; font-size: 15px;">${val.Cedula}</td>
//                     <td colspan="1" class="text-right" style="text-align: center; vertical-align: bottom; font-size: 15px;">${val.Nombre}</td>
//                     <td colspan="1" class="text-right" style="text-align: center; vertical-align: bottom; font-size: 15px;">${val.Tarea}</td>
//                     <td colspan="1" class="text-right" style="text-align: center; vertical-align: bottom; font-size: 15px;">${horas}:${minutos}:${segundos}</td>
//                     <td colspan="1" class="text-right" style="text-align: center; vertical-align: bottom; font-size: 15px;">${tyd}</td>
//                     <td colspan="1" class="text-right" style="text-align: center; vertical-align: bottom; font-size: 15px;">${eq}</td>
//                 </tr>`;
//             });

//         contentHtml += `</tbody>
//         </table>

//         </body>
//         </html>`;

//         let res = await nodeHtmlToImage({
//             output: `./images/DemorasEnSitio/${namefile}.png`,
//             html: contentHtml,
//             content: { imageSource: dataURI }
//         })
//         .then(() => `Imagen de ${namefile} creado con exito`);

//         return res;

//     } catch (error) {
//         console.log(error);
//     }
// }

const generarExcel = async(data)=>{

    const headers = [
        { header: 'Region', key: 'Region', width: 15 },
        { header: 'Cedula', key: 'Cedula', width: 15 },
        { header: 'Nombre', key: 'Nombre', width: 50 },
        { header: 'Tarea', key: 'Tarea', width: 15 },
        { header: 'Tiempo En Sitio', key: 'TiempoEnSitio', width: 15 },
        { header: 'TYD_Categoria2', key: 'TYD_Categoria2', width: 15 },
        { header: 'Ingreso_EQ', key: 'Ingreso_EQ', width: 15 },
    ]


    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('My Sheet');

    worksheet.columns = headers;

    itemToSave = {};
        itemsToSave = [];   
        data.forEach(data=>{
            let segundosP = data.Segundos
            let segundos = (Math.round(segundosP % 0x3C)).toString();
            let horas    = (Math.floor(segundosP / 0xE10)).toString();
            let minutos  = (Math.floor(segundosP / 0x3C ) % 0x3C).toString();
            let tiempo =  horas+':'+minutos+':'+segundos
            let tyd = ""

            if(data.TYD_Categoria2 == -3){
                tyd = 'No requiere T&D'
            }else if(data.TYD_Categoria2 == 0){
                tyd = 'OK'
            }

            let eq = ""
            if(data.Cantidad_Equipos_Activados > 1){
                eq = "SI"
            }else{
                eq= "NO"
            }

            itemToSave.Region= data.Region
            itemToSave.Cedula= data.Cedula
            itemToSave.Nombre= data.Nombre
            itemToSave.Tarea= data.Tarea
            itemToSave.TiempoEnSitio  = tiempo
            itemToSave.TYD_Categoria2= tyd
            itemToSave.Ingreso_EQ= eq
            itemsToSave.push(itemToSave);
            itemToSave = {};
        });
    worksheet.addRows(itemsToSave);

    await workbook.xlsx.writeFile('./images/DemorasEnSitio/Reporte_Tecnicos_Demora_En_Sito.xlsx')
}

const initReporteDemorasEnSitio = async () => {

    try {
        const [
            aprovisionamiento,
        ] = await Promise.all([
            getDataClick(),
        ]);

        let dataaprovisionamiento = []
        for (const value of aprovisionamiento) {
            dataaprovisionamiento.push(value);
        }

       await generarExcel( dataaprovisionamiento);

        
    } catch (error) {

        console.log('Error ejecución:', error);

    }
};

module.exports = {
    initReporteDemorasEnSitio
}
