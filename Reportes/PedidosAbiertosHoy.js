const nodeHtmlToImage = require('node-html-to-image');
const fs = require('fs');
const mssqlDB = require('../database/conn-mssql');

const getDataClick = async (categoria) => {
    try {

        const result = await mssqlDB.dbConnectionMssql(`

        SELECT r.DateTimeExtraction, r.Region, COUNT (r.CallID) AS 'Ordenes' FROM(
        SELECT
        GETDATE() DateTimeExtraction,
        CONVERT(VARCHAR(8),FORMAT(GETDATE(), 'dd-MM-yy')) AS FechaExtraccion,
        left(CONVERT(varchar(8), GETDATE(), 108),2) AS HoraExtraccion,
        TK.CallID,
        TK.UNEPedido,
        STA.Name AS Estado,
        DIS.Name AS Distrito,
        TK.UNEMunicipio,
        CASE WHEN DIS.Name IN (
         'BARRANC',
         'BARRANC_2',
         'BARRANC_3'
        ) THEN
         'Barrancabermeja' WHEN DIS.Name IN ('BUGA') THEN
         'Buga' WHEN DIS.Name IN ('IGUANA') THEN
         'Antioquia Norte' WHEN RG.Name IN ('Otros_Municipios_Norte') THEN
         'Guajira' WHEN DIS.Name IN ('TULUA') THEN
         'Tulua' WHEN DIS.Name IN (
         'MICROZONA 10',
         'MICROZONA 22',
         'MICROZONA 21'
        ) THEN
         'Cartago' WHEN DIS.Name IN ('MICROZONA 20') THEN
         'Risaralda' WHEN RG.Name IN (
         'Cundinamarca Sur',
         'Cundinamarca Default',
         'Cundinamarca Norte'
        ) THEN
         'Bogota' ELSE
         RG.Name END AS Region,
        CASE WHEN RG.Name IN (
         'Cundinamarca Sur',
         'Cundinamarca Default',
         'Cundinamarca Norte'
        ) THEN
         'Bogota' WHEN RG.Name IN ('Meta') THEN
         'Sur' WHEN AR.Name IN ('Default', 'Default_Pais') THEN
         'Noroccidente' WHEN AR.Name IN ('Norte') THEN
         'Costa' WHEN RG.Name IN (
         'Nariño',
         'Buga',
         'Cauca',
         'Tulua',
         'Valle',
         'Caldas',
         'Cartago',
         'Quindio',
         'Risaralda',
         'Tolima',
         'Valle Quindío'
        ) THEN
         'Occidente' WHEN RG.Name IN (
         'Antioquia Municipios',
         'Antioquia_Edatel'
        ) THEN
         'Edatel' WHEN RG.Name IN ('Default_Pais') THEN
         'Edatel' ELSE
         AR.Name END AS AREA,
        CASE WHEN TT.Name in ('Reparacion Bronce Plus','Reparacion GPON Bronce Plus') THEN 'Aseguramiento BSC' ELSE TTC.Name
        END AS Categoria,
        TT.Name AS TipoTarea,
        SS.Name AS Sistema,
        CASE
        WHEN upper(TK.[UNEHoraCita]) IN ('HO FIJ ARM','HOR FIJ MAN','07:30','06:00','07:00','AM','MA','MAÑANA') THEN 'AM'
        WHEN upper(TK.[UNEHoraCita]) IN ('PM','18:00','TARDE') THEN 'PM'
        ELSE 'TD'
        END AS UNEHoraCita,
        TK.UNEFechaCita,
        TK.AppointmentStart AS FechaInicioCita,
        TK.AppointmentFinish AS FechaFinCita,
        AG.StartTime AS FechaInicioAsignado,
        AG.FinishTime AS FechaFinAsignado,
        TK.Duration, TK.UNETecnologias,
        TK.UNESegmento,
        TK.UNEUENcalculada,
        TK.UNEAgendamientos UNEAgendamientoClick,
        TK.CompletionDate AS FechaCumplido,
        TK.CancellationDate AS FechaCancelado,
        TK.TimeModified AS FechaModificado,
        TK.EngineerID AS Tecnico,
        TK.EngineerName,
        TK.UNEProvisioner,
        TK.OnSiteDate,
        (DATEDIFF( MINUTE, TK.OnSiteDate,GETDATE())) "TiempoenSitio"
        FROM W6TASKS AS TK LEFT OUTER JOIN
        W6UNESOURCESYSTEMS AS SS ON TK.UNESourceSystem = SS.W6Key LEFT OUTER JOIN
        W6REGIONS AS RG ON TK.Region = RG.W6Key LEFT OUTER JOIN
        W6AREA AS AR ON AR.W6Key = TK.Area LEFT OUTER JOIN
        W6ASSIGNMENTS AS AG ON AG.Task = TK.W6Key LEFT OUTER JOIN
        W6TASK_STATUSES AS STA ON TK.Status = STA.W6Key LEFT OUTER JOIN
        W6DISTRICTS AS DIS ON TK.District = DIS.W6Key LEFT OUTER JOIN
        W6TASK_TYPES AS TT ON TK.TaskType = TT.W6Key LEFT OUTER JOIN
        W6TASKTYPECATEGORY AS TTC ON TK.TaskTypeCategory = TTC.W6Key
        WHERE (
        (
        TK.AppointmentStart > DATEADD(day, +0, CONVERT (date, SYSDATETIME()))
        AND TK.AppointmentStart < DATEADD(day, +1, CONVERT (date, SYSDATETIME()))
        AND STA.Name IN ('Abierto')
        AND TTC.Name IN ('Aprovisionamiento','Aseguramiento','Aprovisionamiento BSC','Mantenimiento Integral','Reconexion','Corte')
        AND TK.UNEMunicipio IS NOT NULL
        and TK.UNEHoraCita<>'Noche'
        )
        )
        ) AS r
        --ORDER BY TK.CallID, TK.CompletionDate DESC
        WHERE
         r.Categoria IN ('${categoria}') AND r.Region != 'Default_Pais' AND r.Region != 'Antioquia Default' AND r.Region != 'Antioquia Municipios' AND r.TipoTarea != 'Reparacion Infraestructura'
        GROUP BY
         r.DateTimeExtraction,
         r.AREA,
         r.Region,
         r.Categoria
        ORDER BY
         r.AREA ASC,
         r.Region ASC;`);

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



const generarImagen = async (titulo, namefile, data, fecha_full) => {

    try {

        const image = fs.readFileSync('./assets/LogoTigoBlanco.png');
        const base64Image = new Buffer.from(image).toString('base64');
        const dataURI = 'data:image/jpeg;base64,' + base64Image

        let contentHtml = `<!DOCTYPE html>
        <html>
        <head>
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
        <style>
        body {    
        width: 480px;
        font-family: arial, sans-serif;
        font-size: 11px
        }
        
        table {
        border-collapse: collapse;
        width: 100%;
        }
        
        td, th {
        border: 1px solid #dddddd;
        text-align: left;
        padding: 5px 5px 5px 5px;
        }
        
        tr:nth-child(even) {
        background-color: #dddddd;
        }
        
        table td:nth-child(2) {
        width: 25%;
        }
        
        .text-right {
        text-align: right
        }
        
        .text-center {
        text-align: center
        }
        </style>
        </head>
        <body>
        
        <table>
        <thead>
            <tr style="background-color:#1F3764; color: white;">
                <th colspan="1" style="text-align: center; border-right-color: #1F3764 !important;"><img src="{{imageSource}}" style="width: 50px;"></th>
                <th colspan="4" style="text-align: center; vertical-align: bottom; font-size: 20px; border-right-color: #1F3764 !important;">${titulo}</th>
                <th colspan="2" style="text-align: right; vertical-align: bottom; font-size: 8px;">Fecha actualización: ${fecha_full}</th>
            </tr>
            <tr style="background-color:#1F3764; color: white;">
                <th colspan="2">Region</th>
                <th colspan="6">Ordenes</th>
        </thead>
            <tbody>`;

            data.forEach(val => {
                contentHtml += `<tr>
                    <td colspan="2">${val.Region}</td>
                    <td colspan="6" class="text-right">${val.Ordenes}</td>
                </tr>`;
            });

        contentHtml += `</tbody>
        </table>
        
        </body>
        </html>`;

        let res = await nodeHtmlToImage({
            output: `./images/PedidosAbiertosHoy/${namefile}.png`,
            html: contentHtml,
            content: { imageSource: dataURI }
        })
        .then(() => `Imagen de ${namefile} creado con exito`);

        return res;

    } catch (error) {
        console.log(error);
    }
}


const initReportePedidos = async () => {

    try {

        let date = new Date();
        let gethour = date.getHours();

        let anio = date.getFullYear();
        let mes = ((date.getMonth() + 1) < 10) ? '0'+(date.getMonth() + 1) : (date.getMonth() + 1);
        let dia = ((date.getDate()) < 10) ? '0'+(date.getDate()) : (date.getDate());
        let hora = ((date.getHours()) < 10) ? '0'+(date.getHours()) : (date.getHours());
        let minutos = ((date.getMinutes()) < 10) ? '0'+(date.getMinutes()) : (date.getMinutes());
        let segundos = ((date.getSeconds()) < 10) ? '0'+(date.getSeconds()) : (date.getSeconds());

        let fecha_full = `${anio}-${mes}-${dia} ${hora}:${minutos}:${segundos}`;

        const [
            aprovisionamiento,
            aprovisionamientobsc,
            aseguramiento,
            aseguramientobsc,
        ] = await Promise.all([
            getDataClick('Aprovisionamiento'),
            getDataClick('Aprovisionamiento BSC'),
            getDataClick('Aseguramiento'),
            getDataClick('Aseguramiento BSC'),
        ]);

        let dataaprovisionamiento = []
        for (const value of aprovisionamiento) {
            dataaprovisionamiento.push(value);
        }
        let aprov = await generarImagen('Aprovisionamiento', 'aprovisionamiento', dataaprovisionamiento, fecha_full);

        let dataaprovisionamientobsc = []
        for (const value of aprovisionamientobsc) {
            dataaprovisionamientobsc.push(value);
        }
        let aprovbsc = await generarImagen('Aprovisionamiento BSC', 'aprovisionamientobsc', dataaprovisionamientobsc, fecha_full);

        let dataaseguramiento = []
        for (const value of aseguramiento) {
            dataaseguramiento.push(value);
        }
        let aseg = await generarImagen('Aseguramiento', 'aseguramiento', dataaseguramiento, fecha_full);

        let dataaseguramientobsc = []
        for (const value of aseguramientobsc) {
            dataaseguramientobsc.push(value);
        }
        let asegbsc = await generarImagen('Aseguramiento BSC', 'aseguramientobsc', dataaseguramientobsc, fecha_full);


        console.log('dataaprovisionamiento', aprov);
        console.log('dataaprovisionamientobsc', aprovbsc);
        console.log('dataaseguramiento', aseg);
        console.log('dataaseguramientobsc', asegbsc);

    } catch (error) {

        console.log('Error ejecución:', error);

    }
};

module.exports = {
    initReportePedidos
}
