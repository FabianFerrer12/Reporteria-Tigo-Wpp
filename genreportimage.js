const { Op } = require("sequelize");
const nodeHtmlToImage = require('node-html-to-image');
const fs = require('fs');

const mssqlDB = require('./database/conn-mssql');
const db = require("./models");

const getDataClick = async (categoria, estadosclick, horacita, tipotarea) => {
    try {

        const result = await mssqlDB.dbConnectionMssql(`SELECT
                                                            r.DateTimeExtraction,
                                                            r.AREA,
                                                            r.Region,
                                                            r.Categoria,
                                                            COUNT (r.CallID) AS 'cant_tareas',
                                                            COUNT (DISTINCT(r.Tecnico)) AS 'cant_tecnico',
                                                            CASE
                                                        WHEN CAST (
                                                            COUNT (DISTINCT(r.Tecnico)) AS NUMERIC
                                                        ) = 0 THEN
                                                            0
                                                        ELSE
                                                            CAST (
                                                                CAST (COUNT(r.CallID) AS NUMERIC) / CAST (
                                                                    COUNT (DISTINCT(r.Tecnico)) AS NUMERIC
                                                                ) AS DECIMAL (10, 1)
                                                            )
                                                        END AS 'ratio'
                                                        FROM
                                                            (
                                                                SELECT
                                                                    GETDATE() DateTimeExtraction,
                                                                    TK.CallID,
                                                                    TK.UNEPedido,
                                                                    STA.Name AS Estado,
                                                                    DIS.Name AS Distrito,
                                                                    TK.UNEMunicipio,
                                                                    CASE
                                                                WHEN DIS.Name IN (
                                                                    'BARRANC',
                                                                    'BARRANC_2',
                                                                    'BARRANC_3'
                                                                ) THEN
                                                                    'Barrancabermeja'
                                                                WHEN DIS.Name IN ('BUGA') THEN
                                                                    'Buga'
                                                                WHEN DIS.Name IN ('IGUANA') THEN
                                                                    'Antioquia Norte'
                                                                WHEN RG.Name IN ('Otros_Municipios_Norte') THEN
                                                                    'Guajira'
                                                                WHEN DIS.Name IN ('TULUA') THEN
                                                                    'Tulua'
                                                                WHEN DIS.Name IN (
                                                                    'MICROZONA 10',
                                                                    'MICROZONA 22',
                                                                    'MICROZONA 21'
                                                                ) THEN
                                                                    'Cartago'
                                                                WHEN DIS.Name IN ('MICROZONA 20') THEN
                                                                    'Risaralda'
                                                                WHEN RG.Name IN (
                                                                    'Cundinamarca Sur',
                                                                    'Cundinamarca Default',
                                                                    'Cundinamarca Norte'
                                                                ) THEN
                                                                    'Bogota'
                                                                ELSE
                                                                    RG.Name
                                                                END AS Region,
                                                                CASE
                                                            WHEN RG.Name IN (
                                                                'Cundinamarca Sur',
                                                                'Cundinamarca Default',
                                                                'Cundinamarca Norte'
                                                            ) THEN
                                                                'Bogota'
                                                            WHEN RG.Name IN ('Meta') THEN
                                                                'Sur'
                                                            WHEN AR.Name IN ('Default', 'Default_Pais') THEN
                                                                'Noroccidente'
                                                            WHEN AR.Name IN ('Norte') THEN
                                                                'Costa'
                                                            WHEN RG.Name IN (
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
                                                                'Occidente'
                                                            WHEN RG.Name IN (
                                                                'Antioquia Municipios',
                                                                'Antioquia_Edatel'
                                                            ) THEN
                                                                'Edatel'
                                                            WHEN RG.Name IN ('Default_Pais') THEN
                                                                'Edatel'
                                                            ELSE
                                                                AR.Name
                                                            END AS AREA,
                                                            CASE
                                                        WHEN TT.Name IN (
                                                            'Reparacion Bronce Plus',
                                                            'Reparacion GPON Bronce Plus',
                                                            'Reparacion Bronce Plus Internet',
                                                            'Reparacion GPON Bronce Plus Internet'
                                                        ) THEN
                                                            'Aseguramiento BSC'
                                                        ELSE
                                                            TTC.Name
                                                        END AS Categoria,
                                                        TT.Name AS TipoTarea,
                                                        SS.Name AS Sistema,
                                                        CASE
                                                        WHEN UPPER (TK.[UNEHoraCita]) IN (
                                                            'HO FIJ ARM',
                                                            'HOR FIJ MAN',
                                                            '07:30',
                                                            '06:00',
                                                            '07:00',
                                                            'AM',
                                                            'MA',
                                                            'MAÑANA'
                                                        ) THEN
                                                            'AM'
                                                        WHEN UPPER (TK.[UNEHoraCita]) IN ('PM', '18:00', 'TARDE') THEN
                                                            'PM'
                                                        ELSE
                                                            'TD'
                                                        END AS UNEHoraCita,
                                                        TK.UNEFechaCita,
                                                        TK.AppointmentStart AS FechaInicioCita,
                                                        TK.AppointmentFinish AS FechaFinCita,
                                                        AG.StartTime AS FechaInicioAsignado,
                                                        AG.FinishTime AS FechaFinAsignado,
                                                        TK.Duration,
                                                        TK.UNETecnologias,
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
                                                        (
                                                            DATEDIFF(
                                                                MINUTE,
                                                                TK.OnSiteDate,
                                                                COALESCE (
                                                                    TK.CompletionDate,
                                                                    GETDATE()
                                                                )
                                                            )
                                                        ) "TiempoenSitio"
                                                        FROM
                                                            W6TASKS AS TK
                                                        LEFT OUTER JOIN W6UNESOURCESYSTEMS AS SS ON TK.UNESourceSystem = SS.W6Key
                                                        LEFT OUTER JOIN W6REGIONS AS RG ON TK.Region = RG.W6Key
                                                        LEFT OUTER JOIN W6AREA AS AR ON AR.W6Key = TK.Area
                                                        LEFT OUTER JOIN W6ASSIGNMENTS AS AG ON AG.Task = TK.W6Key
                                                        LEFT OUTER JOIN W6TASK_STATUSES AS STA ON TK.Status = STA.W6Key
                                                        LEFT OUTER JOIN W6DISTRICTS AS DIS ON TK.District = DIS.W6Key
                                                        LEFT OUTER JOIN W6TASK_TYPES AS TT ON TK.TaskType = TT.W6Key
                                                        LEFT OUTER JOIN W6TASKTYPECATEGORY AS TTC ON TK.TaskTypeCategory = TTC.W6Key
                                                        WHERE
                                                            (
                                                                (
                                                                    TK.AppointmentStart > DATEADD(
                                                                        DAY,
                                                                        + 0,
                                                                        CONVERT (DATE, SYSDATETIME())
                                                                    )
                                                                    AND TK.AppointmentStart < DATEADD(
                                                                        DAY,
                                                                        + 1,
                                                                        CONVERT (DATE, SYSDATETIME())
                                                                    )
                                                                    AND STA.Name IN (
                                                                        ${estadosclick}
                                                                    )
                                                                    AND TTC.Name IN (
                                                                        'Aprovisionamiento',
                                                                        'Aseguramiento',
                                                                        'Aprovisionamiento BSC'
                                                                    )
                                                                    AND TT.Name NOT IN (
                                                                        'Reparacion Infraestructura'
                                                                    )
                                                                    ${horacita}
                                                                    AND TK.UNEMunicipio IS NOT NULL
                                                                )
                                                                OR (
                                                                    AG.StartTime > DATEADD(
                                                                        DAY,
                                                                        + 0,
                                                                        CONVERT (DATE, SYSDATETIME())
                                                                    )
                                                                    AND AG.StartTime < DATEADD(
                                                                        DAY,
                                                                        + 1,
                                                                        CONVERT (DATE, SYSDATETIME())
                                                                    )
                                                                    AND STA.Name IN (
                                                                        ${estadosclick}
                                                                    )
                                                                    AND TTC.Name IN (
                                                                        'Mantenimiento Integral',
                                                                        'Reconexion',
                                                                        'Corte',
                                                                        'Reparacion Bronce Plus Internet',
                                                                        'Reparacion GPON Bronce Plus Internet',
                                                                        'Reparacion Infraestructura'
                                                                    )
                                                                    AND TK.UNEMunicipio IS NOT NULL
                                                                )
                                                                OR (
                                                                    STA.Name IN (
                                                                        ${estadosclick}
                                                                    )
                                                                    AND TTC.Name IN (
                                                                        'Mantenimiento Integral',
                                                                        'Reconexion',
                                                                        'Corte',
                                                                        'Reparacion Bronce Plus Internet',
                                                                        'Reparacion GPON Bronce Plus Internet',
                                                                        'Reparacion Infraestructura'
                                                                    )
                                                                    AND TK.UNEMunicipio IS NOT NULL 
                                                                )
                                                            )  --ORDER BY TK.CallID, TK.CompletionDate DESC
                                                            ) AS r
                                                        WHERE
                                                            r.Categoria IN ('${categoria}') 
                                                            AND r.Region != 'Default_Pais'
                                                            AND r.Region != 'Antioquia Default'
                                                            AND r.Region != 'Antioquia Municipios'
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

const processRatio = async (region, categoria, gethour) => {

    try {

        let hour = 0;

        if (gethour >= 5 &&  gethour < 6) {
            hour = 5;
        } else if (gethour >= 6 &&  gethour < 7) {
            hour = 6;
        } else if (gethour >= 7 &&  gethour < 10) {
            hour = 7;
        } else if (gethour >= 10 &&  gethour < 12) {
            hour = 10;
        } else if (gethour >= 12 &&  gethour < 13) {
            hour = 12;
        } else if (gethour >= 13 &&  gethour < 14) {
            hour = 13;
        } else if (gethour >= 14 &&  gethour < 15) {
            hour = 14;
        } else if (gethour >= 15 &&  gethour < 16) {
            hour = 15;
        } else if (gethour >= 16 &&  gethour < 17) {
            hour = 17;
        } else if (gethour >= 17 &&  gethour < 18) {
            hour = 18;
        } else {
            hour = 18;
        }

        let dataratio = await db['Ratio'].findAll({
            where: {
                [Op.and]: [
                    { region: region },
                    { categoria: categoria },
                    { hora: hour }
                ]
            }
        });

        if (dataratio.length == 0) {
            return false;
        }

        return dataratio;

    } catch (error) {
        console.log('Error process ratio: ', error);
        throw('Error process ratio: ', error);
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
                <th>AREA</th>
                <th>REGION</th>
                <th>CANTIDAD ORDENES</th>
                <th>CANTIDAD FRENTES</th>
                <th>RATIO</th>
                <th colspan="2">RATIO ESPERADO</th>
            </tr>
        </thead>
            <tbody>`;

            data.forEach(val => {
                let estado = (val.validation) ? '<i class="fa fa-check" style="color:green;"></i>' : '<i class="fa fa-times" style="color:red;"></i>';
                contentHtml += `<tr>
                    <td>${val.AREA}</td>
                    <td>${val.Region}</td>
                    <td class="text-right">${val.cant_tareas}</td>
                    <td class="text-right">${val.cant_tecnico}</td>
                    <td class="text-right">${val.ratio}</td>
                    <td class="text-right">${val.ratioesperado}</td>
                    <td class="text-center">${estado}</td>
                </tr>`;
            });

        contentHtml += `</tbody>
        </table>
        
        </body>
        </html>`;

        let res = await nodeHtmlToImage({
            output: `./images/Inicio y seguimiento/${namefile}.png`,
            html: contentHtml,
            content: { imageSource: dataURI }
        })
        .then(() => `Imagen de ${namefile} creado con exito`);

        return res;

    } catch (error) {
        console.log(error);
    }
}


const initReporte = async () => {

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

        let estadosclick = '';
        let horacita = ''
        let tipotarea = '';

        if (gethour >= 7 &&  gethour < 10) {
            estadosclick = "\'Abierto\', \'Asignado\', \'Despachado\', \'En Camino\'";
            horacita = "AND UNEHoraCita = \'AM\'";
            tipotarea = "AND TT.W6Key != 124829709";
        }else if(gethour >= 10){
            estadosclick = "\'Abierto\', \'Asignado\', \'Despachado\', \'En Camino\'";
            horacita = "";
            tipotarea = "AND TT.W6Key != 124829709";
        }

        const [
            aprovisionamiento,
            aprovisionamientobsc,
            aseguramiento,
            aseguramientobsc,
        ] = await Promise.all([
            getDataClick('Aprovisionamiento', estadosclick, horacita, tipotarea),
            getDataClick('Aprovisionamiento BSC', estadosclick, horacita, tipotarea),
            getDataClick('Aseguramiento', estadosclick, horacita, tipotarea),
            getDataClick('Aseguramiento BSC', estadosclick, horacita, tipotarea),
        ]);

        let dataaprovisionamiento = []
        for (const value of aprovisionamiento) {
            /* let date = new Date();
            let gethour = date.getHours(); */

            let ratioesp = await processRatio(value.Region, value.Categoria, gethour)

            let ratiowait = (!ratioesp) ? 0 : ratioesp[0].ratioesperado;
            value.ratioesperado = ratiowait;
            value.validation = (value.ratio <= ratiowait) ? true : false;

            dataaprovisionamiento.push(value);
        }
        let aprov = await generarImagen('Aprovisionamiento', 'aprovisionamiento', dataaprovisionamiento, fecha_full);

        let dataaprovisionamientobsc = []
        for (const value of aprovisionamientobsc) {
            /* let date = new Date();
            let gethour = date.getHours(); */

            let ratioesp = await processRatio(value.Region, value.Categoria, gethour)

            let ratiowait = (!ratioesp) ? 0 : ratioesp[0].ratioesperado;
            value.ratioesperado = ratiowait;
            value.validation = (value.ratio <= ratiowait) ? true : false;

            dataaprovisionamientobsc.push(value);
        }
        let aprovbsc = await generarImagen('Aprovisionamiento BSC', 'aprovisionamientobsc', dataaprovisionamientobsc, fecha_full);

        let dataaseguramiento = []
        for (const value of aseguramiento) {
            /* let date = new Date();
            let gethour = date.getHours(); */

            let ratioesp = await processRatio(value.Region, value.Categoria, gethour)

            let ratiowait = (!ratioesp) ? 0 : ratioesp[0].ratioesperado;
            value.ratioesperado = ratiowait;
            value.validation = (value.ratio <= ratiowait) ? true : false;

            dataaseguramiento.push(value);
        }
        let aseg = await generarImagen('Aseguramiento', 'aseguramiento', dataaseguramiento, fecha_full);

        let dataaseguramientobsc = []
        for (const value of aseguramientobsc) {
            /* let date = new Date();
            let gethour = date.getHours(); */

            let ratioesp = await processRatio(value.Region, value.Categoria, gethour)

            let ratiowait = (!ratioesp) ? 0 : ratioesp[0].ratioesperado;
            value.ratioesperado = ratiowait;
            value.validation = (value.ratio <= ratiowait) ? true : false;

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
    initReporte
}
