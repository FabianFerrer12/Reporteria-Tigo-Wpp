const { Op } = require("sequelize");
const nodeHtmlToImage = require("node-html-to-image");
const fs = require("fs");

const mssqlDB = require("../database/conn-mssql");
const db = require("../models");

const getDataClick = async (categoria) => {
  try {
    const result = await mssqlDB.dbConnectionMssql(`
    
    BEGIN TRAN
    SELECT
        r.DateTimeExtraction,
        r.AREA,
        r.Region,
        r.Categoria,
        COUNT (DISTINCT r.Tecnico) AS cant_tareas,
        COUNT (
            DISTINCT CASE                 
            WHEN r.Estado IN (
                'Finalizada',
                'En Sitio',
                'Incompleto'
            )                  THEN
                r.Tecnico               
            END
        ) AS cant_tecnico
    FROM
        (
            SELECT DISTINCT
                GETDATE() DateTimeExtraction,
                TK.CallID,
                STA.Name AS Estado,
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
        'Reparacion GPON Bronce Plus'
    ) THEN
        'Aseguramiento BSC'
    ELSE
        TTC.Name
    END AS Categoria,
     TK.EngineerID AS Tecnico
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
                TK.OnSiteDate > DATEADD(
                    DAY,
                    + 0,
                    CONVERT (DATE, SYSDATETIME())
                )
                AND STA.Name IN (
                    'Incompleto',
                    'Finalizada',
                    'En Sitio'
                )
                AND TTC.Name IN (
                    'Aprovisionamiento',
                    'Aseguramiento'
                )
            )
            OR (
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
                    'Abierto',
                    'Asignado',
                    'Despachado',
                    'En Camino',
                    'En Sitio'
                )
                AND TTC.Name IN (
                    'Aprovisionamiento',
                    'Aseguramiento',
                    'Aprovisionamiento BSC'
                )
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
                    'Asignado',
                    'Despachado',
                    'En Camino'
                )
                AND TTC.Name IN (
                    'Mantenimiento Integral',
                    'Reconexion',
                    'Corte'
                )
            )
            OR (
                STA.Name IN ('Abierto')
                AND TTC.Name IN (
                    'Mantenimiento Integral',
                    'Reconexion',
                    'Corte'
                )
            )
        )
        ) AS r
    WHERE
        r.Categoria = '${categoria}'
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
        r.Region ASC
        
        COMMIT TRANSACTION ;`);
    if (!result) {
      console.log("Error en la conexion de la BD", result);
      return false;
    }

    if (result.recordset.length == 0) {
      console.log("Sin datos para listar", result.recordset.length);
    }

    let res = result.recordset;

    return res;
  } catch (error) {
    console.log("ERROR: ", error);
  }
};

const generarImagen = async (titulo, namefile, data, fecha_full) => {
  try {
    const image = fs.readFileSync("./assets/LogoTigoBlanco.png");
    const base64Image = new Buffer.from(image).toString("base64");
    const dataURI = "data:image/jpeg;base64," + base64Image;

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
                <th colspan="3" style="text-align: center; vertical-align: bottom; font-size: 20px; border-right-color: #1F3764 !important;">${titulo}</th>
                <th colspan="1" style="text-align: right; vertical-align: bottom; font-size: 8px;">Fecha actualización: ${fecha_full}</th>
            </tr>
            <tr style="background-color:#1F3764; color: white;">
                <th>AREA</th>
                <th>REGION</th>
                <th>CANTIDAD ORDENES</th>
                <th>CANTIDAD FRENTES</th>
                <th>PORCENTAJE</th>
            </tr>
        </thead>
            <tbody>`;

    data.forEach((val) => {
      let n = Math.round((val.cant_tecnico / val.cant_tareas) * 100);
      let color;
      if (n >= 95) {
        color = `#198754`;
      } else if (n >= 85 && n < 95) {
        color = `#ffc107`;
      } else if (n < 85) {
        color = `#bb2d3b`;
      }
      contentHtml += `<tr>
                    <td>${val.AREA}</td>
                    <td>${val.Region}</td>
                    <td class="text-right">${val.cant_tareas}</td>
                    <td class="text-right">${val.cant_tecnico}</td>
                    <td class="text-right" style="background: ${color}">${n}%</td>
                </tr>`;
    });

    contentHtml += `</tbody>
        </table>
        
        </body>
        </html>`;

    let res = await nodeHtmlToImage({
      output: `./images/porcentaje/${namefile}.png`,
      html: contentHtml,
      content: { imageSource: dataURI },
    }).then(() => `Imagen de ${namefile} creado con exito`);

    return res;
  } catch (error) {
    console.log(error);
  }
};

const initReportTecnico = async () => {
  try {
    let date = new Date();
    let gethour = date.getHours();

    let anio = date.getFullYear();
    let mes =
      date.getMonth() + 1 < 10
        ? "0" + (date.getMonth() + 1)
        : date.getMonth() + 1;
    let dia = date.getDate() < 10 ? "0" + date.getDate() : date.getDate();
    let hora = date.getHours() < 10 ? "0" + date.getHours() : date.getHours();
    let minutos =
      date.getMinutes() < 10 ? "0" + date.getMinutes() : date.getMinutes();
    let segundos =
      date.getSeconds() < 10 ? "0" + date.getSeconds() : date.getSeconds();

    let fecha_full = `${anio}-${mes}-${dia} ${hora}:${minutos}:${segundos}`;

    const [
      aprovisionamiento,
      aprovisionamientobsc,
      aseguramiento,
      aseguramientobsc,
    ] = await Promise.all([
      getDataClick("Aprovisionamiento"),
      getDataClick("Aprovisionamiento BSC"),
      getDataClick("Aseguramiento"),
      getDataClick("Aseguramiento BSC"),
    ]);

    let dataaprovisionamiento = [];
    for (const value of aprovisionamiento) {
      dataaprovisionamiento.push(value);
    }
    let aprov;
    if (dataaprovisionamiento) {
      aprov = await generarImagen(
        "Porcentaje tecnicos en sitio Aprovisionamiento",
        "aprovisionamiento",
        dataaprovisionamiento,
        fecha_full
      );
    }

    let dataaprovisionamientobsc = [];
    for (const value of aprovisionamientobsc) {
      dataaprovisionamientobsc.push(value);
    }
    let aprovbsc;
    if (dataaprovisionamientobsc) {
      aprovbsc = await generarImagen(
        "Porcentaje tecnicos en sitio Aprovisionamiento BSC",
        "aprovisionamientobsc",
        dataaprovisionamientobsc,
        fecha_full
      );
    }

    let dataaseguramiento = [];
    for (const value of aseguramiento) {
      dataaseguramiento.push(value);
    }
    let aseg;

    if (dataaseguramiento) {
      aseg = await generarImagen(
        " Porcentaje tecnicos en sitio Aseguramiento",
        "aseguramiento",
        dataaseguramiento,
        fecha_full
      );
    }

    let dataaseguramientobsc = [];
    for (const value of aseguramientobsc) {
      dataaseguramientobsc.push(value);
    }
    let asegbsc;
    if (dataaseguramientobsc) {
      asegbsc = await generarImagen(
        "Porcentaje tecnicos en sitio Aseguramiento BSC",
        "aseguramientobsc",
        dataaseguramientobsc,
        fecha_full
      );
    }

    console.log("dataaprovisionamiento", aprov);
    console.log("dataaprovisionamientobsc", aprovbsc);
    console.log("dataaseguramiento", aseg);
    console.log("dataaseguramientobsc", asegbsc);
  } catch (error) {
    console.log("Error ejecución:", error);
  }
};

module.exports = {
  initReportTecnico,
};
