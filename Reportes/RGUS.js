const XLSX = require("xlsx");
const mssqlDB = require("../database/conn-mssql");
const nodeHtmlToImage = require("node-html-to-image");
const fs = require("fs");
const leerRgus = async (ruta) => {
  const workbook = XLSX.readFile(ruta);
  const workbookSheets = workbook.SheetNames;
  const sheet = workbookSheets[0];
  const dataExcel = XLSX.utils.sheet_to_json(workbook.Sheets[sheet]);
  return dataExcel;
};

const getDataClick = async (variable) => {
  try {
    const result = await mssqlDB.dbConnectionMssql(`
      BEGIN TRAN
      SELECT R.Regiones, SUM(R.RGU) Cumplidos FROM(SELECT TK.CompletionDate,TK.AppointmentStart ,TK.CallID, TK.UNEPedido,TTC.Name Categoria, 
            (CASE 
            WHEN TT.Name LIKE('%Adicion HFC Plus%') THEN 'Nuevo'
            WHEN TT.Name LIKE ('%Nuevo Mono%') THEN 'Nuevo'
            WHEN TT.Name LIKE('%Nuevo Duo%') THEN 'Nuevo' 
            WHEN TT.Name LIKE('%Nuevo Trio%') THEN 'Nuevo' 
            WHEN TT.Name LIKE('%Nuevo Producto%') THEN 'Nuevo' 
            WHEN TT.Name LIKE('%Cambio_Domicilio%') THEN 'Traslado' 
            WHEN TT.Name LIKE('%Cambio_Tecnologia%') THEN 'Migracion'
            WHEN TT.Name LIKE('%Migracion%') THEN 'Migracion'
            WHEN TT.Name LIKE('%Cambio_Equipo%') THEN 'Extension'
            WHEN TT.Name LIKE('%Garantia%') THEN 'Extension'
            WHEN TT.Name LIKE('%Extension%') THEN 'Extension'
            WHEN TT.Name LIKE('%Nuevo Multiple%') THEN 'Nuevo'
            WHEN TT.Name LIKE('%Reparacion Bronce Plus%') THEN 'Reparacion BSC'
            ELSE 'Reparacion' END) 'TIPO_RGU', SS.Name Sistema,(CASE WHEN DIS.Name IN('FACATATIVA_1','FUNZA_1','FUSAGASUGA_1','MOSQUERA_1','FUNZA_MOSQUERA','CHIA_CAJICA') THEN 'Cundinamarca'
            WHEN DIS.Name IN ('SOACHA_1','CAJICA_1','CHIA_1','BOSA','BOSA_1','BOSA_2','BOSA_3','BOSA_4','BOSA_MC','CAN','CAN_1_E6','ECA','ENG','ENG_2','ENGATIVA_MC','FRAGUA_3_MC','FRAGUA_4_MC','FRAGUA_MC','FRG','FRG_1','FRG_2','FRG_3','FRG_4','FRG_5',
            'QCA','QCA_E6','SUBA','SUBA_MC_1','SUBA_MC_2','TIMIZA','TIMIZA_1','TIMIZA_2','TIMIZA_3','TIMIZA_4','TIMIZA_5','TIMIZA_6','TIMIZA_7','TIMIZA_MC_1','TIMIZA_MC_2','TIMIZA_MC_3','TIMIZA_MC_4') THEN 'Bogota' 
            WHEN DIS.Name IN('CAUCASIA_1') THEN 'Antioquia_Edatel'
            WHEN DIS.Name IN ('GUAJIRA_MUNICIPIOS')THEN 'Guajira'
            ELSE RG.Name END)Regiones, DIS.Name Distrito,TK.UNEMunicipio,STA.Name Estado, 
            
            (CASE 
            WHEN TK.UNETecnologias LIKE('%GPON%') THEN 'GPON'
            ELSE 'HFC-COBRE' END) 'TECNOLOGIA', TK.EngineerName, 
            (CASE
            WHEN TT.Name LIKE('%Nuevo Mono%') THEN 1
            WHEN TT.Name LIKE('%Duo%') THEN 2
            WHEN TT.Name LIKE('%Trio%') THEN 3
            WHEN TT.Name LIKE('%Nuevo Producto%') THEN 1
            ELSE 1 END) 'RGU'
            FROM W6TASKS TK
            INNER JOIN W6UNESOURCESYSTEMS SS 	ON TK.UNESourceSystem=SS.W6Key
            INNER JOIN W6REGIONS RG 		ON TK.Region=RG.W6Key
            INNER JOIN W6TASK_STATUSES STA 		ON TK."Status"=STA.W6Key
            INNER JOIN W6DISTRICTS DIS 		ON TK.District=DIS.W6Key
            INNER JOIN W6TASKTYPECATEGORY TTC 	ON TK.TaskTypeCategory=TTC.W6Key
            INNER JOIN W6TASK_TYPES TT 	                                ON TK.TaskType=TT.W6Key
            
            
            WHERE STA.Name IN ('Finalizada')
            AND TTC.Name IN ( 'Aprovisionamiento','Aprovisionamiento BSC','Aseguramiento')
            
            AND TK.CompletionDate > CONVERT (date, SYSDATETIME())
            AND TK.CompletionDate < DATEADD(day, +1, CONVERT (date, SYSDATETIME()))
            
            GROUP BY TK.CompletionDate,TK.CallID, TK.UNEPedido,TTC.Name,TT.Name, SS.Name, RG.Name, DIS.Name, STA.Name,TK.UNETecnologias, TK.EngineerName,TK.AppointmentStart,TK.UNEMunicipio) AS R
            
            WHERE ${variable}
            
            GROUP BY R.Regiones
            ORDER BY R.Regiones ASC
            
            COMMIT TRANSACTION`);

    if (!result) {
      console.log("Error en la conexion de la BD", result);
      return false;
    }

    if (result.recordset.length == 0) {
      console.log("Sin datos para listar", result.recordset.length);
    }

    let res = result.recordset;

    return res;
  } catch (error) {}
};

const getDataAseguramiento = async () => {
  try {
    const result = await mssqlDB.dbConnectionMssql(`
      BEGIN TRAN
      SELECT R.Regiones, COUNT(*) Cumplidos FROM(SELECT TK.CallID, TK.UNEPedido, TTC.Name Categoria, SS.Name Sistema,(CASE WHEN DIS.Name IN('FACATATIVA_1','FUNZA_1','FUSAGASUGA_1','MOSQUERA_1','FUNZA_MOSQUERA','CHIA_CAJICA') THEN 'Cundinamarca'
      WHEN DIS.Name IN ('SOACHA_1','CAJICA_1','CHIA_1','BOSA','BOSA_1','BOSA_2','BOSA_3','BOSA_4','BOSA_MC','CAN','CAN_1_E6','ECA','ENG','ENG_2','ENGATIVA_MC','FRAGUA_3_MC','FRAGUA_4_MC','FRAGUA_MC','FRG','FRG_1','FRG_2','FRG_3','FRG_4','FRG_5',
      'QCA','QCA_E6','SUBA','SUBA_MC_1','SUBA_MC_2','TIMIZA','TIMIZA_1','TIMIZA_2','TIMIZA_3','TIMIZA_4','TIMIZA_5','TIMIZA_6','TIMIZA_7','TIMIZA_MC_1','TIMIZA_MC_2','TIMIZA_MC_3','TIMIZA_MC_4') THEN 'Bogota' 
      WHEN DIS.Name IN('CAUCASIA_1') THEN 'Antioquia_Edatel'
      WHEN DIS.Name IN ('GUAJIRA_MUNICIPIOS')THEN 'Guajira'
      ELSE RG.Name END)Regiones,
      DIS.Name Distrito,TK.UNEMunicipio, STA.Name Estado,TK.UNETecnologias,
      TK.UNEStatusChangeTimeStamp FechaCumplido, SERV_STA.Name EstadoServicio, 
      TK.EngineerID Tecnico, Left(SERV.RespuestaServicio,150) RespuestaServicio,INCOM.Name Incompleto,PINCOM.Name Pendiente,TT.Name As TaskType
      
      FROM W6TASKS TK WITH(NOLOCK)
      INNER JOIN W6UNESOURCESYSTEMS SS ON TK.UNESourceSystem=SS.W6Key
      INNER JOIN W6REGIONS RG ON TK.Region=RG.W6Key
      INNER JOIN W6TASK_STATUSES STA ON TK."Status"=STA.W6Key
      INNER JOIN W6DISTRICTS DIS ON TK.District=DIS.W6Key
      INNER JOIN W6TASKTYPECATEGORY TTC ON TK.TaskTypeCategory=TTC.W6Key
      INNER JOIN W6TASK_TYPES TT ON TK.TaskType=TT.W6Key
      INNER JOIN W6TASKS_UNESERVICES TK_SERV ON TK.W6Key=TK_SERV.W6Key
      INNER JOIN W6UNESERVICES SERV ON TK_SERV.UNEService=SERV.W6Key
      INNER JOIN W6UNESERVICESTATUS SERV_STA ON SERV."Status"=SERV_STA.W6Key
      LEFT JOIN W6UNEINCOMPLETECODES INCOM		ON SERV."IncompleteCode"=INCOM.W6Key
      LEFT JOIN W6UNEPENDINGCODES PINCOM		ON SERV."PendingCode"=PINCOM.W6Key
      
      
      
      
      WHERE  STA.Name IN ('Pendiente','Incompleto')
      AND TTC.Name IN ('Aseguramiento')
      
      AND ( SERV.IncompleteCode IN ('1166033018','1166033085','1166106657','1166106654','1166106653','1166106661','1166106648','1166033232','1166033392','1166033084')
      OR   SERV.PendingCode IN ('1166058309','1166058313','1166058418','1166058425','1166058432','1166172185','1166172186','1166172189','1166172190','1166172193','1166172197','1166058433'))
      
      AND TK.CompletionDate > DATEADD(day, 0, CONVERT (date, SYSDATETIME()))
      AND TK.CompletionDate < DATEADD(day, +1, CONVERT (date, SYSDATETIME()))
      )AS r
      
      GROUP BY R.Regiones
      COMMIT TRANSACTION`);

    if (!result) {
      console.log("Error en la conexion de la BD", result);
      return false;
    }

    if (result.recordset.length == 0) {
      console.log("Sin datos para listar", result.recordset.length);
    }

    let res = result.recordset;

    return res;
  } catch (error) {}
};

const generarImagen = async (titulo, namefile, data, fecha_full) => {
  try {
    const image = fs.readFileSync("./assets/LogoTigoBlanco.png");
    const base64Image = new Buffer.from(image).toString("base64");
    const dataURI = "data:image/jpeg;base64," + base64Image;
    let date = new Date();
    let gethour = date.getHours();
    let porcentaje;
    let totalProyeccion = 0;
    let totalCumplidos = 0;
    let totalPorcentaje = 0;
    let totalNoroocidentePro = 0;
    let totalOrientePro = 0;
    let totalNortePro = 0;
    let totalCentroPro = 0;
    let totalSurPro = 0;
    let totalEjeCafeteroPro = 0;
    let totalNoroocidenteCum = 0;
    let totalOrienteCum = 0;
    let totalNorteCum = 0;
    let totalCentroCum = 0;
    let totalSurCum = 0;
    let totalEjeCafeteroCum = 0;
    let totalSurYEjeCum = 0;
    let totalSurYEjePro = 0;

    let contador = 0;

    data.forEach((val) => {
      totalProyeccion = totalProyeccion + val.Proyeccion;
      totalCumplidos = totalCumplidos + val.Cumplidos;
      if (val.Area == "Noroocidente") {
        totalNoroocidentePro = totalNoroocidentePro + val.Proyeccion;
        totalNoroocidenteCum = totalNoroocidenteCum + val.Cumplidos;
      } else if (val.Area == "Oriente") {
        totalOrientePro = totalOrientePro + val.Proyeccion;
        totalOrienteCum = totalOrienteCum + val.Cumplidos;
      } else if (val.Area == "Norte") {
        totalNortePro = totalNortePro + val.Proyeccion;
        totalNorteCum = totalNorteCum + val.Cumplidos;
      } else if (val.Area == "Centro") {
        totalCentroPro = totalCentroPro + val.Proyeccion;
        totalCentroCum = totalCentroCum + val.Cumplidos;
      } else if (val.Area == "Sur") {
        totalSurPro = totalSurPro + val.Proyeccion;
        totalSurCum = totalSurCum + val.Cumplidos;
      } else if (val.Area == "Eje Cafetero") {
        totalEjeCafeteroPro = totalEjeCafeteroPro + val.Proyeccion;
        totalEjeCafeteroCum = totalEjeCafeteroCum + val.Cumplidos;
      }
    });

    totalNoroocidentePro = Math.round(totalNoroocidentePro);
    totalOrientePro = Math.round(totalOrientePro);
    totalNortePro = Math.round(totalNortePro);
    totalCentroPro = Math.round(totalCentroPro);
    totalSurYEjePro = Math.round(totalSurPro + totalEjeCafeteroPro);
    totalNoroocidenteCum = Math.round(totalNoroocidenteCum);
    totalOrienteCum = Math.round(totalOrienteCum);
    totalNorteCum = Math.round(totalNorteCum);
    totalCentroCum = Math.round(totalCentroCum);
    totalSurYEjeCum = Math.round(totalSurCum + totalEjeCafeteroCum);

    totalPorcentaje = (totalCumplidos / totalProyeccion) * 100;
    if (
      titulo == "HFC-COBRE" ||
      titulo == "GPON" ||
      titulo == "HFC-COBRE & GPON"
    ) {
      switch (gethour) {
        case 10:
          porcentaje = 13;
          break;
        case 11:
          porcentaje = 24;
          break;
        case 12:
          porcentaje = 35;
          break;
        case 13:
          porcentaje = 47;
          break;
        case 14:
          porcentaje = 56;
          break;
        case 15:
          porcentaje = 64;
          break;
        case 16:
          porcentaje = 73;
          break;
        case 17:
          porcentaje = 85;
          break;
        case 18:
          porcentaje = 94;
          break;
        case 19:
          porcentaje = 99;
          break;
        case 20:
          porcentaje = 100;
          break;
        default:
          porcentaje = 0;
      }
    } else if (titulo == "Masivos Hogares") {
      switch (gethour) {
        case 10:
          porcentaje = 22;
          break;
        case 11:
          porcentaje = 33;
          break;
        case 12:
          porcentaje = 43;
          break;
        case 13:
          porcentaje = 53;
          break;
        case 14:
          porcentaje = 62;
          break;
        case 15:
          porcentaje = 71;
          break;
        case 16:
          porcentaje = 81;
          break;
        case 17:
          porcentaje = 90;
          break;
        case 18:
          porcentaje = 96;
          break;
        case 19:
          porcentaje = 99;
          break;
        case 20:
          porcentaje = 100;
          break;
        default:
          porcentaje = 0;
      }
    } else if (titulo == "BSC") {
      switch (gethour) {
        case 10:
          porcentaje = 6;
          break;
        case 11:
          porcentaje = 17;
          break;
        case 12:
          porcentaje = 29;
          break;
        case 13:
          porcentaje = 42;
          break;
        case 14:
          porcentaje = 55;
          break;
        case 15:
          porcentaje = 63;
          break;
        case 16:
          porcentaje = 75;
          break;
        case 17:
          porcentaje = 90;
          break;
        case 18:
          porcentaje = 100;
          break;
        case 19:
          porcentaje = 100;
          break;
        case 20:
          porcentaje = 100;
          break;
        default:
          porcentaje = 0;
      }
    } else if (titulo == "Aseguramiento") {
      switch (gethour) {
        case 10:
          porcentaje = 21;
          break;
        case 11:
          porcentaje = 32;
          break;
        case 12:
          porcentaje = 43;
          break;
        case 13:
          porcentaje = 53;
          break;
        case 14:
          porcentaje = 62;
          break;
        case 15:
          porcentaje = 72;
          break;
        case 16:
          porcentaje = 82;
          break;
        case 17:
          porcentaje = 91;
          break;
        case 18:
          porcentaje = 97;
          break;
        case 19:
          porcentaje = 99;
          break;
        case 20:
          porcentaje = 100;
          break;
        default:
          porcentaje = 0;
      }
    }

    let contentHtml = `<!DOCTYPE html>
        <html>
        <head>
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
        <style>
        body {    
        width: 500px;
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
                <th colspan="5" style="text-align: center; border-right-color: #1F3764 !important;"><img src="{{imageSource}}" style="width: 50px;"></th>
                <th colspan="7" style="text-align: center; vertical-align: bottom; font-size: 20px; border-right-color: #1F3764 !important;">${titulo}</th>
                <th colspan="3" style="text-align: right; vertical-align: bottom; font-size: 8px;">Fecha actualización: ${fecha_full}</th>
            </tr>
            <tr style="background-color:#1F3764; color: white;">
                <th colspan="3" >Region</th>
                <th colspan="2" >Proyección</th>
                <th colspan="2" >∑Proyección</th>
                <th colspan="2" >RGUs Cumplidos</th>
                <th colspan="2" >∑Cumplidos</th>
                <th colspan="2" >% Meta Hora</th>
                <th colspan="2" >% Cumplidos</th>
        </thead>
            <tbody>`;
    data.forEach((val) => {
      contador = contador + 1;
      let n = Math.round((val.Cumplidos / val.Proyeccion) * 100);
      if (n < 1 || n == Infinity || isNaN(n)) {
        n = 0;
      }

      if (n >= porcentaje) {
        color = `#a9d08e`;
      } else {
        color = `#fcb4b2`;
      }
      let pro = Math.round(val.Proyeccion);

      if (val.Area == "Noroocidente") {
        contentHtml += `<tr>
        <td colspan="3"  style="background-color:#dddddd;border: 1px solid #FFFFFF;
        padding: 5px 5px 5px 5px;">${val.Regiones}</td>
        <td colspan="2" class="text-center" style="background-color:#dddddd; border: 1px solid #FFFFFF;
        padding: 5px 5px 5px 5px;">${pro}</td>`;
      }
      if (val.Area == "Oriente") {
        contentHtml += `<tr>
        <td colspan="3"  style="background-color:#FFFFFF;">${val.Regiones}</td>
        <td colspan="2" class="text-center" style="background-color:#FFFFFF;">${pro}</td>`;
      }
      if (val.Area == "Norte") {
        contentHtml += `<tr>
        <td colspan="3"  style="background-color:#dddddd;border: 1px solid #FFFFFF;
        padding: 5px 5px 5px 5px;">${val.Regiones}</td>
        <td colspan="2" class="text-center" style="background-color:#dddddd;border: 1px solid #FFFFFF;
        padding: 5px 5px 5px 5px;">${pro}</td>`;
      }
      if (val.Area == "Centro") {
        contentHtml += `<tr>
        <td colspan="3"  style="background-color:#FFFFFF;">${val.Regiones}</td>
        <td colspan="2" class="text-center" style="background-color:#FFFFFF;">${pro}</td>`;
      }

      if (val.Area == "Sur" || val.Area == "Eje Cafetero") {
        contentHtml += `<tr>
        <td colspan="3"  style="background-color:#dddddd;border: 1px solid #FFFFFF;
        padding: 5px 5px 5px 5px;">${val.Regiones}</td>
        <td colspan="2" class="text-center" style="background-color:#dddddd;border: 1px solid #FFFFFF;
        padding: 5px 5px 5px 5px;">${pro}</td>`;
      }

      if (contador == 1) {
        contentHtml += `
        <td rowspan="5" colspan="2 class="text-center" style="background-color:#dddddd; border: 1px solid #FFFFFF;
        padding: 5px 5px 5px 5px;">${totalNoroocidentePro}</td> `;
      } else if (contador == 6) {
        contentHtml += `
        <td rowspan="5" colspan="2 class="text-center" style="background-color:	#FFFFFF;">${totalOrientePro}</td> `;
      } else if (contador == 11) {
        contentHtml += `
        <td rowspan="11" colspan="2 class="text-center" style="background-color:#dddddd; border: 1px solid #FFFFFF;
        padding: 5px 5px 5px 5px;">${totalNortePro}</td> `;
      } else if (contador == 22) {
        contentHtml += `
        <td rowspan="2" colspan="2 class="text-center" style="background-color:	#FFFFFF;">${totalCentroPro}</td> `;
      } else if (contador == 24) {
        contentHtml += `
        <td rowspan="9" colspan="2 class="text-center" style="background-color:#dddddd; border: 1px solid #FFFFFF;
        padding: 5px 5px 5px 5px;">${totalSurYEjePro}</td> `;
      }

      if (val.Area == "Noroocidente") {
        contentHtml += `
        <td colspan="2" colspan="2 class="text-center" style="background-color:#dddddd; border: 1px solid #FFFFFF;
        padding: 5px 5px 5px 5px;">${val.Cumplidos}</td> `;
      }
      if (val.Area == "Oriente") {
        contentHtml += `
        <td colspan="2" colspan="2 class="text-center" style="background-color:	#FFFFFF;">${val.Cumplidos}</td> `;
      }
      if (val.Area == "Norte") {
        contentHtml += `
        <td colspan="2" colspan="2 class="text-center" style="background-color:#dddddd; border: 1px solid #FFFFFF;
        padding: 5px 5px 5px 5px;">${val.Cumplidos}</td> `;
      }
      if (val.Area == "Centro") {
        contentHtml += `
        <td colspan="2" colspan="2 class="text-center" style="background-color:	#FFFFFF;">${val.Cumplidos}</td> `;
      }

      if (val.Area == "Sur" || val.Area == "Eje Cafetero") {
        contentHtml += `
        <td colspan="2" colspan="2 class="text-center" style="background-color:#dddddd; border: 1px solid #FFFFFF;
        padding: 5px 5px 5px 5px;">${val.Cumplidos}</td> `;
      }

      if (contador == 1) {
        contentHtml += `
        <td rowspan="5" colspan="2 class="text-center" style="background-color:#dddddd;border: 1px solid #FFFFFF;
        padding: 5px 5px 5px 5px;">${totalNoroocidenteCum}</td> `;
      } else if (contador == 6) {
        contentHtml += `
        <td rowspan="5" colspan="2 class="text-center" style="background-color:	#FFFFFF;">${totalOrienteCum}</td> `;
      } else if (contador == 11) {
        contentHtml += `
        <td rowspan="11" colspan="2 class="text-center" style="background-color:#dddddd;border: 1px solid #FFFFFF;
        padding: 5px 5px 5px 5px;">${totalNorteCum}</td> `;
      } else if (contador == 22) {
        contentHtml += `
        <td rowspan="2" colspan="2 class="text-center" style="background-color:	#FFFFFF;">${totalCentroCum}</td> `;
      } else if (contador == 24) {
        contentHtml += `
        <td rowspan="9" colspan="2 class="text-center" style="background-color:#dddddd;border: 1px solid #FFFFFF;
        padding: 5px 5px 5px 5px;">${totalSurYEjeCum}</td> `;
      }

      if (val.Area == "Noroocidente") {
        contentHtml += `
        <td colspan="2" colspan="2 class="text-center" style="background-color:#dddddd;border: 1px solid #FFFFFF;
        padding: 5px 5px 5px 5px;">${porcentaje}%</td> `;
      }
      if (val.Area == "Oriente") {
        contentHtml += `
        <td colspan="2" colspan="2 class="text-center" style="background-color:	#FFFFFF;">${porcentaje}%</td> `;
      }
      if (val.Area == "Norte") {
        contentHtml += `
        <td colspan="2" colspan="2 class="text-center" style="background-color:#dddddd; border: 1px solid #FFFFFF;
        padding: 5px 5px 5px 5px;">${porcentaje}%</td> `;
      }
      if (val.Area == "Centro") {
        contentHtml += `
        <td colspan="2" colspan="2 class="text-center" style="background-color:	#FFFFFF;">${porcentaje}%</td> `;
      }

      if (val.Area == "Sur" || val.Area == "Eje Cafetero") {
        contentHtml += `
        <td colspan="2" colspan="2 class="text-center" style="background-color:#dddddd; border: 1px solid #FFFFFF;
        padding: 5px 5px 5px 5px;">${porcentaje}%</td> `;
      }

      contentHtml += `
                    <td colspan="2" class="text-center" style="background-color: ${color}">${n}%</td> 
                    </tr>`;
    });

    contentHtml += `
        </tbody>

        <tfoot >
        <tr >
        <td colspan=3 style="background-color:#1F3764; color: white;">Total</td>
        <td colspan=2 class="text-center" style="background-color:#1F3764; color: white;" >${Math.round(
          totalProyeccion
        )}</td>
        <td colspan=2 class="text-center" style="background-color:#1F3764; color: white;" >${Math.round(
          totalProyeccion
        )}</td>
        <td colspan=2 class="text-center" style="background-color:#1F3764; color: white;">${Math.round(
          totalCumplidos
        )}</td>
        <td colspan=2 class="text-center" style="background-color:#1F3764; color: white;">${Math.round(
          totalCumplidos
        )}</td>
        <td colspan=2 class="text-center" style="background-color:#1F3764; color: white;">${Math.round(
          porcentaje
        )}%</td>`;
    let n = Math.round(totalPorcentaje);
    if (n < 1 || n == Infinity || isNaN(n)) {
      n = 0;
    }

    if (n >= porcentaje) {
      color = `#a9d08e`;
    } else {
      color = `#fcb4b2`;
    }
    contentHtml += `<td colspan=2 class="text-center" style="background-color: ${color}">${n}%</td>
        </tr>
    </tfoot>
        </table>
        
        </body>
        </html>`;

    let res = await nodeHtmlToImage({
      output: `./images/RGUs/${namefile}.png`,
      html: contentHtml,
      content: { imageSource: dataURI },
    }).then(() => `Imagen de ${namefile} creado con exito`);

    return res;
  } catch (error) {
    console.log(error);
  }
};

const HFCCOBRE = async () => {
  let Rgus = await leerRgus("./Reportes/Base_Carga_RGUs.xlsx");
  let variable =
    "R.Categoria = 'Aprovisionamiento' AND R.TECNOLOGIA = 'HFC-COBRE' AND R.TIPO_RGU = 'Nuevo'";
  let dataConsulta = await getDataClick(variable);

  var posibles = [];

  Rgus.forEach((reg) => {
    const existeEnBD = dataConsulta.find(
      (item) => item.Regiones === reg.Regiones
    );
    const agregarEnArrayRevisar = {
      Area: reg.Area,
      Regiones: reg.Regiones,
      Proyeccion: reg["HFC-COBRE"],
    };
    if (existeEnBD) {
      agregarEnArrayRevisar.Cumplidos = existeEnBD.Cumplidos;
    } else {
      agregarEnArrayRevisar.Cumplidos = 0;
    }
    posibles.push(agregarEnArrayRevisar);
  });
  return posibles;
};
const GPON = async () => {
  let Rgus = await leerRgus("./Reportes/Base_Carga_RGUs.xlsx");
  let variable =
    "R.Categoria = 'Aprovisionamiento' AND R.TECNOLOGIA = 'GPON' AND R.TIPO_RGU = 'Nuevo'";
  let dataConsulta = await getDataClick(variable);

  var posibles = [];

  Rgus.forEach((reg) => {
    const existeEnBD = dataConsulta.find(
      (item) => item.Regiones === reg.Regiones
    );
    const agregarEnArrayRevisar = {
      Area: reg.Area,
      Regiones: reg.Regiones,
      Proyeccion: reg["GPON"],
    };
    if (existeEnBD) {
      agregarEnArrayRevisar.Cumplidos = existeEnBD.Cumplidos;
    } else {
      agregarEnArrayRevisar.Cumplidos = 0;
    }
    posibles.push(agregarEnArrayRevisar);
  });
  return posibles;
};
const HFCCOBRE_GPON = async () => {
  let Rgus = await leerRgus("./Reportes/Base_Carga_RGUs.xlsx");
  let variable = "R.Categoria = 'Aprovisionamiento' AND R.TIPO_RGU = 'Nuevo'";
  let dataConsulta = await getDataClick(variable);

  var posibles = [];

  Rgus.forEach((reg) => {
    const existeEnBD = dataConsulta.find(
      (item) => item.Regiones === reg.Regiones
    );
    const agregarEnArrayRevisar = {
      Area: reg.Area,
      Regiones: reg.Regiones,
      Proyeccion: reg["HFC-COBRE & GPON"],
    };
    if (existeEnBD) {
      agregarEnArrayRevisar.Cumplidos = existeEnBD.Cumplidos;
    } else {
      agregarEnArrayRevisar.Cumplidos = 0;
    }
    posibles.push(agregarEnArrayRevisar);
  });
  return posibles;
};
const MasivosHogares = async () => {
  let Rgus = await leerRgus("./Reportes/Base_Carga_RGUs.xlsx");
  let variable =
    "R.Categoria = 'Aprovisionamiento' AND R.TIPO_RGU IN ('Extension','Traslado','Migracion')";
  let dataConsulta = await getDataClick(variable);

  var posibles = [];

  Rgus.forEach((reg) => {
    const existeEnBD = dataConsulta.find(
      (item) => item.Regiones === reg.Regiones
    );
    const agregarEnArrayRevisar = {
      Area: reg.Area,
      Regiones: reg.Regiones,
      Proyeccion: reg["Masivos Hogares"],
    };
    if (existeEnBD) {
      agregarEnArrayRevisar.Cumplidos = existeEnBD.Cumplidos;
    } else {
      agregarEnArrayRevisar.Cumplidos = 0;
    }
    posibles.push(agregarEnArrayRevisar);
  });
  return posibles;
};
const BSC = async () => {
  let Rgus = await leerRgus("./Reportes/Base_Carga_RGUs.xlsx");
  let variable =
    "R.Categoria = 'Aprovisionamiento BSC' AND R.TIPO_RGU = 'Nuevo'";
  let dataConsulta = await getDataClick(variable);

  var posibles = [];

  Rgus.forEach((reg) => {
    const existeEnBD = dataConsulta.find(
      (item) => item.Regiones === reg.Regiones
    );
    const agregarEnArrayRevisar = {
      Area: reg.Area,
      Regiones: reg.Regiones,
      Proyeccion: reg["BSC"],
    };
    if (existeEnBD) {
      agregarEnArrayRevisar.Cumplidos = existeEnBD.Cumplidos;
    } else {
      agregarEnArrayRevisar.Cumplidos = 0;
    }
    posibles.push(agregarEnArrayRevisar);
  });
  return posibles;
};
const Aseguramiento = async () => {
  let Rgus = await leerRgus("./Reportes/Base_Carga_RGUs.xlsx");
  let variable = "R.Categoria = 'Aseguramiento' AND  R.TIPO_RGU = 'Reparacion'";
  let dataConsulta = await getDataClick(variable);
  let dataEscalado = await getDataAseguramiento();

  var posibles = [];
  var posiblesfinales = [];

  Rgus.forEach((reg) => {
    const existeEnBD = dataConsulta.find(
      (item) => item.Regiones === reg.Regiones
    );
    const agregarEnArrayRevisar = {
      Area: reg.Area,
      Regiones: reg.Regiones,
      Proyeccion: reg["Aseguramiento"],
    };
    if (existeEnBD) {
      agregarEnArrayRevisar.Cumplidos = existeEnBD.Cumplidos;
    } else {
      agregarEnArrayRevisar.Cumplidos = 0;
    }
    posibles.push(agregarEnArrayRevisar);
  });

  posibles.forEach((reg) => {
    const existeEnBD = dataEscalado.find(
      (item) => item.Regiones === reg.Regiones
    );
    const agregarEnArrayRevisar = {
      Area: reg.Area,
      Regiones: reg.Regiones,
      Proyeccion: reg.Proyeccion,
      Cumplidos: reg.Cumplidos,
    };

    if (existeEnBD) {
      agregarEnArrayRevisar.Cumplidos += existeEnBD.Cumplidos;
    }
    posiblesfinales.push(agregarEnArrayRevisar);
  });

  return posiblesfinales;
};

const initRgus = async () => {
  try {
    const [hfccobre, gpon, hfcgpon, masivos, bsc, aseguramiento] =
      await Promise.all([
        HFCCOBRE(),
        GPON(),
        HFCCOBRE_GPON(),
        MasivosHogares(),
        BSC(),
        Aseguramiento(),
      ]);
    // let hfccobre = await HFCCOBRE();
    // let gpon = await GPON();
    // let hfcgpon = await HFCCOBRE_GPON();
    // let masivos = await MasivosHogares();
    // let bsc = await BSC();
    // let aseguramiento = await Aseguramiento();
    let date = new Date();
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

    let hfci = await generarImagen(
      "HFC-COBRE",
      "HFC-COBRE",
      hfccobre,
      fecha_full
    );
    let gponI = await generarImagen("GPON", "GPON", gpon, fecha_full);
    let hfcgponi = await generarImagen(
      "HFC-COBRE & GPON",
      "HFC-COBRE & GPON",
      hfcgpon,
      fecha_full
    );
    let masivosi = await generarImagen(
      "Masivos Hogares",
      "Masivos Hogares",
      masivos,
      fecha_full
    );
    let bsci = await generarImagen("BSC", "BSC", bsc, fecha_full);
    let aseguramientoi = await generarImagen(
      "Aseguramiento",
      "Aseguramiento",
      aseguramiento,
      fecha_full
    );

    console.log("Data HFC", hfci);
    console.log("Data GPON", gponI);
    console.log("Data HFC-GPON", hfcgponi);
    console.log("Data Masivos", masivosi);
    console.log("Data BSC", bsci);
    console.log("Data Aseguramiento", aseguramientoi);
  } catch (error) {
    console.log("Error ejecución:", error);
  }
};

module.exports = {
  initRgus,
};
