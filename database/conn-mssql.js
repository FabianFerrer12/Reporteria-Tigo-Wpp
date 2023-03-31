
const sql = require('mssql')

const config = {
    user: "BI_Clicksoftware",
    password: "6n`Vue8yYK7Os4D-y",
    //server: "NETV-PSQL09-05",
    server: "10.100.74.16",
    database: "Service Optimization",
    connectionTimeout: 30000,
    requestTimeout: 30000,
    pool: {
        max: 10,
        min: 0,
        idleTimeoutMillis: 30000
    },
    options: {
        encrypt: false,
        trustServerCertificate: true
    }
}


//instantiate a connection pool
const dbConnectionMssql = async (query, params = null) => {
    console.log("Intentando conectar");
    let result = null;   
    try{
        const pool = await sql.connect(config);
        if(params){
            result = await pool.request().query(query, params);
        }
        else{
            result = await pool.request().query(query);
        }
        await pool.close();
        return result;
    }
    catch(err) {
        console.log(err.name); // ReferenceError
        console.log(err.message);
        //console.log(err.stack);
        //console.log('ERROR CONEXION: ', err);
        return false;
    }
}

module.exports = {
    dbConnectionMssql
}
