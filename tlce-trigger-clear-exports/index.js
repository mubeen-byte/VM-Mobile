const async = require("async");
const ENV = process.env.ENV;
const REGION = process.env.AWS_REGION;
const DB_HOST = process.env.DB_HOST;
const DB_USER = process.env.DB_USER;
const dao = require('foundry-dao/dynamodb')(REGION, `_${ENV}`);
const EXPORT_STATUS_TYPE_ID = process.env.EXPORT_STATUS_TYPE_ID;
let rdsDao;

exports.handler = (event, context, callback) => {
    async.waterfall(
        [
            initRDS,
            getExportCSVList,
            deleteOldObjects
        ],
        (err) => {
            if (err) {
                console.error(err);
            }
            return context.succeed("success");
        }
    );

    function initRDS(cbInitRDS) {
        rdsDao = require("foundry-dao/rds")({
            host: DB_HOST,
            user: DB_USER,
            port: 5432,
            database: 'AdvancedSearch'
        }, REGION, ENV);

        rdsDao.connect((err) => {
            if (err) {
                return cbInitRDS(err);
            }

            return cbInitRDS(null);
        });
    }

    function getExportCSVList(cbGetExportCSVList) {
        const query = `SELECT * FROM "_type_${EXPORT_STATUS_TYPE_ID}" WHERE "Created" < ${new Date().getTime() - 24*60*60*1000}`;
        console.log(query);
        rdsDao.query(query, (err, data) => {
            rdsDao.closeConnection();
            if (err) {
                console.log(err);
                return cbGetExportCSVList("Error while query database");
            }
            if (!data?.length) {
                return cbGetExportCSVList("no data");
            }
            console.log("Data to delete:", JSON.stringify(data));
            return cbGetExportCSVList(null, data);
        });
    }

    function deleteOldObjects(data, cbDeleteOldObjects) {
        async.eachSeries(data, (obj, next) => {
            dao.deviceInfo.Remove(obj.deviceId, (err, res) => {
                if (err) {
                    console.error(`Error deleting object ${obj.deviceId}:`, err);
                }
                next();
            });
        }, cbDeleteOldObjects);
    }
};


