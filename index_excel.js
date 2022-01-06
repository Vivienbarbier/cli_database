const { initializeApp, applicationDefault, cert } = require('firebase-admin/app');
const { getFirestore, Timestamp, FieldValue } = require('firebase-admin/firestore')
require('dotenv').config();
const serviceAccount = require('./serviceAccountKey.json');

initializeApp({
  credential: cert(serviceAccount)
});

var db = getFirestore();

const admin = require('firebase-admin');
const headersMapping = require("./headersMapping.json");
const collectionKey = "case"; //name of the collection
const business_unit_code = "OCPL"
const business_unit_name = "Centre Pays de Loire"


var XLSX = require('xlsx')
var workbook = XLSX.readFile("./data.xlsx");
var sheet_name_list = workbook.SheetNames;
var worksheet = workbook.Sheets[sheet_name_list[0]];


var col="";
for (let i = 0; i < 80; i++) {
    col = ((i>=26) ? String.fromCharCode(65+(parseInt(i/26)-1)) : "") + String.fromCharCode(65+i%26)+"1";
    if (worksheet[col] !== undefined){
        if (headersMapping[worksheet[col].w] !== undefined){
            worksheet[col].w = headersMapping[worksheet[col].w]
        }
    }    
}
var data = XLSX.utils.sheet_to_json(worksheet);
loadDataInToFireStore();


async function loadDataInToFireStore() {
    var teaser = Object({ list : [] })      
    for ( docKey in data) {
        var doc = data[docKey];
        
        if (doc.initial_revenu === undefined)               doc.initial_revenu              = 0; 
        if (doc.initial_marging  === undefined)             doc.initial_marging             = 0;
        if (doc.budget_eng_cost === undefined)              doc.budget_eng_cost             = 0;
        if (doc.budget_field_hours === undefined)           doc.budget_field_hours          = 0;
        if (doc.budget_workforce_cost === undefined)        doc.budget_workforce_cost       = 0;
        if (doc.budget_material_cost === undefined)         doc.budget_material_cost        = 0;
        if (doc.budget_subcontracting_cost === undefined)   doc.budget_subcontracting_cost  = 0;
        if (doc.budget_total_cost === undefined)            doc.budget_total_cost           = 0;
        
        if (doc.conso_field_hours === undefined)           doc.conso_field_hours          = 0;
        if (doc.conso_workforce_cost === undefined)        doc.conso_workforce_cost       = 0;
        if (doc.conso_material_cost === undefined)         doc.conso_material_cost        = 0;
        if (doc.conso_subcontracting_cost === undefined)   doc.conso_subcontracting_cost  = 0;
        if (doc.conso_total_cost === undefined)            doc.conso_total_cost           = 0;
       
        if (doc.prev_field_hours === undefined)           doc.prev_field_hours          = 0;
        if (doc.prev_workforce_cost === undefined)        doc.prev_workforce_cost       = 0;
        if (doc.prev_material_cost === undefined)         doc.prev_material_cost        = 0;
        if (doc.prev_subcontracting_cost === undefined)   doc.prev_subcontracting_cost  = 0;
        if (doc.prev_total_cost === undefined)            doc.prev_total_cost           = 0;
        for ( k in headersMapping) {
            if (doc[headersMapping[k]] === undefined) doc[headersMapping[k]] = "-";
        }


        var record = Object({
            work_order : doc.work_order_number,
            imputation_code : doc.imputation_code,
            progress : parseInt(doc.progress*100),
            type : doc.type,
            comments : "",
            status : (doc.progress == 100) ? "closed" : ((doc.progress == 0) ? "planify" : "wip"),
            owners: {
                business : doc.business_owner,
                execution : doc.execution_owner,
                field : doc.field_owner,
                business_unit: business_unit_name
            },

            contract : {
                signing_date : ((doc.signing_date !== null)  ? ExcelDateToDate(doc.signing_date) : null),
                client : doc.client,
                name : doc.name,
                revenu: doc.initial_revenu
            },
            schedule : {
                start_date_ts :  ((doc.start_date !== null)  ? ExcelDateToJSTimestamp(doc.start_date) : null),
                start_date :  ((doc.start_date !== null)  ? ExcelDateToDate(doc.start_date) : null),
                end_date_ts :  ((doc.end_date !== null)  ? ExcelDateToJSTimestamp(doc.end_date) : null),
                end_date :  ((doc.start_date !== null)  ? ExcelDateToDate(doc.end_date) : null),
            },
            budget : {
                workloads : [
                    { 
                        type : 1,
                        workload : doc.budget_eng_cost/100, // TO fix
                        rate : 100
                    },
                    { 
                        type : 2,
                        workload : doc.budget_field_hours,
                        rate : (doc.budget_field_hours == 0) ? 0 : (doc.budget_workforce_cost - doc.budget_eng_cost) / (doc.budget_field_hours)                   },
                ],
                workforce : doc.budget_workforce_cost,
                material: doc.budget_material_cost,
                subcontracting :  doc.budget_subcontracting_cost,
                total :  doc.budget_total_cost
            },
            conso : {
                workloads : [
                    { 
                        type : 1,
                        workload : doc.conso_eng_cost/100, // TO fix
                        rate : 100
                    },
                    { 
                        type : 2,
                        workload : doc.conso_field_hours,
                        rate : (doc.conso_field_hours == 0) ? 0 : (doc.conso_workforce_cost - doc.conso_eng_cost) / (doc.conso_field_hours)
                    },                
                ],
                workforce : doc.conso_workforce_cost,
                material : doc.conso_material_cost,
                subcontracting :  doc.conso_subcontracting_cost,
                total :  doc.conso_total_cost
            },
            prev : {
                workloads : [
                    { 
                        type : 1,
                        workload : doc.prev_eng_cost/100, // TO fix
                        rate : 100
                    },
                    { 
                        type : 2,
                        workload : doc.prev_field_hours,
                        rate : (doc.prev_field_hours == 0) ? 0 : (doc.prev_workforce_cost - doc.prev_eng_cost) / (doc.prev_field_hours)
                    },                
                ],
                workforce : doc.prev_workforce_cost,
                material : doc.prev_material_cost,
                subcontracting :  doc.prev_subcontracting_cost,
                total :  doc.prev_total_cost,
                revenu : doc.prev_revenu,
                marging : doc.prev_marging,
                risk_amount : doc.risk_amount,
                risk_description : doc.risk_description,
            },
            provision : {
                    amount : doc.provision_amount,
                    description : doc.provision_description
            },
            ytd : {
                workforce : doc.YTD_field_cost,
                material : doc.YTD_material_cost,
                subcontracting : doc.YTD_subcontracting_cost,
                total : doc.YTD_total_cost,
                stock : doc.YTD_stock,
                marging: doc.YTD_marging
            },
            situation : {
                activity : doc.produit_revenu,
                billed_amount : doc.billed_amount,
                fae : doc.FAE,
                received_amount : doc.received_amount,
                last_years_revenu : doc.last_years_revenu,
                producted_current_year_revenu : doc.producted_current_year_revenu,
                remaining_current_year_revenu : doc.remaining_current_year_revenu,
                projection_next_years_revenu  : doc.projection_next_years_revenu,
                PAR : doc.PAR,
            }
        });

           const refDoc = await db.collection("case").doc("OCPL").collection("current").add(record).catch((error) => {
               console.error("Error writing document: ", error);
           });
           console.log("Document " + refDoc.id + " successfully written!");
           teaser.list.push({
               work_order_number : record.work_order,
               imputation_code   : record.imputation_code,
               owner : record.owners.execution,
               name : record.contract.name,
               client: record.contract.client,
               start_date: record.schedule.start_date,
               revenu : record.contract.revenu,
               progress : parseInt(record.progress*100),
               status : record.status,
               id : refDoc.id,
           });
        }
    db.collection("case").doc(business_unit_code).set(teaser).then((res) => {
        console.log("Teaser Created");
    }).catch((error) => {
        console.error("Error writing document: ", error);
    });   
}

function ExcelDateToJSTimestamp(serial) {
    if (serial === "a completer") return null;  
    var utc_days  = Math.floor(serial - 25569);
    var utc_value = utc_days * 86400;                                        
    var date_info = new Date(utc_value * 1000);
    return date_info.getTime();
 }

 function ExcelDateToDate(serial) {
    if (serial === "a completer") return "-";
    var utc_days  = Math.floor(serial - 25569);
    var utc_value = utc_days * 86400;                                        
    var date_info = new Date(utc_value * 1000);
    return date_info.getFullYear() + "-" +(date_info.getMonth()<9 ? "0" : "") + (date_info.getMonth()+1) + "-" + (date_info.getDate()<10 ? "0" : "") + date_info.getDate() ;
  
}