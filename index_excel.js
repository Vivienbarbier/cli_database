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
    var teaser = Object({ list : [], records :[], previsions : [] })
    db.collection("case").doc(business_unit_code).set(teaser).then((res) => {
        console.log("Empty Teaser creation : \tOK");
    }).catch((error) => {
        console.error("Empty Teaser creation : \tError ", error);
    }); 
    
    

    for (docKey in data) {
        var doc = data[docKey];
        
        if (doc.initial_revenu === undefined)               doc.initial_revenu              = 0; 
        if (doc.initial_margin  === undefined)              doc.initial_margin             = 0;
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

        var margin_rate = (doc.initial_revenu === 0) ? 0 : doc.initial_margin / doc.initial_revenu;

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
                margin : doc.prev_margin,
                risk_amount : doc.risk_amount,
                risk_description : doc.risk_description,
            },
            provision : {
                    amount : doc.provision_amount,
                    description : doc.provision_description
            },
            ytd : {
                workforce :         doc.YTD_field_cost,
                material :          doc.YTD_material_cost,
                subcontracting :    doc.YTD_subcontracting_cost,
                total :             doc.YTD_total_cost,
                stock :             doc.YTD_stock,
                margin:             doc.YTD_margin,
                revenu :            doc.YTD_total_cost + doc.YTD_stock + doc.YTD_margin,
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
            },
            records : [],
            previsions : []
        });
        
        record.records = CreateRecords(
            doc.initial_revenu,
            doc.initial_margin,
            record.budget.workforce,
            record.budget.material,
            record.budget.subcontracting,
            record.schedule.start_date,
            record.schedule.end_date,
            teaser.records);
            
        record.previsions = CreatePrevisions(
            doc.initial_revenu,
            doc.initial_margin,
            record.budget.workforce,
            record.budget.material,
            record.budget.subcontracting,
            record.schedule.start_date,
            record.schedule.end_date,
            teaser.previsions);
        

        const refDoc = await db.collection("case").doc("OCPL").collection("current").add(record).catch((error) => {
            console.error("Error writing document: ", error);
        });
        console.log("Document " + refDoc.id + " successfully written!");         
    }
    db.collection("case").doc(business_unit_code).update( {
        records : teaser.records.sort((a,b) => {
            return a.ts - b.ts; 
        }), 
        previsions : teaser.previsions.sort((a,b) => {
            return a.ts - b.ts; 
        }), 
    } ).catch((error) => {
        console.error("Error writing teaser document: ", error);
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
    return date_info.getFullYear() + "-" +(date_info.getMonth()<9 ? "0" : "") + (date_info.getMonth()+1) + "-" + (date_info.getDate()<10 ? "0" : "") + date_info.getDate();
}

function getWeekNumber(date){
    var oneJan = new Date(date.getFullYear(),0,1);
    var numberOfDays = Math.floor((date - oneJan) / (24 * 60 * 60 * 1000));
    var result = Math.ceil(( date.getDay() + 1 + numberOfDays) / 7);
    return result;    
}


function CreateRecords(revenu, margin, workforce, material, subcontracting, start_date, end_date, teaser ){
    var records = [];
    let start = new Date(start_date);
    let end = new Date(end_date);
    let duration_weeks = (end.getTime() - start.getTime()) / (7000 * 3600 * 24);
    var date = start;
    let today = new Date();

    for (var i =0; i < duration_weeks ; i++){    
        if(date >= today) break;
        let dateStr = date.toISOString().slice(0, 10)
        const record = {
            date :  dateStr , 
            ts :    date.getTime(),
            revenu : (revenu / duration_weeks),
            margin : (margin / duration_weeks),
            workforce : (workforce / duration_weeks),
            material : (material / duration_weeks),
            subcontracting : (subcontracting / duration_weeks)
        }
        records.push(record)
        const recIndex = teaser.findIndex( r => r.date === dateStr);
        if(recIndex === -1){
            teaser.push(record); 
        }else{
            teaser[recIndex].revenu          += record.revenu;
            teaser[recIndex].margin          += record.margin;
            teaser[recIndex].workforce       += record.workforce;
            teaser[recIndex].material        += record.material;
            teaser[recIndex].subcontracting  += record.subcontracting;
        }
        date.setDate(date.getDate()+7);
    }
    return records;
}

function CreatePrevisions(revenu, margin, workforce, material, subcontracting, start_date, end_date,teaser){
    var previsions = [];
    let start = new Date(start_date);
    let end = new Date(end_date);
    let duration_weeks = (end.getTime() - start.getTime()) / (7000 * 3600 * 24);
    var key ="";
    var date = new Date(); // Today
    for (var i =0; i < duration_weeks ; i++){    
        if(date >= end) break;
        let dateStr = date.toISOString().slice(0, 10)
        const prevision = {
            date :  dateStr ,
            ts: date.getTime(),
            revenu : (revenu / duration_weeks),
            margin : (margin / duration_weeks),
            workforce : (workforce / duration_weeks),
            material : (material / duration_weeks),
            subcontracting : (subcontracting / duration_weeks)
        };
        previsions.push(prevision);
        const recIndex = teaser.findIndex( r => r.date === dateStr);
        if(recIndex === -1){
            teaser.push(prevision); 
        }else{
            teaser[recIndex].revenu          += prevision.revenu;
            teaser[recIndex].margin          += prevision.margin;
            teaser[recIndex].workforce       += prevision.workforce;
            teaser[recIndex].material        += prevision.material;
            teaser[recIndex].subcontracting  += prevision.subcontracting;
        }
        date.setDate(date.getDate()+7);
    }
    return previsions;
}



//function CreateHistoryRecords(value, start_date, end_date){
//    var records = [];
//    let start = new Date(start_date);
//    let end = new Date(end_date);
//    let duration_weeks = (end.getTime() - start.getTime()) / (7000 * 3600 * 24);
//    var key ="";
//    var date = start;
//    let today = new Date();
//    for (var i =0; i < duration_weeks ; i++){    
//        if(date >= today) break;
//        records.push({date :  date.toISOString().slice(0, 10) , v: (value / duration_weeks) })
//        date.setDate(date.getDate()+7);
//    }
//    return records;
//}
//
//function CreatePrevisions(value, start_date, end_date){
//    var previsions = [];
//    let start = new Date(start_date);
//    let end = new Date(end_date);
//    let duration_weeks = (end.getTime() - start.getTime()) / (7000 * 3600 * 24);
//    var key ="";
//    var date = new Date(); // Today
//    for (var i =0; i < duration_weeks ; i++){    
//        if(date >= end) break;
//        previsions.push({date : date.toISOString().slice(0, 10) , v: (value / duration_weeks) })
//        date.setDate(date.getDate()+7);
//    }
//    return previsions;
//}
