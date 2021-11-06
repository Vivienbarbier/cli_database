const { initializeApp, applicationDefault, cert } = require('firebase-admin/app');
const { getFirestore, Timestamp, FieldValue } = require('firebase-admin/firestore');
const serviceAccount = require('./serviceAccountKey.json');

initializeApp({
  credential: cert(serviceAccount)
});

const firestore = getFirestore();
const admin = require('./node_modules/firebase-admin');
const data = require("./data2.json");
const collectionKey = "case"; //name of the collection

if (data && (typeof data === "object")) {

    Object.keys(data).forEach(docKey => {
        var doc = data[docKey];
        // Check values
       doc.imputation_code_generic_flag = (doc.imputation_code_generic_flag == "Specifique") ? false : true; 
       if (doc.type == "")                          doc.type                        = null;
       if (doc.signing_date == "")                  doc.signing_date                = null;
       if (doc.client == "")                        doc.client                      = null ;
       if (doc.name == "")                          doc.name                        = null;
       if (doc.initial_revenu == "")                doc.initial_revenu              = null;
       if (doc.start_date == "")                    doc.start_date                  = null;
       if (doc.duration == "")                      doc.duration                    = null;
       if (doc.end_date == "")                      doc.end_date                    = null;
       if (doc.preparation_status == "")            doc.preparation_status          = null;
       if (doc.field_staff_std_rate == "")          doc.field_staff_std_rate        = null
       if (doc.budget_workforce_total_hours == "")  doc.budget_workforce_total_hours = null;
       if (doc.budget_workforce_hours_type1 == "")  doc.budget_workforce_hours_type1 = null;
       if (doc.budget_workforce_hours_type2 == "")  doc.budget_workforce_hours_type2 = null;
       if (doc.budget_workforce_hours_type3 == "")  doc.budget_workforce_hours_type3 = null;
       if (doc.budget_workforce_hours_type4 == "")  doc.budget_workforce_hours_type4 = null;
       if (doc.budget_workforce_total_cost == "")   doc.budget_workforce_total_cost   = null;
       if (doc.budget_material_cost == "")          doc.budget_material_cost         = null;
       if (doc.budget_subcontracting_cost == "")    doc.budget_subcontracting_cost   = null;
       if (doc.budget_total_cost == "")             doc.budget_total_cost_1         = null;
       if (doc.conso_workforce_total_cost == "")    doc.conso_workforce_total_cost  = null;
       if (doc.conso_workforce_total_hours == "")   doc.conso_workforce_total_hours = null;
       if (doc.conso_material_cost == "")           doc.conso_material_cost         = null;
       if (doc.conso_subcontracting_cost == "")     doc.conso_subcontracting_cost   = null;
       if (doc.conso_total_cost == "")              doc.conso_total_cost            = null;
       if (doc.progress == "")                      doc.progress                    = null;
       if (doc.status == "")                        doc.status                      = null;
       if (doc.risk_amount == "")                   doc.risk_amount                 = null;
       if (doc.prev_workforce_total_cost == "")     doc.prev_workforce_total_cost   = null;
       if (doc.prev_workforce_total_hours == "")    doc.prev_workforce_total_hours  = null;
       if (doc.prev_material_cost == "")            doc.prev_material_cost          = null;
       if (doc.prev_subcontracting_cost == "")      doc.prev_subcontracting_cost    = null;
       if (doc.prev_marging = "")                   doc.prev_marging                = null;
       if (doc.prev_marging_ratio == "")            doc.prev_marging_ratio          = null;
       if (doc.fae == "")                           doc.fae                         = null;
       if (doc.billed_amout == "")                  doc.billed_amout                = null;
       if (doc.recoverable_amount == "")            doc.recoverable_amount          = null;
       if (doc.paid_amount == "")                   doc.paid_amount                 = null;
       if (doc.realised_amount == "")               doc.realised_amount             = null;
       if (doc.prev_revenu_current_year == "")      doc.prev_revenu_current_year    = null;
     

        //Convert to floats
       if (typeof doc.prev_revenu_current_year     === 'string')    doc.prev_revenu_current_year = parseFloat(doc.prev_revenu_current_year.replace(",","."));
       if (typeof doc.prev_total_cost              === 'string') doc.prev_total_cost                  = parseFloat(doc.prev_total_cost.replace(",","."));
       if (typeof doc.initial_revenu               === 'string') doc.initial_revenu                   = parseFloat(doc.initial_revenu.replace(",","."));
       if (typeof doc.duration                     === 'string') doc.duration                         = parseFloat(doc.duration.replace(",","."));
       if (typeof doc.field_staff_std_rate         === 'string') doc.field_staff_std_rate             = parseFloat(doc.field_staff_std_rate.replace(",","."));       
       if (typeof doc.budget_workforce_total_hours === 'string') doc.budget_workforce_total_hours     = parseFloat(doc.budget_workforce_total_hours.replace(",","."));
       if (typeof doc.budget_workforce_hours_type1 === 'string') doc.budget_workforce_hours_type1     = parseFloat(doc.budget_workforce_hours_type1.replace(",","."));
       if (typeof doc.budget_workforce_hours_type2 === 'string') doc.budget_workforce_hours_type2     = parseFloat(doc.budget_workforce_hours_type2.replace(",","."));
       if (typeof doc.budget_workforce_hours_type3 === 'string') doc.budget_workforce_hours_type3     = parseFloat(doc.budget_workforce_hours_type3.replace(",","."));
       if (typeof doc.budget_workforce_hours_type4 === 'string') doc.budget_workforce_hours_type4     = parseFloat(doc.budget_workforce_hours_type4.replace(",","."));
       if (typeof doc.budget_total_cost            === 'string') doc.budget_total_cost                = parseFloat(doc.budget_total_cost.replace(",","."));
       if (typeof doc.budget_material_cost         === 'string') doc.budget_material_cost             = parseFloat(doc.budget_material_cost.replace(",","."));
       if (typeof doc.budget_subcontracting_cost   === 'string') doc.budget_subcontracting_cost       = parseFloat(doc.budget_subcontracting_cost.replace(",","."));
       if (typeof doc.budget_workforce_total_cost  === 'string') doc.budget_workforce_total_cost      = parseFloat(doc.budget_workforce_total_cost.replace(",","."));
       if (typeof doc.conso_workforce_total_cost   === 'string') doc.conso_workforce_total_cost       = parseFloat(doc.conso_workforce_total_cost.replace(",","."));
       if (typeof doc.conso_workforce_total_hours  === 'string') doc.conso_workforce_total_hours      = parseFloat(doc.conso_workforce_total_hours.replace(",","."));
       if (typeof doc.conso_material_cost          === 'string') doc.conso_material_cost              = parseFloat(doc.conso_material_cost.replace(",","."));
       if (typeof doc.conso_subcontracting_cost    === 'string') doc.conso_subcontracting_cost        = parseFloat(doc.conso_subcontracting_cost.replace(",","."));
       if (typeof doc.conso_total_cost             === 'string') doc.conso_total_cost                 = parseFloat(doc.conso_total_cost.replace(",","."));
       if (typeof doc.risk_amount                  === 'string') doc.risk_amount                      = parseFloat(doc.risk_amount.replace(",","."));
       if (typeof doc.prev_workforce_total_cost    === 'string') doc.prev_workforce_total_cost        = parseFloat(doc.prev_workforce_total_cost.replace(",","."));
       if (typeof doc.prev_workforce_total_hours   === 'string') doc.prev_workforce_total_hours       = parseFloat(doc.prev_workforce_total_hours.replace(",","."));
       if (typeof doc.prev_material_cost           === 'string') doc.prev_material_cost               = parseFloat(doc.prev_material_cost.replace(",","."));
       if (typeof doc.prev_subcontracting_cost     === 'string') doc.prev_subcontracting_cost         = parseFloat(doc.prev_subcontracting_cost.replace(",","."));
       if (typeof doc.prev_marging                 === 'string') doc.prev_marging                     = parseFloat(doc.prev_marging.replace(",","."));
       if (typeof doc.prev_marging_ratio           === 'string') doc.prev_marging_ratio               = parseFloat(doc.prev_marging_ratio.replace(",","."));
       if (typeof doc.fae                          === 'string') doc.fae                              = parseFloat(doc.fae.replace(",","."));
       if (typeof doc.billed_amout                 === 'string') doc.billed_amout                     = parseFloat(doc.billed_amout.replace(",","."));
       if (typeof doc.recoverable_amount           === 'string') doc.recoverable_amount               = parseFloat(doc.recoverable_amount.replace(",","."));
       if (typeof doc.paid_amount                  === 'string') doc.paid_amount                      = parseFloat(doc.paid_amount.replace(",","."));
       if (typeof doc.realised_amount              === 'string') doc.realised_amount                  = parseFloat(doc.realised_amount.replace(",","."));
       if (typeof doc.prev_revenu_current_year     === 'string') doc.prev_revenu_current_year         = parseFloat(doc.prev_revenu_current_year.replace(",","."));
       if (typeof doc.progress                     === 'string') doc.progress                         = parseFloat(doc.progress.replace(",","."));

        var record = Object({
            imputation_code : doc.imputation_code,
            progress : doc.progress,
            status : doc.status,
            owners: {
                business : doc.business_owner,
                execution : doc.execution_owner,
                field : doc.field_owner
            },
            type : doc.type,
            contract : {
                signing_date : doc.signing_date,
                client : doc.client,
                name : doc.name,
                revenu: doc.initial_revenu
            },
            schedule : {
                start_date :  ((doc.start_date !== null)  ? ExcelDateToJSDate(doc.start_date) : null),
                duration : doc.duration,
            },
            preparation : {
                status : doc.preparation_status
            },
            budget : {
                workloads : [
                    { 
                        type : 1,
                        workload : doc.budget_workforce_hours_type1,
                        rate : doc.field_staff_std_rate
                    },
                    { 
                        type : 2,
                        workload : doc.budget_workforce_hours_type2,
                        rate : doc.field_staff_std_rate
                    },
                    { 
                        type : 3,
                        workload : doc.budget_workforce_hours_type3,
                        rate : doc.field_staff_std_rate
                    },
                    { 
                        type : 4,
                        workload : doc.budget_workforce_hours_type4,
                        rate : doc.field_staff_std_rate
                    }
                ],
                workforce : doc.budget_workforce_total_cost,
                material: doc.budget_material_cost,
                subcontracting :  doc.budget_subcontracting_cost,
                total :  doc.budget_total_cost
            },
            conso : {
                workload : doc.conso_workforce_total_hours,
                workforce : doc.conso_workforce_total_cost,
                material : doc.conso_material_cost,
                subcontracting :  doc.conso_subcontracting_cost,
                total :  doc.conso_total_cost
            },
            prev : {
                workload : doc.prev_workforce_total_hours,
                workforce : doc.prev_workforce_total_cost,
                material : doc.prev_material_cost,
                subcontracting :  doc.prev_subcontracting_cost,
                total :  doc.prev_total_cost
            },
            billing : {
                billable : doc.fae, 
                billed : doc.billed_amout,
                paid : doc.paid_amount,
                realised : doc.realised_amount,
                current_year_revenu : doc.prev_revenu_current_year
            }        
        });
       
        firestore.collection(collectionKey).doc().set(record).then((res) => {
            console.log("Document " + docKey + " successfully written!");
        }).catch((error) => {
            console.error("Error writing document: ", error);
        });
    });
}

function ExcelDateToJSDate(serial) {
    var utc_days  = Math.floor(serial - 25569);
    var utc_value = utc_days * 86400;                                        
    var date_info = new Date(utc_value * 1000);
 
    var fractional_day = serial - Math.floor(serial) + 0.0000001;
 
    var total_seconds = Math.floor(86400 * fractional_day);
 

 
    return new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate(), 0, 0, 0);
 }