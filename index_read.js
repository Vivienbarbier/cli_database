const { initializeApp, applicationDefault, cert } = require('firebase-admin/app');
const { getFirestore, Timestamp, FieldValue } = require('firebase-admin/firestore')
require('dotenv').config();
const serviceAccount = require('./serviceAccountKey.json');

initializeApp({
  credential: cert(serviceAccount)
});

var db = getFirestore();

const admin = require('firebase-admin');
const collectionKey = "case"; //name of the collection
const business_unit_code = "OCPL"

    db.collection(collectionKey).doc(business_unit_code).get().then((docRef) => {
        console.log("Sucess");
        console.log(JSON.stringify(docRef.data));
    }).catch((error) => {
        console.error("Error reading document: ", error);
    });   
