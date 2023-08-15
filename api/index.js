const express = require('express');
const bodyParser = require('body-parser');
const Finance = require('financejs');
const yahooFinance = require('yahoo-finance2');
const admin = require("firebase-admin");
//const serviceAccount = require("./serviceAccount.json");

//! To do
//admin.initializeApp({
    //credential: admin.credential.cert(serviceAccount),
    //databaseURL: "placeholder",
    //storageBucket: "placeholder"
//})

const app = express();
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.json())
app.use(cors({ origin: true }));
const db = admin.database()
const auth = admin.auth()

app.post('/api/createbudgetEntry', async(req, res)=>{
    const {user_ID, amount_type, amount_category, amount, notes, date, timestamp} = req.body
    const user_budget_ref = ref(db, `/${user_ID}/budgeting`)
    try{
        const postData = {
            user_ID: user_ID,
            category: amount_category,
            amount_type: amount_type,
            amount: amount,
            notes: notes,
            date: date,
            timestamp: timestamp
        }
        const newPostKey = push(user_budget_ref).key
        const updates = {};
        updates[`/${user_ID}/budgeting/${newPostKey}`] = postData;
        updates[`/user_budget_data/${user_ID}/${newPostKey}`] = postData;
        res.status(200)
        return await update(ref(db), updates)
    }
    catch(error){
        res.send(500)
    }
})




app.post('/api/investment', async(req, res)=>{
    const {ticker} = req.body
    try{
        const queryModules = ['price', 'financialData', 'summaryDetail']
        const financial_data = await yahooFinance.quoteSummary(ticker, queryModules)
        res.status(200).json(financial_data)
    }
    catch(error){
        res.status(500).json({"error": error})
    }

})


app.post('/api/saveInvestment', async(req, res)=>{
    const {user_ID, company, shareCount, investment_amount, company_type, company_industry, investment_date, investment_id, timestamp} = req.body
    const user_investment_ref = ref(db, `/user_investment_data/${user_ID}`)
    try{
        const postData = {
            user_ID: user_ID,
            company: company,
            shareCount: shareCount,
            investment_amount: investment_amount,
            amount: investment_amount,
            category: company_industry,
            price_per_share: investment_amount / shareCount,
            company_type: company_type,
            company_industry: company_industry,
            investment_date: investment_date,
            investment_id: investment_id,
            timestamp: timestamp
        }
        const newPostKey = push(user_investment_ref).key
        const updates = {};
        updates[`/${user_ID}/investments/${newPostKey}`] = postData;
        updates[`/user_investment_data/${user_ID}/${newPostKey}`] = postData;
        res.status(200)
        return await update(ref(db), updates)
    }
    catch(error){
        res.status(500).json({"error": error})
    }
})