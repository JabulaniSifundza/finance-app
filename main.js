import { initializeApp } from "https://www.gstatic.com/firebasejs/9.8.2/firebase-app.js";
import { getAnalytics } from "https://www.gstatic.com/firebasejs/9.8.2/firebase-analytics.js";
import { getDatabase, ref, child, get, onValue, query, equalTo, orderByChild } from "https://www.gstatic.com/firebasejs/9.8.2/firebase-database.js";
import { getAuth, onAuthStateChanged, signOut, createUserWithEmailAndPassword } from "https://www.gstatic.com/firebasejs/9.8.2/firebase-auth.js";
import { getFirestore, collection, getDoc, doc, deleteDoc} from "https://www.gstatic.com/firebasejs/9.8.2/firebase-firestore.js";




const firebaseConfig = {
  apiKey: "AIzaSyAqMaNzWL8du_EyUBtfnBH8tozqfiFQy38",
  authDomain: "personal-finance-models.firebaseapp.com",
  projectId: "personal-finance-models",
  storageBucket: "personal-finance-models.appspot.com",
  messagingSenderId: "545682365018",
  appId: "1:545682365018:web:c95098d90b5ad412aad663",
  measurementId: "G-Y50XMW7D2E"
};

// Initialize Firebase
const app = initializeApp(firebaseConfig);
const analytics = getAnalytics(app);
const db = getFirestore(app);
const auth = getAuth(app);

const createAccount = document.getElementById("createAccount")
const create_account_btn = document.getElementById("create-account-btn")
const amountType = document.getElementById("amountType").value
const amountCategory = document.getElementById("category").value
const amount = document.getElementById("entryAmount").value
const amountDate = document.getElementById("amountDate").value
const notes = document.getElementById("amountNote").value

const user_email = document.getElementById("user-email").value;
const user_pass = document.getElementById("user-password").value;

let budgetHistory = document.getElementById("budgetActivityBreakdown")

function addToHistory(budgetEntry){
    const html = `
        <div class=${budgetEntry.amountType === "Income" ? "breakDownSymbol" : budgetEntry.amountType === "Expenditure" ? "breakDownSymbolRed" : budgetEntry.amountType === "Savings" ? "breakDownSymbolGreen" : "breakDownSymbol"}>
            <span class="material-symbols-outlined">
                ${budgetEntry.amountCategory === "Salary" ? "payments" : budgetEntry.amountCategory === "Investment Return" ? "monitoring" : budgetEntry.amountCategory === "Loan Acquisition" ? "account_balance" : budgetEntry.amountCategory === "Loan Installment" ? "account_balance" : budgetEntry.amountCategory === "Grocery" ? "shopping_basket" : budgetEntry.amountCategory === "Rent" ? "real_estate_agent" : budgetEntry.amountCategory === "Utilities" ? "lightbulb" : budgetEntry.amountCategory === "Entertainment" ? "sports_esports" : budgetEntry.amountCategory === "Shopping" ? "local_mall" : budgetEntry.amountCategory === "Travel" ? "luggage" : budgetEntry.amountCategory === "Other" ? "other_admission" : "more_horiz"}
            </span>
        </div>
        <p>${budgetEntry.amountCategory}</p>
        <h4>$ ${budgetEntry.amount}</h4>
    `;
    let newEntry = document.createElement("div");
    newEntry.classList.add("breakDownCard");
    newEntry.innerHTML = html;
    budgetHistory.prepend(newEntry);
}
const addBudgetEntry = document.getElementById("createEntry")
function saveBudgetEntry(){
    onAuthStateChanged(auth, async(user)=>{
        if(user){
            try{
                const userId = user.uid
                const entry_data = await fetch('/api/createbudgetEntry', {
                    method: 'POST',
                    headers:{
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify({
                        user: userId,
                        amountType: amountType,
                        amountCategory: amountCategory,
                        amount: amount,
                        amountDate: amountDate,
                        notes: notes,
                        budget_entry_timestamp: Math.floor(Date.now() / 1000)
                    })
                })
                const data = await entry_data.json()
                console.log(data)
                //addToHistory()
            }
            catch(error){
                alert(`Oops... We are sorry but an error has occurred: ${error}`)
            }
        }
        else{
            createAccount.showModal()
        }
    })  
};

const create_new_user_account = async(email, password)=>{
  try{
    const user_credential = await createUserWithEmailAndPassword(auth, email, password)
    const user = user_credential.user
    const uid = user.uid
    set(ref(db, `/user_accounts/${uid}`),{
      fullname: document.getElementById("userFirstName").value,
      email: document.getElementById("user-email").value,
      user_id: uid
    }).then(()=>{
      createAccount.close();
    })
  }
  catch(error){
    alert(`Unfortunately, the following error has occurred: ${error}`)
  }
}

addBudgetEntry.addEventListener("click", (event)=>{
    event.preventDefault();
    saveBudgetEntry()
    //console.log("Log")
})

create_account_btn.addEventListener("click", (event)=>{
    event.preventDefault();
    create_new_user_account(user_email, user_pass)
})
document.getElementById("signupModalClose").addEventListener("click", ()=>{
    createAccount.close();
})

document.getElementById('show-password').addEventListener('change', function (e) {
  const passwordInput = document.getElementById('user-password');
  if (e.target.checked) {
    passwordInput.type = 'text';
  } else {
    passwordInput.type = 'password';
  }
});

// Getting and display budget data
const read_user_budget_data = async()=>{
    onAuthStateChanged(auth, async(user)=>{
        if(user){
            try{
                const userId = user.uid
                onValue(`/user_budget_data/${userId}`, (snapshot)=>{
                    const budget_data = snapshot.val()
                    let expense_by_type = d3.nest()
                    .key(d => d.category)
                    .rollup(v => d3.sum(v, d => d.amount))
                    .entries(budget_data)

                    let daily_expenditures = d3.nest()
                    .key(d => d3.timeDay(new Date(d.date)))
                    .rollup(v => d3.sum(v, d => d.amount))
                    .entries(budget_data)

                    let weekly_expenditures = d3.nest()
                    .key(d => d3.timeWeek(new Date(d.date)))
                    .rollup(v => d3.sum(v, d => d.amount))
                    .entries(budget_data)

                    let monthly_expenditures = d3.nest()
                    .key(d => d3.timeMonth(new Date(d.date)))
                    .rollup(v => d3.sum(v, d => d.amount))
                    .entries(budget_data)

                    let annual_expenditures = d3.nest()
                    .key(d => d3.timeYear(new Date(d.date)))
                    .rollup(v => d3.sum(v, d => d.amount))
                    .entries(budget_data)

                    // Expenditure Pie Chart
                    let pie = d3.pie()
                    .value(d => d.value)
                    let data_ready = pie(expense_by_type)
                    let svg = d3.select("#expenditure_pie")
                    .append("svg")
                    .attr('width', 120)
                    .attr('height', 120)
                    .append('g')
                    .attr('transform', 'translate(' + 120 / 2 + ',' + 120 / 2 + ')');
                    let color = d3.scaleOrdinal(["#98abc5", "#8a89a6", "#7b6888", "#6b486b", "#a05d56", "#d0743c", "#ff8c00"]);
                    svg
                    .selectAll('mySlices')
                    .data(data_ready)
                    .enter()
                    .append('path')
                    .attr('d', d3.arc()
                        .innerRadius(0)  
                        .outerRadius(200)
                    )
                    .attr('fill', d => color(d.data.key))
                    .attr("stroke", "white")
                    .style("stroke-width", "2px")
                    svg.selectAll("path")
                    .data(pie)
                    .enter().append("path")
                    .attr("d", d3.arc())
                    .attr("fill", "steelblue")
                    .transition()
                    .ease(d3.easeBounce)
                    .duration(1000)
                    .delay((d, i) => i * 500)
                    .attrTween("d", d3.arcTween());


                    // Expenditure Trends

                    // Daily spending
                    let svg_daily = d3.select("#daily_expenditure")
                    .append('svg')
                    .attr('width', 120)
                    .attr('height', 120)
                    let xScale = d3.scaleTime().domain([d3.min(budget_data, d => new Date(d.date)), d3.max(budget_data, d => new Date(d.date))]).range([0, 120]);
                    let yScale = d3.scaleLinear().domain([0, d3.max(budget_data, d => d.amount)]).range([120, 0]);
                    let line = d3.line()
                    .x(d => xScale(new Date(d.key)))
                    .y(d => yScale(d.value));

                    svg_daily.append("defs").append("filter")
                    .attr("id", "drop-shadow")
                    .attr("height", "130%")
                    .append("feGaussianBlur")
                    .attr("in", "SourceAlpha")
                    .attr("stdDeviation", 5)
                    .attr("result", "blur");
                    
                    svg_daily.append("path")
                    .datum(daily_expenditures)
                    .attr("fill", "none")
                    .attr("stroke", "steelblue")
                    .attr("stroke-width", 1.5)
                    .attr("d", line);

                    svg_daily.selectAll("path")
                    .data(line)
                    .enter().append("path")
                    .attr("d", line)
                    .attr("fill", "steelblue")
                    .transition()
                    .ease(d3.easeBounce)
                    .duration(1000)
                    .delay((d, i) => i * 500)
                    .attrTween("d", d3.arcTween());

                    // Weekly spending
                    let svg_weekly = d3.select("#weekly_expenditure")
                    .append('svg')
                    .attr('width', 120)
                    .attr('height', 120)
                    let x_weekly_Scale = d3.scaleTime().domain([d3.min(budget_data, d => new Date(d.date)), d3.max(budget_data, d => new Date(d.date))]).range([0, 120]);
                    let y_weekly_Scale = d3.scaleLinear().domain([0, d3.max(budget_data, d => d.amount)]).range([120, 0]);
                    let weekly_line = d3.line()
                    .x(d => x_weekly_Scale(new Date(d.key)))
                    .y(d => y_weekly_Scale(d.value));
                    svg_weekly.append("defs").append("filter")
                    .attr("id", "drop-shadow")
                    .attr("height", "130%")
                    .append("feGaussianBlur")
                    .attr("in", "SourceAlpha")
                    .attr("stdDeviation", 5)
                    .attr("result", "blur");

                    svg_weekly.append("path")
                    .datum(weekly_expenditures)
                    .attr("fill", "none")
                    .attr("stroke", "steelblue")
                    .attr("stroke-width", 1.5)
                    .attr("d", weekly_line);

                    svg_weekly.selectAll("path")
                    .data(weekly_line)
                    .enter().append("path")
                    .attr("d", weekly_line)
                    .attr("fill", "steelblue")
                    .transition()
                    .ease(d3.easeBounce)
                    .duration(1000)
                    .delay((d, i) => i * 500)
                    .attrTween("d", d3.arcTween());

                    // Monthly spending
                    let svg_Monthly = d3.select("#weekly_expenditure")
                    .append('svg')
                    .attr('width', 120)
                    .attr('height', 120)
                    let x_monthly_Scale = d3.scaleTime().domain([d3.min(budget_data, d => new Date(d.date)), d3.max(budget_data, d => new Date(d.date))]).range([0, 120]);
                    let y_monthly_Scale = d3.scaleLinear().domain([0, d3.max(budget_data, d => d.amount)]).range([120, 0]);
                    let monthly_line = d3.line()
                    .x(d => x_monthly_Scale(new Date(d.key)))
                    .y(d => y_monthly_Scale(d.value));

                    svg_Monthly.append("defs").append("filter")
                    .attr("id", "drop-shadow")
                    .attr("height", "130%")
                    .append("feGaussianBlur")
                    .attr("in", "SourceAlpha")
                    .attr("stdDeviation", 5)
                    .attr("result", "blur");
                    
                    svg_Monthly.append("path")
                    .datum(monthly_expenditures)
                    .attr("fill", "none")
                    .attr("stroke", "steelblue")
                    .attr("stroke-width", 1.5)
                    .attr("d", monthly_line);

                    svg_Monthly.selectAll("path")
                    .data(monthly_line)
                    .enter().append("path")
                    .attr("d", monthly_line)
                    .attr("fill", "steelblue")
                    .transition()
                    .ease(d3.easeBounce)
                    .duration(1000)
                    .delay((d, i) => i * 500)
                    .attrTween("d", d3.arcTween());

                    // Yearly spending
                    let svg_annual = d3.select("#weekly_expenditure")
                    .append('svg')
                    .attr('width', 120)
                    .attr('height', 120)
                    let x_annual_Scale = d3.scaleTime().domain([d3.min(budget_data, d => new Date(d.date)), d3.max(budget_data, d => new Date(d.date))]).range([0, 120]);
                    let y_annual_Scale = d3.scaleLinear().domain([0, d3.max(budget_data, d => d.amount)]).range([120, 0]);
                    let annual_line = d3.line()
                    .x(d => x_annual_Scale(new Date(d.key)))
                    .y(d => y_annual_Scale(d.value));

                    svg_annual.append("defs").append("filter")
                    .attr("id", "drop-shadow")
                    .attr("height", "130%")
                    .append("feGaussianBlur")
                    .attr("in", "SourceAlpha")
                    .attr("stdDeviation", 5)
                    .attr("result", "blur");

                    svg_annual.append("path")
                    .datum(annual_expenditures)
                    .attr("fill", "none")
                    .attr("stroke", "steelblue")
                    .attr("stroke-width", 1.5)
                    .attr("d", annual_line);

                    svg_annual.selectAll("path")
                    .data(annual_line)
                    .enter().append("path")
                    .attr("d", annual_line)
                    .attr("fill", "steelblue")
                    .transition()
                    .ease(d3.easeBounce)
                    .duration(1000)
                    .delay((d, i) => i * 500)
                    .attrTween("d", d3.arcTween());

                })

                const budget_ref = query(
                    ref(db, `/${userId}/budgeting`),
                    orderByChild("timestamp"),
                    limitToLast(5)
                )
                 onValue(budget_ref, (snapshot)=>{
                    snapshot.forEach((entry)=>{
                        const entry_info = entry.val()
                        addToHistory(entry_info)
                    })
                })
                
            }
            catch(error){
                alert(`Oops... We are sorry but an error has occurred: ${error}`)
            }
        }
        else{
            createAccount.showModal()
        }
    })  
}
// Investment Section
// Public
const tickerSymbol = document.getElementById("tickerSymbol").value
const tickerInfo = document.getElementById("tickerInfo")
const publicCompanyIndustry = document.getElementById("publicCompanyIndustry").value
let investmentDetails = document.getElementById("publicCompanyInfo")
let publicSharePrice = ""
//Private company
const privateCompanyName = document.getElementById("privateCompanyName").value
const privateCompanyIndustry = document.getElementById("privateCompanyIndustry").value
const addPrivateInvestment = document.getElementById("addPrivateInvestment")

let investmentHistory = document.getElementById("investmentBreakdown")

document.getElementById("publicCompanyLookup").addEventListener("click", ()=>{
    getPublicCompany()
})

const read_user_investments = async()=>{
    onAuthStateChanged(auth, async(user)=>{
        if(user){
            try{
                const user_id = user.uid
                onValue(`/user_investment_data/${user_id}`, (snapshot)=>{
                    const investment_data = snapshot.val()
                    let investment_by_type = d3.nest()
                    .key(d => d.category)
                    .rollup(v => d3.sum(v, d => d.amount))
                    entries(investment_data)

                    let annual_investment = d3.nest()
                    .key(d => d3.timeYear(new Date(d.date)))
                    .rollup(v => d3.sum(v, d => d.amount))
                    entries(investment_data)


                    let pie = d3.pie()
                    .value(d => d.value)
                    let data_ready = pie(investment_by_type)
                    let svg = d3.select("#investment_pie")
                    .append("svg")
                    .attr('width', 120)
                    .attr('height', 120)
                    .append('g')
                    .attr('transform', 'translate(' + 120 / 2 + ',' + 120 / 2 + ')');
                    let color = d3.scaleOrdinal(["#98abc5", "#8a89a6", "#7b6888", "#6b486b", "#a05d56", "#d0743c", "#ff8c00"]);
                    svg
                    .selectAll('mySlices')
                    .data(data_ready)
                    .enter()
                    .append('path')
                    .attr('d', d3.arc()
                        .innerRadius(0)  
                        .outerRadius(200)
                    )
                    .attr('fill', d => color(d.data.key))
                    .attr("stroke", "white")
                    .style("stroke-width", "2px")
                    svg.selectAll("path")
                    .data(pie)
                    .enter().append("path")
                    .attr("d", d3.arc())
                    .attr("fill", "steelblue")
                    .transition()
                    .ease(d3.easeBounce)
                    .duration(1000)
                    .delay((d, i) => i * 500)
                    .attrTween("d", d3.arcTween());

                    // Annual Investing
                    let svg_annual = d3.select("#daily_expenditure")
                    .append('svg')
                    .attr('width', 120)
                    .attr('height', 120)
                    let xScale = d3.scaleTime().domain([d3.min(investment_data, d => new Date(d.date)), d3.max(investment_data, d => new Date(d.date))]).range([0, 120]);
                    let yScale = d3.scaleLinear().domain([0, d3.max(investment_data, d => d.amount)]).range([120, 0]);
                    let line = d3.line()
                    .x(d => xScale(new Date(d.key)))
                    .y(d => yScale(d.value));

                    svg_annual.append("defs").append("filter")
                    .attr("id", "drop-shadow")
                    .attr("height", "130%")
                    .append("feGaussianBlur")
                    .attr("in", "SourceAlpha")
                    .attr("stdDeviation", 5)
                    .attr("result", "blur");
                    
                    svg_annual.append("path")
                    .datum(annual_investment)
                    .attr("fill", "none")
                    .attr("stroke", "steelblue")
                    .attr("stroke-width", 1.5)
                    .attr("d", line);

                    svg_annual.selectAll("path")
                    .data(line)
                    .enter().append("path")
                    .attr("d", line)
                    .attr("fill", "steelblue")
                    .transition()
                    .ease(d3.easeBounce)
                    .duration(1000)
                    .delay((d, i) => i * 500)
                    .attrTween("d", d3.arcTween());
                })
                const investment_ref = query(
                    ref(db, `/${user_id}/investments`),
                    orderByChild("timeStamp"),
                    limitToLast(5)
                )
                onValue(investment_ref, (snapshot)=>{
                    snapshot.forEach((investment)=>{
                        const investment_info = investment.val()
                        addToInvestmentHistory(investment_info)
                    })
                })
            }
            catch(error){
                alert(`Oops... We are sorry but an error has occurred: ${error}`)
            }
        }
        else{
            
        }
        
    })
}

async function getPublicCompany(){
    const companyInfo = await fetch(`/api/investment`,{
        method: 'POST',
        headers:{
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            tickerSymbol: tickerSymbol
        })
    })
    const data = await companyInfo.json()
    const sharePrice = formatCurrency(data.summaryDetail.previousClose)
    const high = formatCurrency(data.summaryDetail.dayHigh)
    publicSharePrice += sharePrice
    const html = `
        <div>
            <p id="publicCompanySharePrice">
                Company Stock Price: $ ${sharePrice}
            </p>
            <p>
                Return on Equity: ${Number(data.financialData.returnOnEquity)}
            </p>
            <p>
                High: $ ${high}
            </p>
            <p>
                Analyst Recommendation: ${data.financialData.recommendationKey}
            </p>
        </div>
        <div>
            <input type="number" placeholder="Number of shares" id="publicCompanyShares">
            <button class="addToInvestment" id="addPublicCompany">Add Investment</button>
        </div>
    `;
    let newPublicCompany = document.createElement("div");
    newPublicCompany.classList.add("breakDownCard");
    newPublicCompany.innerHTML = html;
    investmentDetails.append(newPublicCompany);
    assignPublicInvestmentEvents()
}


const saveInvestment = async(shareCount, investmentAmount, userId) =>{
    const response = await fetch(`/api/saveInvestment`,{
        method: 'POST',
        headers:{
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            user: userId,
            company: tickerSymbol ? tickerSymbol : privateCompanyName,
            shareCount: shareCount,
            investmentAmount: investmentAmount,
            companyType: publicCompanyIndustry ? "Public" : "Private",
            companyIndustry: publicCompanyIndustry ? publicCompanyIndustry : privateCompanyIndustry,
            investmentDate: `${year}-${month}-${day}`,
            investmentId: tickerSymbol ? `${tickerSymbol-userId}` : `${privateCompanyName-userId}`,
            timeStamp: Math.floor(Date.now() / 1000)
        })
    })
    const data = await response.json()
    console.log(data)
}
document.getElementById("addPrivateInvestment").addEventListener("click", ()=>{
    onAuthStateChanged(auth, (user)=>{
        if(user){
            try{
                const user_id = user.id
                const today = new Date();
                const year = today.getFullYear();
                const month = today.getMonth() + 1;
                const day = today.getDate();
                const shareCount = "Unknown"
                const privateInvestmentValue = document.getElementById("privateInvestmentValue").value
                saveInvestment(shareCount, privateInvestmentValue, user_id)
                //addToInvestmentHistory(investmentEntry)
            }
            catch(error){
                alert(`Oops... We are sorry but an error has occurred: ${error}`)
            }
        }
        else{
            createAccount.showModal()
        }
    })

})

function assignPublicInvestmentEvents(){
    const shareCount = document.getElementById("publicCompanyShares")
    const addPublicCompany = document.getElementById("addPublicCompany")
    const publicCompanySharePrice = document.getElementById("publicCompanySharePrice").innerText;
    const [ , ,monetaryValue] = publicCompanySharePrice.split(" ")
    const publicInvestmentAmount = Number(shareCount) * Number(monetaryValue)
    addPublicCompany.addEventListener("click", ()=>{
        onAuthStateChanged(auth, (user)=>{
            if(user){
                const user_id = user.id
                const today = new Date();
                const year = today.getFullYear();
                const month = today.getMonth() + 1;
                const day = today.getDate();
                saveInvestment(shareCount, publicInvestmentAmount, user_id)
                //addToInvestmentHistory(investmentEntry)
            }
            else{
                createAccount.showModal()
            }
        })
    })
}




function addToInvestmentHistory(investmentEntry){
    const html = `
        <div class=${investmentEntry.companyIndustry === "Food" ? "breakDownSymbol" : investmentEntry.companyIndustry === "Financial Services" ? "breakDownSymbolGreen" : investmentEntry.companyIndustry === "Entertainment" ? "breakDownSymbol" : investmentEntry.companyIndustry === "Technology" ? "breakDownSymbolGreen" : investmentEntry.companyIndustry === "Agriculture" ? "breakDownSymbol" : investmentEntry.companyIndustry === "Pharmaceuticals" ? "breakDownSymbolGreen" : investmentEntry.companyIndustry === "Hospitality" ? "breakDownSymbol" : investmentEntry.companyIndustry === "Retail" ? "breakDownSymbolGreen" : investmentEntry.companyIndustry === "Telecommunications" ? "breakDownSymbol" : investmentEntry.companyIndustry === "Other" ? "breakDownSymbolGreen" : "breakDownSymbol"}>
            <span class="material-symbols-rounded">
                ${investmentEntry.companyIndustry === "Food & Beverage" ? "fastfood" : investmentEntry.companyIndustry === "Financial Services" ? "paid" : investmentEntry.companyIndustry === "Entertainment" ? "sports_esports" : investmentEntry.companyIndustry === "Technology" ? "memory" : investmentEntry.companyIndustry === "Agriculture" ? "agriculture" : investmentEntry.companyIndustry === "Pharmaceuticals" ? "medication" : investmentEntry.companyIndustry === "Hospitality" ? "room_service" : investmentEntry.companyIndustry === "Retail" ? "local_mall" : investmentEntry.companyIndustry === "Telecommunications" ? "perm_phone_msg" : investmentEntry.companyIndustry === "Other" ? "other_admission" : "more_horiz"}
            <span>
            <p class="companyName" data-investmentID=${investmentEntry.investmentId}>
                ${investmentEntry.company} 
            </p>
            <h4>$ ${investmentEntry.investmentAmount}</h4>
        </div>
    `;
    let newInvestment = document.createElement("div")
    newInvestment.classList.add("breakDownCard")
    newInvestment.innerHTML = html
    investmentHistory.prepend(newInvestment)
    assignInvestmentHistoryEvent()
}
function assignInvestmentHistoryEvent(){
    const investmentName = document.getElementsByClassName("companyName")
    for(let i = 0; i < investmentName.length; i++){
        investmentName[i].addEventListener("click", openInvestmentModal)
    }
}
async function openInvestmentModal(event){
    const editInvestment = document.getElementById("editInvestment")
    editInvestment.showModal()
    const investmentId = event.taget.getAttribute("data-investmentID")
    const docRef = doc(db, "Investments", investmentId);
    const docSnap = await getDoc(docRef);
    const data =  docSnap.data()
    let investmentInfoDiv = document.getElementById("investmentInfo")
    const html = `
        <p>Company Name: ${data.company}</p>
        <div>
            <p id="investShareCount">Share Count: ${data.shareCount}</p>
            <span class="material-symbols-rounded" data-id=${investmentId}>
                remove
            </span>
            <span class="material-symbols-rounded" data-id=${investmentId}>
                add
            </span>
        </div>
        <div>
            <p>Investment Amount: </p>
            <span id="investmentAmount" data-id=${investmentId}>${data.investmentAmount}</span>
        </div>
        <div>
            <button id="updateEdited">Update</button>
            <button id="deleteEdited" data-id=${investmentId}>Delete</button>
        </div>
    `;
    let investmentDetails = document.createElement("div")
    investmentDetails.classList.add("investmentDetails")
    investmentDetails.innerHTML = html
    investmentInfoDiv.append(investmentDetails)
    assignInvestmentEditEvents()
}

function assignInvestmentEditEvents(){
    const addShare = document.getElementById("addShare")
    const removeShare = document.getElementById("removeShare")
    const deleteInvestment = document.getElementById("deleteEdited")
    const updateCurrentInvestment = document.getElementById("updateEdited")
    const investmentAmount = document.getElementById("investmentAmount")
    addShare.addEventListener("click", ()=>{
        increaseShareCount()
    })
   removeShare.addEventListener("click", ()=>{
        decreaseShareCount()
    })
    deleteInvestment.addEventListener("click", ()=>{
        removeInvestment()
    })
    updateCurrentInvestment.addEventListener("click", ()=>{
        updateInvestment()
    })
    investmentAmount.addEventListener("click", ()=>{
        changeInvestmentAmount()
    })
}
function updateInvestment(event){
    const investmentId = event.target.getAttribute("data-id")
    const investmentAmount = document.getElementById("investmentAmount")
    const docRef = firebase.firestore().collection("Investments").doc(investmentId);
    docRef.update({
        investmentAmount: investmentAmount.innerHTML
    })

}

function increaseShareCount(event){
    const investmentId = event.target.getAttribute("data-id")
    const count = document.getElementById("investShareCount")
    const docRef = firebase.firestore().collection("Investments").doc(investmentId);
    docRef.update({
        shareCount: firebase.firestore.FieldValue.increment(1)
    })
}

function decreaseShareCount(event){
    const investmentId = event.target.getAttribute("data-id")
    const count = document.getElementById("investShareCount")
    const docRef = firebase.firestore().collection("Investments").doc(investmentId);
    docRef.update({
        shareCount: firebase.firestore.FieldValue.decrement(1)
    })
}

function changeInvestmentAmount(event){
    const element = event.target
    if(!element.contentEditable === true){
        element.contentEditable = true
        element.focus()
    }
    else{
        element.contentEditable = false
    }
}

async function removeInvestment(event){
    const investmentId = event.target.getAttribute("data-id")
    await deleteDoc(doc(db, "Investments", investmentId))
    location.reload()
}

function create_LBO(){
    // LBO Model Data Processing

const Historical_Revenue = document.getElementById("Historical_Revenue").value
const Current_Revenue = document.getElementById("Current_Revenue").value
const Historical_COGS = document.getElementById("Historical_COGS").value
const Current_COGS = document.getElementById("Current_COGS").value
const Historical_SGA = document.getElementById("Historical_SG&A").value
const Current_SGA = document.getElementById("Current_SG&A").value
const Historical_Depreciation = document.getElementById("Historical_Depreciation").value
const Current_Depreciation = document.getElementById("Current_Depreciation").value
const Historical_Amortization = document.getElementById("Historical_Amortization").value
const Current_Amortization = document.getElementById("Current_Amortization").value
const Historical_Interest_Expense = document.getElementById("Historical_Interest_Expense").value
const Current_Interest_Expense = document.getElementById("Current_Interest_Expense").value
const Historical_Cash_Balance = document.getElementById("Historical_Cash_Balance").value
const LBO_Accounts_Receivable = document.getElementById("LBO_Accounts_Receivable").value
const LBO_Inventory = document.getElementById("LBO_Inventory").value
const LBO_Prepaid_Expenses = document.getElementById("LBO_Prepaid_Expenses").value
const PPE_Depreciation = document.getElementById("PP&E_Depreciation").value
const Capitalized_Financing_Fee = document.getElementById("Capitalized_Financing_Fee").value
const LBO_Goodwill = document.getElementById("LBO_Goodwill").value
const LBO_Other_Long_Term_Asset = document.getElementById("LBO_Other_Long_Term_Asset").value
const LBO_Accounts_Payable = document.getElementById("LBO_Accounts_Payable").value
const LBO_Notes_Payable = document.getElementById("BO_Notes_Payable").value
const LBO_Long_Term_Debt = document.getElementById("LBO_Long_Term_Debt").value
const LBO_Common_Stock = document.getElementById("LBO_Common_Stock").value
const LBO_Retained_Earnings = document.getElementById("LBO_Retained_Earnings").value
const LBO_Interest_Paid_Revolving = document.getElementById("LBO_Interest_Paid_Revolving").value
const LBO_Interest_Paid_Subordinated = document.getElementById("LBO_Interest_Paid_Subordinated").value
const LBO_Loan_Period_Senior = document.getElementById("LBO_Loan_Period_Senior").value
const LBO_Interest_Paid_Senio = document.getElementById("LBO_Interest_Paid_Senior").value
const LBO_Asset_Useful_Life_Duration = document.getElementById("LBO_Asset_Useful_Life_Duration").value
const LBO_CapEx = document.getElementById("LBO_CapEx").value
const LBO_Subordinated_Debt = document.getElementById("LBO_Subordinated_Debt").value
const LBO_EBITDA_Multiple = document.getElementById("LBO_EBITDA_Multiple").value
const LBO_Purchase_Price = document.getElementById("LBO_Purchase_Price").value
const Subordinated_Debt_EBITDA_Multiple = document.getElementById("Subordinated_Debt_EBITDA_Multiple").value
const Senior_Debt_EBITDA_Multiple = document.getElementById("Senior_Debt_EBITDA_Multiple").value
const LBO_Transaction_Expense = document.getElementById("LBO_Transaction_Expense").value
const LBO_Financing_Fees = document.getElementById("LBO_Financing_Fees").value
const createFinancialModel = document.getElementById("createFinancialModel")

const revenue_growth_rate = ((Current_Revenue / Historical_Revenue) - 1).toFixed(2)
const projectedRevenueYr1 = (revenue_growth_rate + 1) * Current_Revenue
const projectedRevenueYr2 = (revenue_growth_rate + 1) * projectedRevenueYr1
const projectedRevenueYr3 = (revenue_growth_rate + 1) * projectedRevenueYr2
const projectedRevenueYr4 = (revenue_growth_rate + 1) * projectedRevenueYr3
const projectedRevenueYr5 = (revenue_growth_rate + 1) * projectedRevenueYr4

const projectedRevnues = {
    "revenue_growth_rate": revenue_growth_rate,
    "current_revenue": Current_Revenue,
    "historical_revenue": Historical_Revenue,
    "year_1": projectedRevenueYr1,
    "year_2": projectedRevenueYr2,
    "year_3": projectedRevenueYr3,
    "year_4": projectedRevenueYr4,
    "year_5": projectedRevenueYr5
}

const cogsRatio = (Current_COGS / Historical_Revenue).toFixed(2)
const projectedCOGSYr1 = cogsRatio * Current_Revenue
const projectedCOGSYr2 = cogsRatio * projectedRevenueYr1
const projectedCOGSYr3 = cogsRatio * projectedRevenueYr2
const projectedCOGSYr4 = cogsRatio * projectedRevenueYr3
const projectedCOGSYr5 = cogsRatio * projectedRevenueYr4

const projectedCOGS = {
    "cogs_ratio": cogsRatio,
    "current_cogs": Current_COGS,
    "historical_cogs": Historical_COGS,
    "year_1": projectedCOGSYr1,
    "year_2": projectedCOGSYr2,
    "year_3": projectedCOGSYr3,
    "year_4": projectedCOGSYr4,
    "year_5": projectedCOGSYr5
}

const historicalGrossProfit = Historical_Revenue - Historical_COGS
const currentGrossProfit = Current_Revenue - Current_COGS
const projectedGrossProfitYr1 = projectedRevenueYr1 - projectedCOGSYr1
const projectedGrossProfitYr2 = projectedRevenueYr2 - projectedCOGSYr2
const projectedGrossProfitYr3 = projectedRevenueYr3 - projectedCOGSYr3
const projectedGrossProfitYr4 = projectedRevenueYr4 - projectedCOGSYr4
const projectedGrossProfitYr5 = projectedRevenueYr5 - projectedCOGSYr5

const grossRatioHistorical = historicalGrossProfit / Historical_Revenue
const grossRatioCurrent = currentGrossProfit / Current_Revenue
const projectedGrossRatioYr1 = projectedGrossProfitYr1 / projectedRevenueYr1
const projectedGrossRatioYr2 = projectedGrossProfitYr2 / projectedRevenueYr2
const projectedGrossRatioYr3 = projectedGrossProfitYr3 / projectedRevenueYr3
const projectedGrossRatioYr4 = projectedGrossProfitYr4 / projectedRevenueYr4
const projectedGrossRatioYr5 = projectedGrossProfitYr5 / projectedRevenueYr5

const projectedGrossRevenues = {
    "historical_gross_profit": historicalGrossProfit,
    "current_gross_profit": currentGrossProfit,
    "year_1": projectedGrossProfitYr1,
    "year_2": projectedGrossProfitYr2,
    "year_3": projectedGrossProfitYr3,
    "year_4": projectedGrossProfitYr4,
    "year_5": projectedGrossProfitYr5,
    "historical_gross_profit_ratio": grossRatioHistorical,
    "current_gross_profit_ratio": grossRatioCurrent,
    "year_1_gross_profit_ratio": projectedGrossRatioYr1,
    "year_2_gross_profit_ratio": projectedGrossRatioYr2,
    "year_3_gross_profit_ratio": projectedGrossRatioYr3,
    "year_4_gross_profit_ratio": projectedGrossRatioYr4,
    "year_5_gross_profit_ratio": projectedGrossRatioYr5
}

const sgaRatio = (Current_SGA / Historical_Revenue).toFixed(2) 
const projectedSGAYr1 = sgaRatio * Current_Revenue
const projectedSGAYr2 = sgaRatio * projectedRevenueYr1
const projectedSGAYr3 = sgaRatio * projectedRevenueYr2
const projectedSGAYr4 = sgaRatio * projectedRevenueYr3
const projectedSGAYr5 = sgaRatio * projectedRevenueYr4

const sgaHistoricalRatio = Historical_SGA / Historical_Revenue
const sgaCurrentRatio = Current_SGA / Current_Revenue
const projectedSGARatioYr1 = projectedSGAYr1 / projectedRevenueYr1
const projectedSGARatioYr2 = projectedSGAYr2 / projectedRevenueYr2
const projectedSGARatioYr3 = projectedSGAYr3 / projectedRevenueYr3
const projectedSGARatioYr4 = projectedSGAYr4 / projectedRevenueYr4
const projectedSGARatioYr5 = projectedSGAYr5 / projectedRevenueYr5

const projectedSGA = {
    "sga_ratio": sgaRatio,
    "current_sga": Current_SGA,
    "historical_sga": Historical_SGA,
    "year_1": projectedSGAYr1,
    "year_2": projectedSGAYr2,
    "year_3": projectedSGAYr3,
    "year_4": projectedSGAYr4,
    "year_5": projectedSGAYr5,
    "historical_sga_ratio": sgaHistoricalRatio,
    "current_sga_ratio": sgaCurrentRatio,
    "current_sga_ratio": sgaCurrentRatio,
    "year_1_sga_ratio": projectedSGARatioYr1,
    "year_2_sga_ratio": projectedSGARatioYr2,
    "year_3_sga_ratio": projectedSGARatioYr3,
    "year_4_sga_ratio": projectedSGARatioYr4,
    "year_5_sga_ratio": projectedSGARatioYr5
}

const EbitdaHistorical = historicalGrossProfit - Historical_SGA
const EbitdaCurrent = currentGrossProfit - Current_SGA
const EbitdaYr1 = projectedGrossProfitYr1 - projectedSGAYr1
const EbitdaYr2 = projectedGrossProfitYr2 - projectedSGAYr2
const EbitdaYr3 = projectedGrossProfitYr3 - projectedSGAYr3
const EbitdaYr4 = projectedGrossProfitYr4 - projectedSGAYr4
const EbitdaYr5 = projectedGrossProfitYr5 - projectedSGAYr5

const EbitdaRatioHistorical = EbitdaHistorical / Historical_Revenue
const EbitdaRatioCurrent = EbitdaCurrent / Current_Revenue
const EbitdaRatioYr1 = EbitdaYr1 / projectedRevenueYr1
const EbitdaRatioYr2 = EbitdaYr2 / projectedRevenueYr2
const EbitdaRatioYr3 = EbitdaYr3 / projectedRevenueYr3
const EbitdaRatioYr4 = EbitdaYr4 / projectedRevenueYr4
const EbitdaRatioYr5 = EbitdaYr5 / projectedRevenueYr5

const projectedEbitda = {
    "historical_ebitda": EbitdaHistorical,
    "current_ebitda": EbitdaCurrent,
    "year_1": EbitdaYr1,
    "year_2": EbitdaYr2,
    "year_3": EbitdaYr3,
    "year_4": EbitdaYr4,
    "year_5": EbitdaYr5,
    "historical_ebitda_ratio": EbitdaRatioHistorical,
    "current_ebitda_ratio": EbitdaRatioCurrent,
    "year_1_ebitda_ratio": EbitdaRatioYr1,
    "year_2_ebitda_ratio": EbitdaRatioYr2,
    "year_3_ebitda_ratio": EbitdaRatioYr3,
    "year_4_ebitda_ratio": EbitdaRatioYr4,
    "year_5_ebitda_ratio": EbitdaRatioYr5
}

//! To do
// Depreciation
// Historical_Depreciation
// Current_Depreciation

// Amortization
// Historical_Amortization
// Current_Amortization

// ! To do
// Interest expense
// Historical_Interest_Expense
// Current_Interest_Expense

// EBIT
const EbitHistorical = EbitdaHistorical - Historical_Depreciation - Historical_Amortization
const EbitCurrent = EbitdaCurrent - Current_Depreciation - Current_Amortization
// ! To do


// Pretax income
const PreTaxIncomeHistorical = EbitHistorical - Historical_Interest_Expense
const PreTaxIncomeCurrent = EbitCurrent - Current_Interest_Expense
// ! To do

// Income tax expense
const IncomeTaxExpenseHistorical = PreTaxIncomeHistorical * 0.35
const IncomeTaxExpenseCurrent = PreTaxIncomeCurrent * 0.35
// ! To do

// Net income
const NetIncomeHistorical = PreTaxIncomeHistorical - IncomeTaxExpenseHistorical
const NetIncomeCurrent = PreTaxIncomeCurrent - IncomeTaxExpenseCurrent
// ! To do

const NetIncome = {
    "historical_net_income": NetIncomeHistorical,
    "current_net_income": NetIncomeCurrent
}

// Balance Sheet
const newCompanyRevolvingCredit = document.getElementById("").value
const newCompanySubordinatedDebt = document.getElementById("").value
const newCompanySeniorDebt = document.getElementById("").value


const TotalCurrentAssets = Historical_Cash_Balance + LBO_Accounts_Receivable + LBO_Inventory +LBO_Prepaid_Expenses
const TotalOtherAssets = Capitalized_Financing_Fee + LBO_Goodwill + LBO_Other_Long_Term_Asset
const TotalAssets = TotalCurrentAssets + TotalOtherAssets
const TotalCurrentLiabilities = LBO_Accounts_Payable
const LongTermLiabilities = LBO_Long_Term_Debt + LBO_Notes_Payable
const TotalLiabilities = TotalCurrentLiabilities + LongTermLiabilities

const TotalEquity = LBO_Common_Stock + LBO_Retained_Earnings
const TotalLiabilitiesAndEquity = TotalLiabilities + TotalEquity

const ARDays = LBO_Accounts_Receivable / (Historical_Revenue / 365)
const InventoryDays = LBO_Inventory / (Historical_COGS / 365)
const APDays = LBO_Accounts_Payable / (Historical_COGS / 365)
// Cash Flow statement
// !To do

// Debt schedule
// !To do

// Sources and Uses
const EbitdaAtClose = EbitdaCurrent
const EnterpriseValMultiple = LBO_Purchase_Price / EbitdaAtClose
const subordinatedDebtVal = Subordinated_Debt_EBITDA_Multiple * EbitdaAtClose
const seniorDebtVal = Senior_Debt_EBITDA_Multiple * EbitdaAtClose
const sourcesTotal = LBO_Purchase_Price + (LBO_Transaction_Expense + LBO_Financing_Fees)
const sourcesEquity = sourcesTotal - (subordinatedDebtVal + seniorDebtVal)

const oldCompanyDebt = LBO_Notes_Payable + LBO_Long_Term_Debt
const sellerProceeds = sourcesTotal - (oldCompanyDebt + LBO_Transaction_Expense + LBO_Financing_Fees)
const usesTotal = sourcesTotal

// Balance sheet adjustments
const sourceEquityVal = sourcesEquity
const sourceDebtVal = subordinatedDebtVal + seniorDebtVal
const useOldCompanyEquity = -sellerProceeds
const useOldCompanyDebt = -oldCompanyDebt 
const useTransactionFees = -LBO_Transaction_Expense
const useFinancingFees = -LBO_Financing_Fees
// PostClosingVals
}




function three_statement_workbook(){
    // Three Statement Model
// Income statement
const three_state_company_name = document.getElementById("three_statement_company_name").value

const three_state_revenue = document.getElementById("three_statement_revenue").value
const three_state_cogs = document.getElementById("three_statement_COGS").value
const three_state_marketing_exp = document.getElementById("three_statement_marketing").value
const three_state_SGA = document.getElementById("three_statement_SG&A").value
const three_state_depreciation = document.getElementById("three_statement_depreciation").value
const three_state_amortization = document.getElementById("three_statement_amortization").value
const three_state_interest_exp = document.getElementById("three_statement_interest_expense").value

const grossProfit = three_state_revenue - three_state_cogs
const operatingProfit = grossProfit - three_state_marketing_exp - three_state_SGA
const EBIT = operatingProfit - three_state_depreciation - three_state_amortization
const three_state_tax_expense = 0.35 * EBIT
const three_state_net_income = EBIT - three_state_tax_expense - three_state_interest_exp

const currencySelector = document.getElementById("three_statement_currency");
const selectedCurrency = currencySelector.options[currencySelector.selectedIndex].value;


const three_state_income_data = [
    ["Income Statement",],
    [,],
    ["Revenue", {t: "n", v: three_state_revenue, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`}],
    ["Cost of Goods Sold", {t: "n", v: three_state_cogs, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`},],
    [,],
    ["Gross Income", {t: "n", v: grossProfit, f: "B3 - B4", z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`},],
    [,],
    ["Selling, General & Administrative", {t: "n", v: three_state_SGA, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`},],
    ["Marketing Expense", {t: "n", v: three_state_marketing_exp, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`},],
    ["Operating Income", {t: "n", v: operatingProfit, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`}
,],
    ["Depreciation", {t: "n", v: three_state_depreciation, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`}
,],
    ["Amortization", {t: "n", v: three_state_amortization, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`}
,],
    [, ],
    ["EBIT", {t: "n", v: EBIT, f:"B6 - (B8 + B9 + B10 + B11 + B12)", z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`},],
    [,],
    ["Tax Expense", {t: "n", v: three_state_tax_expense, f:"B14 * 0.35", z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`},],
    ["Interest Expense", {t: "n", v: three_state_interest_exp, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`}
,],
    [,],
    [,],
    ["Net Income", {t: "n", v: three_state_net_income, f: "B14 - (B16 + B17)", z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`},]
]

//balance sheet
const three_state_cash_balance = document.getElementById("three_statement_cash_balance").value
const three_state_accounts_receivable = document.getElementById("three_statement_accounts_receivable").value
const three_state_inventory = document.getElementById("three_statement_inventory").value
const three_state_prepaid_expenses = document.getElementById("three_statement_prepaid_expenses").value
const three_state_goodwill = document.getElementById("three_statement_goodwill").value
const three_state_other_long_term_assets = document.getElementById("three_statement_other_long_term_assets").value
const three_state_accounts_payable = document.getElementById("three_statement_accounts_payable").value
const three_state_short_term_loans = document.getElementById("three_statement_short_term_loans").value
const three_state_accrued_comp = document.getElementById("three_statement_accrued_comp").value
const three_state_notes_payable = document.getElementById("three_statement_notes_payable").value
const three_state_long_term_debt = document.getElementById("three_statement_long_term_debt").value
const three_state_common_stock = document.getElementById("three_statement_common_stock").value
const three_state_retained_earnings = document.getElementById("three_statement_retained_earnings").value

const three_state_current_assets = three_state_cash_balance + three_state_accounts_receivable + three_state_inventory + three_state_prepaid_expenses
const three_state_total_current_assets = three_state_current_assets + three_state_goodwill + three_state_other_long_term_assets
const three_state_current_liabilties = three_state_accounts_payable + three_state_short_term_loans + three_state_accrued_comp
const three_state_total_liabilities = three_state_current_liabilties + three_state_long_term_debt + three_state_notes_payable
const three_state_shareholder_equity = three_state_common_stock + three_state_retained_earnings




const three_state_balance_sheet_data = [
    ["Balance Sheet",],
    [,],
    ["Cash Balance", {t: "n", v: three_state_cash_balance, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`}
,],
    ["Accounts Receivable", {t: "n", v: three_state_accounts_receivable, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`}
,],
    ["Inventory", {t: "n", v: three_state_inventory, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`}
,],
    ["Prepaid Expenses", {t: "n", v: three_state_prepaid_expenses, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`}
,],
    [,],
["Total Current Assets", {t: "n", v: three_state_current_assets, f: "B3 + B4 + B5 + B6", z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`},],
    [,],
    ["Goodwill", {t: "n", v: three_state_goodwill, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`}
,],
    ["Long Term Assets", {t: "n", v: three_state_other_long_term_assets, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`}
,],
    [, ],
["Total Assets", {t: "n", v: three_state_total_current_assets, f:"B8 + (B10 + B11)", z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`},],
    [,],
    ["Accounts Payable", {t: "n", v: three_state_accounts_payable, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`}
,],
    ["Short Term Loans", {t: "n", v: three_state_short_term_loans, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`}
,],
    ["Accrued Compensation and Benefits", {t: "n", v: three_state_accrued_comp, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`}
,],
    [,],
["Current Liabilities", {t: "n", v: three_state_current_liabilties, f:"B15 + B17", z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`},],
    [,],
    ["Notes Payable", {t: "n", v: three_state_notes_payable, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`}
,],
    ["Long Term Debt", {t: "n", v: three_state_long_term_debt, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`}
,],
    [,],
["Total Liabilities", {t: "n", v: three_state_total_liabilities, f:"B19 + (B21 + B22)", z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`},],
    [,],
    ["Common Stock", {t: "n", v: three_state_common_stock, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`}
,],
    ["Retained Earnings", {t: "n", v: three_state_retained_earnings, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`}
,],
    [,],
["Shareholders' Equity", {t: "n", v: three_state_shareholder_equity, f:"B26 + B27", z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`},]
]

// Cash Flow Statement
const cash_flow_net_income = document.getElementById("cash_flow_net_income").value
const three_statement_cash_flow_depreciation = document.getElementById("three_statement_cash_flow_depreciation").value
const three_statement_cash_flow_deferred_income_tax = document.getElementById("three_statement_cash_flow_deferred_income_tax").value
const three_statement_cash_flow_other_operating = document.getElementById("three_statement_cash_flow_other_operating").value
const three_statement_cash_flow_accounts_receivable = document.getElementById("three_statement_cash_flow_accounts_receivable").value
const three_statement_cash_flow_inventories = document.getElementById("three_statement_cash_flow_inventories").value
const three_statement_cash_flow_other_assets = document.getElementById("three_statement_cash_flow_other_assets").value
const three_statement_cash_flow_accounts_payable = document.getElementById("three_statement_cash_flow_accounts_payable").value
const three_statement_cash_flow_loan_interest = document.getElementById("three_statement_cash_flow_loan_interest").value
const three_statement_cash_flow_other_liabilities = document.getElementById("three_statement_cash_flow_other_liabilities").value
const three_statement_cash_flow_marketable_securities = document.getElementById("three_statement_cash_flow_marketable_securities").value
const three_statement_cash_flow_marketable_securities_maturities = document.getElementById("three_statement_cash_flow_marketable_securities_maturities").value
const three_statement_cash_flow_marketable_securities_sales = document.getElementById("three_statement_cash_flow_marketable_securities_sales").value
const three_statement_cash_flow_aquisition_ppe = document.getElementById("three_statement_cash_flow_aquisition_ppe").value
const three_statement_cash_flow_aquisition_intangible = document.getElementById("three_statement_cash_flow_aquisition_intangible").value
const three_statement_cash_flow_sale_ppe = document.getElementById("three_statement_cash_flow_sale_ppe").value
const three_statement_cash_flow_sale_intangible = document.getElementById("three_statement_cash_flow_sale_intangible").value
const three_statement_cash_flow_other_investing = document.getElementById("three_statement_cash_flow_other_investing").value
const three_statement_cash_flow_dividend = document.getElementById("three_statement_cash_flow_dividend").value
const three_statement_cash_flow_stock_repurchase = document.getElementById("three_statement_cash_flow_stock_repurchase").value
const three_statement_cash_flow_stock_cash_from_loans = document.getElementById("three_statement_cash_flow_stock_cash_from_loans").value
const three_statement_cash_flow_loan_principal = document.getElementById("three_statement_cash_flow_loan_principal").value
const three_statement_cash_flow_other_financing = document.getElementById("three_statement_cash_flow_other_financing").value


const three_state_cash_from_operating = cash_flow_net_income + three_statement_cash_flow_depreciation + three_statement_cash_flow_deferred_income_tax + three_statement_cash_flow_other_operating + three_statement_cash_flow_accounts_receivable + three_statement_cash_flow_inventories + three_statement_cash_flow_accounts_payable + three_statement_cash_flow_other_liabilities + three_statement_cash_flow_loan_interest + three_statement_cash_flow_other_assets
const three_state_cash_from_investing =  three_statement_cash_flow_marketable_securities_maturities + three_statement_cash_flow_marketable_securities_sales - three_statement_cash_flow_marketable_securities - three_statement_cash_flow_aquisition_ppe - three_statement_cash_flow_aquisition_intangible + three_statement_cash_flow_sale_ppe + three_statement_cash_flow_sale_intangible - three_statement_cash_flow_other_investing
const three_state_cash_from_financing = three_statement_cash_flow_stock_cash_from_loans - three_statement_cash_flow_dividend - three_statement_cash_flow_stock_repurchase - three_statement_cash_flow_loan_principal + three_statement_cash_flow_other_financing


const three_state_cash_flow_data = [
    ["Cash Flow",],
    [,],
    ["Net Income", {t: "n", v: cash_flow_net_income, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`}
,],
    ["Depreciation", {t: "n", v: three_statement_cash_flow_depreciation, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`}
,],
["Deferred Income Tax", {t: "n", v: three_statement_cash_flow_deferred_income_tax, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`}
,],
    ["Other", {t: "n", v: three_statement_cash_flow_other_operating, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`}
,],
    ["Accounts Receivable", {t: "n", v: three_statement_cash_flow_accounts_receivable, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`}
,],
    ["Inventories", {t: "n", v: three_statement_cash_flow_inventories, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`}
,],
    ["Other Current and Non-Current Assets", {t: "n", v: three_statement_cash_flow_other_assets, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`}
,],
    ["Accounts Payable", {t: "n", v: three_statement_cash_flow_accounts_payable, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`}
,],
    ["Loan Interest Payment", {t: "n", v: three_statement_cash_flow_loan_interest, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`}
,],
    ["Other Current and Non-Current Liabilities", {t: "n", v: three_statement_cash_flow_other_assets, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`}
,],
    [,],
    ["Cash From Operations", {t: "n", v: three_state_cash_from_operating, f: "B3 + B4 + B5 + B6 + B7 + B8 + B9 + B10 + B11 + B12", z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`},],
    [,],
    ["Purchases of Marketable Securities", {t: "n", v: three_statement_cash_flow_marketable_securities, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`}
,],
    ["Proceeds From Maturities of Marketable Securities", {t: "n", v: three_statement_cash_flow_marketable_securities_maturities, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`}
,],
    ["Proceeds From Sales of Marketable Securities", {t: "n", v: three_statement_cash_flow_marketable_securities_sales, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`}
,],
    ["Acquisition of Property, Plant and Equipment", {t: "n", v: three_statement_cash_flow_aquisition_ppe, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`}
,],
    ["Acquisition of Intangible Assets", {t: "n", v: three_statement_cash_flow_aquisition_intangible, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`}
,],
    ["Proceeds from the Sale of Property, Plant and Equipment", {t: "n", v: three_statement_cash_flow_sale_ppe, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`}
,],
    ["Proceeds from Sale of Intangible Assets", {t: "n", v: three_statement_cash_flow_sale_intangible, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`}
,],
    ["Other Investing Activities", {t: "n", v: three_statement_cash_flow_other_investing, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`}
,],
    [,],
["Cash From Investing Activities", {t: "n", v: three_state_cash_from_investing, f:"B16 + B17 + B18 + B19 + B20 + B21 + B22 + B23", z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`},],
    [,],
    ["Dividends and Dividend Equivalent Rights Paid", {t: "n", v: three_statement_cash_flow_dividend, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`}
,],
    ["Common Stock Repurchase", {t: "n", v: three_statement_cash_flow_stock_repurchase, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`}
,],
    ["Proceeds from Issuance of Long-Term Debt", {t: "n", v: three_statement_cash_flow_stock_cash_from_loans, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`}
,],
    ["Loan Principal Payment", {t: "n", v: three_statement_cash_flow_loan_principal, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`}
,],
    ["Other Financing Activities", {t: "n", v: three_statement_cash_flow_other_financing, z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`}
,],
    [,],    
["Cash used in Financing Activities", {t: "n", v: three_state_cash_from_financing, f:"B27 + B28 + B29 + B30 + B31", z: `${selectedCurrency === 'Dollar' ? '"$"#,##0.00_);\\("$"#,##0.00\\)': selectedCurrency === 'Rand' ? '"R "#,##0.00_);\\("R "#,##0.00\\)' : selectedCurrency === 'Langeni' ? '"SZL "#,##0.00_);\\("SZL "#,##0.00\\)' : '"$"#,##0.00_);\\("$"#,##0.00\\)'}`},],
    [,]
]
    const income_state_ws = XLSX.utils.aoa_to_sheet(three_state_income_data)
    const balanc_sheet_ws = XLSX.utils.aoa_to_sheet( three_state_balance_sheet_data)
    const cash_flow_ws = XLSX.utils.aoa_to_sheet(three_state_cash_flow_data)
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, income_state_ws, "Income Statement")
    XLSX.utils.book_append_sheet(workbook, balanc_sheet_ws, "Balance Sheet")
    XLSX.utils.book_append_sheet(workbook, cash_flow_ws, "Cash Flow")
    XLSX.writeFile(workbook, `Three-statement-model-${three_state_company_name}.xlsx`)
}

document.getElementById("createThreeStatement").addEventListener("click", (event)=>{
    event.preventDefault()
    three_statement_workbook()
})


const createValuationSheet=()=>{
    // Discounted Cash Flow/Valuation
//Income statement
const Dcf_Historical_Revenue = document.getElementById("Dcf_Historical_Revenue").value
const Dcf_Current_Revenue = document.getElementById("Dcf_Current_Revenue").value
const Dcf_Historical_COGS = document.getElementById("Dcf_Historical_COGS").value
const Dcf_Current_COGS = document.getElementById("Dcf_Current_COGS").value
const Dcf_Historical_SGA = document.getElementById("Dcf_Historical_SG&A").value
const Dcf_Current_SGA = document.getElementById("Dcf_Current_SG&A").value
const Dcf_Historical_marketing = document.getElementById("Dcf_Historical_marketing").value
const Dcf_Current_marketing = document.getElementById("Dcf_Current_marketing").value
const Dcf_Historical_Depreciation = document.getElementById("Dcf_Historical_Depreciation").value
const Dcf_Current_Depreciation = document.getElementById("Dcf_Current_Depreciation").value
const Dcf_Historical_Amortization = document.getElementById("Dcf_Historical_Amortization").value
const Dcf_Current_Amortization = document.getElementById("Dcf_Current_Amortization").value
const Dcf_Historical_Interest_Expense = document.getElementById("Dcf_Historical_Interest_Expense").value
const Dcf_Current_Interest_Expense = document.getElementById("Dcf_Current_Interest_Expense").value

const dcf_revenue_delta = ((Dcf_Current_Revenue - Dcf_Historical_Revenue) / Dcf_Historical_Revenue)
const dcf_cogs_delta = ((Dcf_Current_COGS - Dcf_Historical_COGS) / Dcf_Historical_COGS)
const dcf_sga_delta = ((Dcf_Current_SGA - Dcf_Historical_SGA) / Dcf_Historical_SGA)
const dcf_marketing_delta = ((Dcf_Current_marketing - Dcf_Historical_marketing) / Dcf_Historical_marketing)
const dcf_depreciation_delta = ((Dcf_Current_Depreciation - Dcf_Historical_Depreciation) / Dcf_Historical_Depreciation)
const dcf_amortization_delta = ((Dcf_Current_Amortization - Dcf_Historical_Amortization) / Dcf_Historical_Amortization)
const dcf_interest_expense_delta = ((Dcf_Current_Interest_Expense - Dcf_Historical_Interest_Expense) / Dcf_Historical_Interest_Expense)

const dcf_gross_profit = Dcf_Current_Revenue - Dcf_Current_COGS
const dcf_operating_profit = dcf_gross_profit - Dcf_Current_SGA
const dcf_ebit = dcf_operating_profit - Dcf_Current_Depreciation - Dcf_Current_Amortization
const dcf_tax_expense = dcf_ebit * 0.35
const dcf_net_profit = dcf_ebit - dcf_tax_expense - Dcf_Current_Interest_Expense


// Projections
const projected_revenue_1 = (1 + dcf_revenue_delta) * Dcf_Current_Revenue
const projected_revenue_2 = (1 + dcf_revenue_delta) * projected_revenue_1
const projected_revenue_3 = (1 + dcf_revenue_delta) * projected_revenue_2
const projected_revenue_4 = (1 + dcf_revenue_delta) * projected_revenue_3
const projected_revenue_5 = (1 + dcf_revenue_delta) * projected_revenue_4

const projected_cogs_1 = (1 + dcf_cogs_delta) * Dcf_Current_COGS
const projected_cogs_2 = (1 + dcf_cogs_delta) * projected_cogs_1
const projected_cogs_3 = (1 + dcf_cogs_delta) * projected_cogs_2
const projected_cogs_4 = (1 + dcf_cogs_delta) * projected_cogs_3
const projected_cogs_5 = (1 + dcf_cogs_delta) * projected_cogs_4

const projected_SGA_1 = (1 + dcf_sga_delta) * Dcf_Current_SGA
const projected_SGA_2 = (1 + dcf_sga_delta) * projected_SGA_1
const projected_SGA_3 = (1 + dcf_sga_delta) * projected_SGA_2
const projected_SGA_4 = (1 + dcf_sga_delta) * projected_SGA_3
const projected_SGA_5 = (1 + dcf_sga_delta) * projected_SGA_4

const projected_marketing_1 = (1 + dcf_marketing_delta) * Dcf_Current_marketing
const projected_marketing_2 = (1 + dcf_marketing_delta) * projected_marketing_1
const projected_marketing_3 = (1 + dcf_marketing_delta) * projected_marketing_2
const projected_marketing_4 = (1 + dcf_marketing_delta) * projected_marketing_3
const projected_marketing_5 = (1 + dcf_marketing_delta) * projected_marketing_4

const projected_depreciation_1 = (1 + dcf_depreciation_delta) * Dcf_Current_Depreciation
const projected_depreciation_2 = (1 + dcf_depreciation_delta) * projected_depreciation_1
const projected_depreciation_3 = (1 + dcf_depreciation_delta) * projected_depreciation_2
const projected_depreciation_4 = (1 + dcf_depreciation_delta) * projected_depreciation_3
const projected_depreciation_5 = (1 + dcf_depreciation_delta) * projected_depreciation_4

const projected_amortization_1 = (1 + dcf_amortization_delta) * Dcf_Current_Amortization
const projected_amortization_2 = (1 + dcf_amortization_delta) * projected_amortization_1
const projected_amortization_3 = (1 + dcf_amortization_delta) * projected_amortization_2
const projected_amortization_4 = (1 + dcf_amortization_delta) * projected_amortization_3
const projected_amortization_5 = (1 + dcf_amortization_delta) * projected_amortization_4

//! Figure out Interest Expense projections
// const projected_interest_expense_1 = (1 + dcf_interest_expense_delta) * Dcf_Current_Interest_Expense
// Balance sheet
const Dcf_Cash_Balance = document.getElementById("Dcf_Cash_Balance").value
const Dcf_Accounts_Receivable = document.getElementById("Dcf_Accounts_Receivable").value
const Dcf_Inventory = document.getElementById("Dcf_Inventory").value
const Dcf_Prepaid_Expenses = document.getElementById("Dcf_Prepaid_Expenses").value
const Dcf_Goodwill = document.getElementById("Dcf_goodwill").value
const Dcf_historical_PPE = document.getElementById("Dcf_historical_PPE").value
const Dcf_PPE = document.getElementById("Dcf_PPE").value
const Dcf_Other_Long_Term_Assets = document.getElementById("Dcf_long_term_assets").value
const Dcf_Accounts_Payable = document.getElementById("Dcf_accounts_payable").value
const Dcf_Short_Term_Loans = document.getElementById("Dcf_short_term_loans").value
const Dcf_accrued_comp = document.getElementById("Dcf_accrued_comp").value
const Dcf_Notes_Payable = document.getElementById("Dcf_Notes_Payable").value
const Dcf_Long_Term_Debt = document.getElementById("Dcf_Long_Term_Debt").value
const Dcf_Common_Stock = document.getElementById("Dcf_Common_Stock").value
const Dcf_Retained_Earnings = document.getElementById("Dcf_Retained_Earnings").value

const cce = Dcf_Cash_Balance + dcf_net_profit + Dcf_Inventory + Dcf_Prepaid_Expenses
const cce_delta = ((cce - Dcf_Cash_Balance) / Dcf_Cash_Balance)

const projected_cce_1 = cce
const projected_cce_2 = (1 + cce_delta) * cce
const projected_cce_3 = (1 + cce_delta) * projected_cce_2
const projected_cce_4 = (1 + cce_delta) * projected_cce_3
const projected_cce_5 = (1 + cce_delta) * projected_cce_4

const AR_projection_days = Dcf_Accounts_Receivable / (Dcf_Current_Revenue / 365)
const projected_AR_1 = AR_projection_days * (projected_revenue_1 / 365)
const projected_AR_2 = AR_projection_days * (projected_revenue_2 / 365)
const projected_AR_3 = AR_projection_days * (projected_revenue_3 / 365)
const projected_AR_4 = AR_projection_days * (projected_revenue_4 / 365)
const projected_AR_5 = AR_projection_days * (projected_revenue_5 / 365)

const PPE_RATIO = Dcf_PPE / Dcf_Current_Revenue

const projected_ppe_1 = PPE_RATIO * projected_revenue_1
const projected_ppe_2 = PPE_RATIO * projected_revenue_2
const projected_ppe_3 = PPE_RATIO * projected_revenue_3
const projected_ppe_4 = PPE_RATIO * projected_revenue_4
const projected_ppe_5 = PPE_RATIO * projected_revenue_5


const dcf_current_assets = Dcf_Cash_Balance + Dcf_Accounts_Receivable + Dcf_Inventory + Dcf_Prepaid_Expenses
const dcf_total_assets = dcf_current_assets + Dcf_Goodwill + Dcf_Other_Long_Term_Assets
const dcf_current_liabilities = Dcf_Accounts_Payable + Dcf_Short_Term_Loans + Dcf_accrued_comp
const curr_liability_ratio = dcf_current_liabilities / Dcf_Current_Revenue
const dcf_total_liabilities = Dcf_Notes_Payable + Dcf_Long_Term_Debt + dcf_current_liabilities
const dcf_shareholders_equity = Dcf_Common_Stock + Dcf_Retained_Earnings

//
const projected_accounts_payable_1 = dcf_current_liabilities / ((Dcf_Current_SGA + Dcf_Current_marketing) * 365)
const projected_accounts_payable_2 = (curr_liability_ratio * projected_revenue_1) / ((projected_SGA_1 + projected_marketing_1) * 365)
const projected_accounts_payable_3 = (curr_liability_ratio * projected_revenue_2) / ((projected_SGA_2 + projected_marketing_2) * 365)
const projected_accounts_payable_4 = (curr_liability_ratio * projected_revenue_3) / ((projected_SGA_3 + projected_marketing_3) * 365)
const projected_accounts_payable_5 = (curr_liability_ratio * projected_revenue_4) / ((projected_SGA_4 + projected_marketing_4) * 365)

// CAPEX formula
//Capital Expenditures (Capex) = Ending PP&E – Beginning PP&E + Depreciation
const projected_capeEx_1 =  (projected_ppe_1 - Dcf_PPE) + Dcf_Current_Depreciation
const projected_capeEx_2 =  (projected_ppe_2 - projected_ppe_1) + projected_depreciation_1
const projected_capeEx_3 =  (projected_ppe_3 - projected_ppe_2) + projected_depreciation_2
const projected_capeEx_4 =  (projected_ppe_4 - projected_ppe_3) + projected_depreciation_3
const projected_capeEx_5 =  (projected_ppe_5 - projected_ppe_4) + projected_depreciation_4
// Free Cash flow formula
// FCFF Calculation = EBIT x (1-tax rate) + Non-Cash Charges + Changes in Working capital – Capital Expenditure
const FCFF_0 = dcf_ebit * (1 - 0.35) + (Dcf_Current_Depreciation + Dcf_Current_Amortization) + (dcf_current_assets - dcf_current_liabilities) - ((Dcf_PPE - Dcf_historical_PPE) + Dcf_Historical_Depreciation)
const FCFF_1 = (projected_revenue_1 - (projected_SGA_1 + projected_marketing_1) - (projected_depreciation_1 + projected_amortization_1)) * (1 - 0.35) + (projected_depreciation_1 + projected_amortization_1) + ((projected_AR_1 + projected_cce_1) - (curr_liability_ratio * projected_revenue_1)) - projected_capeEx_1
const FCFF_2 = (projected_revenue_2 - (projected_SGA_2 + projected_marketing_2) - (projected_depreciation_2 + projected_amortization_2)) * (1 - 0.35) + (projected_depreciation_2 + projected_amortization_2) + ((projected_AR_2 + projected_cce_2) - (curr_liability_ratio * projected_revenue_2)) - projected_capeEx_2
const FCFF_3 = (projected_revenue_3 - (projected_SGA_3 + projected_marketing_3) - (projected_depreciation_3 + projected_amortization_3)) * (1 - 0.35) + (projected_depreciation_3 + projected_amortization_3) + ((projected_AR_3 + projected_cce_3) - (curr_liability_ratio * projected_revenue_3)) - projected_capeEx_3
const FCFF_4 = (projected_revenue_4 - (projected_SGA_4 + projected_marketing_4) - (projected_depreciation_4 + projected_amortization_4)) * (1 - 0.35) + (projected_depreciation_4 + projected_amortization_4) + ((projected_AR_4 + projected_cce_4) - (curr_liability_ratio * projected_revenue_4)) - projected_capeEx_4
const FCFF_5 = (projected_revenue_5 - (projected_SGA_5 + projected_marketing_5) - (projected_depreciation_5 + projected_amortization_5)) * (1 - 0.35) + (projected_depreciation_5 + projected_amortization_5) + ((projected_AR_5 + projected_cce_5) - (curr_liability_ratio * projected_revenue_5)) - projected_capeEx_5
// Cash Flow Statement
const Dcf_cash_flow_net_income = document.getElementById("Dcf_cash_flow_net_income").value
const Dcf_cash_flow_depreciation = document.getElementById("Dcf_cash_flow_depreciation").value
const Dcf_cash_flow_deferred_income_tax = document.getElementById("Dcf_cash_flow_deferred_income_tax").value
const Dcf_cash_flow_other_operating = document.getElementById("Dcf_cash_flow_other_operating").value
const Dcf_cash_flow_accounts_receivable = document.getElementById("Dcf_cash_flow_accounts_receivable").value
const Dcf_cash_flow_inventories = document.getElementById("Dcf_cash_flow_inventories").value
const Dcf_cash_flow_other_assets = document.getElementById("Dcf_cash_flow_other_assets").value
const Dcf_cash_flow_accounts_payable = document.getElementById("Dcf_cash_flow_accounts_payable").value
const Dcf_cash_flow_loan_interest = document.getElementById("Dcf_cash_flow_loan_interest").value
const Dcf_cash_flow_aquisition_ppe = document.getElementById("Dcf_cash_flow_aquisition_ppe").value
const Dcf_cash_flow_aquisition_intangible = document.getElementById("Dcf_cash_flow_aquisition_intangible").value
const Dcf_cash_flow_sale_ppe = document.getElementById("Dcf_cash_flow_sale_ppe").value
const Dcf_cash_flow_sale_intangible = document.getElementById("Dcf_cash_flow_sale_intangible").value
const Dcf_cash_flow_other_investing = document.getElementById("Dcf_cash_flow_other_investing").value
const Dcf_cash_flow_dividend = document.getElementById("Dcf_cash_flow_dividend").value
const Dcf_cash_flow_stock_repurchase = document.getElementById("Dcf_cash_flow_stock_repurchase").value
const Dcf_cash_flow_cash_from_loans = document.getElementById("Dcf_cash_flow_cash_from_loans").value

const dcf_cash_from_operating = Dcf_cash_flow_net_income + Dcf_cash_flow_depreciation + Dcf_cash_flow_deferred_income_tax + Dcf_cash_flow_other_operating - Dcf_cash_flow_accounts_receivable - Dcf_cash_flow_inventories + Dcf_cash_flow_other_assets + Dcf_cash_flow_accounts_payable - Dcf_cash_flow_loan_interest
const dcf_cash_from_investing = Dcf_cash_flow_sale_ppe - Dcf_cash_flow_aquisition_ppe - Dcf_cash_flow_aquisition_intangible + Dcf_cash_flow_sale_intangible - Dcf_cash_flow_other_investing
const dcf_cash_from_financing = Dcf_cash_flow_cash_from_loans - Dcf_cash_flow_dividend - Dcf_cash_flow_stock_repurchase


//
const valuation_income_statement_projections = {
    "Revenue": [Dcf_Current_Revenue, projected_revenue_1, projected_revenue_2, projected_revenue_3, projected_revenue_4, projected_revenue_5],
    "COGS": [Dcf_Current_COGS, projected_cogs_1, projected_cogs_2, projected_cogs_3, projected_cogs_4, projected_cogs_5],
    "SGA": [Dcf_Current_SGA, projected_SGA_1, projected_SGA_2, projected_SGA_3, projected_SGA_4, projected_SGA_5],
    "Marketing": [Dcf_Current_marketing, projected_marketing_1, projected_marketing_2, projected_marketing_3, projected_marketing_4, projected_marketing_5],
    "Operating Profit": [Dcf_Current_Revenue - (Dcf_Current_COGS + Dcf_Current_SGA + Dcf_Current_marketing), projected_revenue_1 - (projected_cogs_1 + projected_SGA_1 + projected_marketing_1), projected_revenue_2 - (projected_cogs_2 + projected_SGA_2 + projected_marketing_2), projected_revenue_3 - (projected_cogs_3 + projected_SGA_3 + projected_marketing_3), projected_revenue_4 - (projected_cogs_4 + projected_SGA_4 + projected_marketing_4), projected_revenue_5 - (projected_cogs_5 + projected_SGA_5 + projected_marketing_5)],
    "Depreciation": [Dcf_Current_Depreciation, projected_depreciation_1, projected_depreciation_2, projected_depreciation_3, projected_depreciation_4, projected_depreciation_5],
    "Amortization": [Dcf_Current_Amortization, projected_amortization_1, projected_amortization_2, projected_amortization_3, projected_amortization_4, projected_amortization_5],
    "EBIT": [Dcf_Current_Revenue - (Dcf_Current_COGS + Dcf_Current_SGA + Dcf_Current_marketing + Dcf_Current_Depreciation + Dcf_Current_Amortization), projected_revenue_1 - (projected_cogs_1 + projected_SGA_1 + projected_marketing_1 + projected_depreciation_1 + projected_amortization_1), projected_revenue_2 - (projected_cogs_2 + projected_SGA_2 + projected_marketing_2 + projected_depreciation_2 + projected_amortization_2), projected_revenue_3 - (projected_cogs_3 + projected_SGA_3 + projected_marketing_3 + projected_depreciation_3 + projected_amortization_3), projected_revenue_4 - (projected_cogs_4 + projected_SGA_4 + projected_marketing_4 + projected_depreciation_4 + projected_amortization_4), projected_revenue_5 - (projected_cogs_5 + projected_SGA_5 + projected_marketing_5 + projected_depreciation_5 + projected_amortization_5)],
    "Interest Expense": [Dcf_Current_Interest_Expense, Dcf_Current_Interest_Expense, Dcf_Current_Interest_Expense, Dcf_Current_Interest_Expense, Dcf_Current_Interest_Expense, Dcf_Current_Interest_Expense],
    "Tax Expense": [(Dcf_Current_Revenue - (Dcf_Current_COGS + Dcf_Current_SGA + Dcf_Current_marketing + Dcf_Current_Depreciation + Dcf_Current_Amortization)) * 0.35, (projected_revenue_1 - (projected_cogs_1 + projected_SGA_1 + projected_marketing_1 + projected_depreciation_1 + projected_amortization_1)) * 0.35, (projected_revenue_2 - (projected_cogs_2 + projected_SGA_2 + projected_marketing_2 + projected_depreciation_2 + projected_amortization_2)) * 0.35, (projected_revenue_3 - (projected_cogs_3 + projected_SGA_3 + projected_marketing_3 + projected_depreciation_3 + projected_amortization_3)) * 0.35, (projected_revenue_4 - (projected_cogs_4 + projected_SGA_4 + projected_marketing_4 + projected_depreciation_4 + projected_amortization_4)) * 0.35, (projected_revenue_5 - (projected_cogs_5 + projected_SGA_5 + projected_marketing_5 + projected_depreciation_5 + projected_amortization_5)) * 0.35],
    "Net Income": [(Dcf_Current_Revenue - (Dcf_Current_COGS + Dcf_Current_SGA + Dcf_Current_marketing + Dcf_Current_Depreciation + Dcf_Current_Amortization + Dcf_Current_Interest_Expense)) - ((Dcf_Current_Revenue - (Dcf_Current_COGS + Dcf_Current_SGA + Dcf_Current_marketing + Dcf_Current_Depreciation + Dcf_Current_Amortization)) * 0.35), (projected_revenue_1 - (projected_cogs_1 + projected_SGA_1 + projected_marketing_1 + projected_depreciation_1 + projected_amortization_1 + Dcf_Current_Interest_Expense)) - ((projected_revenue_1 - (projected_cogs_1 + projected_SGA_1 + projected_marketing_1 + projected_depreciation_1 + projected_amortization_1)) * 0.35), (projected_revenue_2 - (projected_cogs_2 + projected_SGA_2 + projected_marketing_2 + projected_depreciation_2 + projected_amortization_2 + Dcf_Current_Interest_Expense)) - ((projected_revenue_2 - (projected_cogs_2 + projected_SGA_2 + projected_marketing_2 + projected_depreciation_2 + projected_amortization_2)) * 0.35), (projected_revenue_3 - (projected_cogs_3 + projected_SGA_3 + projected_marketing_3 + projected_depreciation_3 + projected_amortization_3 + Dcf_Current_Interest_Expense)) - ((projected_revenue_3 - (projected_cogs_3 + projected_SGA_3 + projected_marketing_3 + projected_depreciation_3 + projected_amortization_3)) * 0.35), (projected_revenue_4 - (projected_cogs_4 + projected_SGA_4 + projected_marketing_4 + projected_depreciation_4 + projected_amortization_4 + Dcf_Current_Interest_Expense)) - ((projected_revenue_4 - (projected_cogs_4 + projected_SGA_4 + projected_marketing_4 + projected_depreciation_4 + projected_amortization_4)) * 0.35), (projected_revenue_5 - (projected_cogs_5 + projected_SGA_5 + projected_marketing_5 + projected_depreciation_5 + projected_amortization_5 + Dcf_Current_Interest_Expense)) - ((projected_revenue_5 - (projected_cogs_5 + projected_SGA_5 + projected_marketing_5 + projected_depreciation_5 + projected_amortization_5)) * 0.35)]
}

const valuation_balance_sheet_projections = {
    "Cash & Cash Equivalents":[cce, projected_cce_1, projected_cce_2, projected_cce_3, projected_cce_4, projected_cce_5],
    "Accounts Receivable": [Dcf_Accounts_Receivable, projected_AR_1, projected_AR_2, projected_AR_3, projected_AR_4, projected_AR_5],
    "Current Assets": [cce + Dcf_Accounts_Receivable, projected_cce_1 + projected_AR_1, projected_cce_2 + projected_AR_2, projected_cce_3 + projected_AR_3, projected_cce_4 + projected_AR_4, projected_cce_5 + projected_AR_5],
    "Property, Plant & Equipment": [Dcf_PPE, projected_ppe_1, projected_ppe_2, projected_ppe_3, projected_ppe_4, projected_ppe_5],
    "Total Assets": [cce + Dcf_Accounts_Receivable + Dcf_PPE, projected_cce_1 + projected_AR_1 + projected_ppe_1, projected_cce_2 + projected_AR_2 + projected_ppe_2, projected_cce_3 + projected_AR_3 + projected_ppe_3, projected_cce_4 + projected_AR_4 + projected_ppe_4, projected_cce_5 + projected_AR_5 + projected_ppe_5],
    "Accounts Payable": [Dcf_Accounts_Payable, projected_accounts_payable_1, projected_accounts_payable_2, projected_accounts_payable_3, projected_accounts_payable_4, projected_accounts_payable_5],
    "Current Liabilities": [dcf_current_liabilities, ],
    "Long Term Debt": [Dcf_Long_Term_Debt],
    "Total Liabilities": [dcf_total_liabilities],
    "Common Stock": [Dcf_Common_Stock],
    "Retained Earnings": [Dcf_Retained_Earnings],

}
const valuation_cash_statement_projections = {
    "Net Income": [(Dcf_Current_Revenue - (Dcf_Current_COGS + Dcf_Current_SGA + Dcf_Current_marketing + Dcf_Current_Depreciation + Dcf_Current_Amortization + Dcf_Current_Interest_Expense)) - ((Dcf_Current_Revenue - (Dcf_Current_COGS + Dcf_Current_SGA + Dcf_Current_marketing + Dcf_Current_Depreciation + Dcf_Current_Amortization)) * 0.35), (projected_revenue_1 - (projected_cogs_1 + projected_SGA_1 + projected_marketing_1 + projected_depreciation_1 + projected_amortization_1 + Dcf_Current_Interest_Expense)) - ((projected_revenue_1 - (projected_cogs_1 + projected_SGA_1 + projected_marketing_1 + projected_depreciation_1 + projected_amortization_1)) * 0.35), (projected_revenue_2 - (projected_cogs_2 + projected_SGA_2 + projected_marketing_2 + projected_depreciation_2 + projected_amortization_2 + Dcf_Current_Interest_Expense)) - ((projected_revenue_2 - (projected_cogs_2 + projected_SGA_2 + projected_marketing_2 + projected_depreciation_2 + projected_amortization_2)) * 0.35), (projected_revenue_3 - (projected_cogs_3 + projected_SGA_3 + projected_marketing_3 + projected_depreciation_3 + projected_amortization_3 + Dcf_Current_Interest_Expense)) - ((projected_revenue_3 - (projected_cogs_3 + projected_SGA_3 + projected_marketing_3 + projected_depreciation_3 + projected_amortization_3)) * 0.35), (projected_revenue_4 - (projected_cogs_4 + projected_SGA_4 + projected_marketing_4 + projected_depreciation_4 + projected_amortization_4 + Dcf_Current_Interest_Expense)) - ((projected_revenue_4 - (projected_cogs_4 + projected_SGA_4 + projected_marketing_4 + projected_depreciation_4 + projected_amortization_4)) * 0.35), (projected_revenue_5 - (projected_cogs_5 + projected_SGA_5 + projected_marketing_5 + projected_depreciation_5 + projected_amortization_5 + Dcf_Current_Interest_Expense)) - ((projected_revenue_5 - (projected_cogs_5 + projected_SGA_5 + projected_marketing_5 + projected_depreciation_5 + projected_amortization_5)) * 0.35)],
   "Depreciation": [Dcf_Current_Depreciation, projected_depreciation_1, projected_depreciation_2, projected_depreciation_3, projected_depreciation_4, projected_depreciation_5],
    "Accounts Receivable": [Dcf_Accounts_Receivable, projected_AR_1, projected_AR_2, projected_AR_3, projected_AR_4, projected_AR_5],
    "Inventories": [Dcf_cash_flow_inventories],
    "Accounts Payable": [Dcf_Accounts_Payable, projected_accounts_payable_1, projected_accounts_payable_2, projected_accounts_payable_3, projected_accounts_payable_4, projected_accounts_payable_5],
    "Cash From Operating Activities": [dcf_cash_from_operating],
    "Property, Plant & Equipment Acquisition(s)": [Dcf_cash_flow_aquisition_ppe],
    "Property, Plant & Equipment Sale(s)": [Dcf_cash_flow_sale_ppe],
    "Cash From Investment Activities": [dcf_cash_from_investing],
    "Stock Repurchases": [Dcf_cash_flow_stock_repurchase],
    "Dividends and Divdend Rights Equivalents Paid": [Dcf_cash_flow_dividend],
    "Cash From Financing Activities": [dcf_cash_from_financing]
}
    let dcf_income_statement_ws = []
    for(const key in valuation_income_statement_projections){
        dcf_income_statement_ws.push([key, ...valuation_income_statement_projections[key]])
    }
    let dcf_balance_sheet_ws = []
    for(const key in valuation_balance_sheet_projections){
        dcf_balance_sheet_ws.push([key, ...valuation_balance_sheet_projections[key]])
    }
    let dcf_cash_statement_ws = []
    for(const key in valuation_cash_statement_projections){
        dcf_cash_statement_ws.push([key, ...valuation_cash_statement_projections[key]])
    }

    const company_name = document.getElementById("dcf_statement_company_name").value;
    const income_state_ws = XLSX.utils.aoa_to_sheet(dcf_income_statement_ws)
    const balanc_sheet_ws = XLSX.utils.aoa_to_sheet(dcf_balance_sheet_ws)
    const cash_flow_ws = XLSX.utils.aoa_to_sheet(dcf_cash_statement_ws)
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, income_state_ws, "Income Statement")
    XLSX.utils.book_append_sheet(workbook, balanc_sheet_ws, "Balance Sheet")
    XLSX.utils.book_append_sheet(workbook, cash_flow_ws, "Cash Flow")
    XLSX.writeFile(workbook, `Discounted-Cash-Flow-statement-model-${company_name}.xlsx`)
}

document.getElementById("createDCFStatement").addEventListener("click", (event)=>{
    event.preventDefault();
    createValuationSheet();
    console.log("clicked")
})

//Accretion Dilution Model

/*
const acquirer_net = document.getElementById("acquirer_net").value
const target_net = document.getElementById("target_net").value
const acquirer_shares = document.getElementById("acquirer_shares").value
const target_shares = document.getElementById("target_shares").value
const acquirer_EPS = document.getElementById("acquirer_EPS").value
const target_EPS = document.getElementById("target_EPS").value
const acquirer_share_price = document.getElementById("acquirer_share_price").value
const target_share_price = document.getElementById("target_share_price").value
const acquirer_PE = document.getElementById("acquirer_PE").value
const target_PE = document.getElementById("target_PE").value

*/

// Create Accretion model





//

