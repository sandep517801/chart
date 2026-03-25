let data = JSON.parse(localStorage.getItem("orders")) || [];

// Add manual data
function addData(){
  let orders = document.getElementById("orders").value.split(/[\n,]+/);
  let reason = document.getElementById("customReason").value || document.getElementById("reason").value;
  if(!reason) { alert("Please select or enter a reason"); return; }
  let date = document.getElementById("date").value || new Date().toISOString().split("T")[0];

  orders.forEach(o=>{
    if(o.trim()){
      data.push({order:o.trim(), reason, date});
    }
  });

  saveData();
  updateAll();
  clearInputs();
}

function clearInputs() {
  document.getElementById("orders").value = "";
  document.getElementById("customReason").value = "";
  document.getElementById("reason").selectedIndex = 0;
  document.getElementById("date").value = "";
}

// Save to localStorage
function saveData(){
  localStorage.setItem("orders", JSON.stringify(data));
}

// Update everything: table, charts, stats
function updateAll(){
  updateTable();
  updateCharts();
  updateStats();
  updateRTOView();
}

// Update table
function updateTable(){
  let table = document.getElementById("table");
  table.innerHTML = "<tr><th>Order ID</th><th>Reason</th><th>Date</th><th>Action</th></tr>";
  data.forEach((d,i)=>{
    table.innerHTML += `<tr>
      <td>${d.order}</td>
      <td>${d.reason}</td>
      <td>${d.date}</td>
      <td><button onclick="deleteRow(${i})">Delete</button></td>
    </tr>`;
  });
}

function deleteRow(i){
  data.splice(i,1);
  saveData();
  updateAll();
}

// Search
function searchTable(){
  let val = document.getElementById("search").value.toLowerCase();
  let rows = document.querySelectorAll("#table tr");
  rows.forEach((row,i)=>{
    if(i===0) return;
    row.style.display = row.innerText.toLowerCase().includes(val) ? "" : "none";
  });
}

// Download CSV
function downloadCSV(){
  let csv = "Order ID,Reason,Date\n";
  data.forEach(d=>{
    csv += `${d.order},${d.reason},${d.date}\n`;
  });
  let blob = new Blob([csv], {type:'text/csv'});
  let a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "orders.csv";
  a.click();
}

// Parse CSV
function parseCSV(text) {
  const lines = text.trim().split('\n');
  const result = [];
  for(let i=0;i<lines.length;i++){
    const row = lines[i].split(',');
    if(row.length >=2){
      result.push({
        order: row[0].trim(),
        reason: row[1].trim(),
        date: row[2]?row[2].trim():new Date().toISOString().split('T')[0]
      });
    }
  }
  return result;
}

// Parse Excel
function parseExcel(dataBinary){
  const workbook = XLSX.read(dataBinary,{type:'binary'});
  const firstSheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[firstSheetName];
  const jsonData = XLSX.utils.sheet_to_json(worksheet,{header:1});
  const result=[];
  for(let i=1;i<jsonData.length;i++){
    const row=jsonData[i];
    if(row.length>=2){
      result.push({
        order:String(row[0]).trim(),
        reason:String(row[1]).trim(),
        date: row[2]?String(row[2]).trim():new Date().toISOString().split('T')[0]
      });
    }
  }
  return result;
}

// File input handler
document.getElementById('fileInput').addEventListener('change', function(e){
  const file = e.target.files[0];
  if(!file) return;
  const reader = new FileReader();

  if(file.name.endsWith('.csv')){
    reader.onload=function(event){
      const newOrders=parseCSV(event.target.result);
      data=data.concat(newOrders);
      saveData();
      updateAll();
    }
    reader.readAsText(file);
  } else {
    reader.onload=function(event){
      const binary=event.target.result;
      const newOrders=parseExcel(binary);
      data=data.concat(newOrders);
      saveData();
      updateAll();
    }
    reader.readAsBinaryString(file);
  }
});

// Charts
let barChart, pieChart, rtoChart;

function updateCharts(){
  let counts={};
  data.forEach(d=>counts[d.reason]=(counts[d.reason]||0)+1);
  let labels=Object.keys(counts);
  let values=Object.values(counts);

  if(barChart) barChart.destroy();
  if(pieChart) pieChart.destroy();

  // Top Reasons Pie Chart
  pieChart=new Chart(document.getElementById("pieChart"),{
    type:"pie",
    data:{ labels, datasets:[{data:values, backgroundColor:["#3498db","#f39c12","#2ecc71","#e74c3c","#9b59b6","#1abc9c","#e67e22","#9b59b6"]}]},
    options:{responsive:false}
  });
}

function updateStats(){
  document.getElementById("totalOrders").innerText = data.length;
  let counts={};
  data.forEach(d=>counts[d.reason]=(counts[d.reason]||0)+1);
  let top=Object.keys(counts).reduce((a,b)=>counts[a]>counts[b]?a:b,"-");
  document.getElementById("topReason").innerText=top;
}

function updateRTOView(){
  let counts={}; let total=data.length;
  data.forEach(d=>counts[d.reason]=(counts[d.reason]||0)+1);

  const table=document.getElementById("rtoTable");
  table.innerHTML=`<tr><th>Top RTO Reasons</th><th>In %</th></tr>`;
  const labels=[], values=[];
  for(const reason in counts){
    const percent=((counts[reason]/total)*100).toFixed(2);
    labels.push(reason);
    values.push(percent);
    table.innerHTML+=`<tr><td>${reason}</td><td>${percent}%</td></tr>`;
  }

  const ctx=document.getElementById('rtoPie').getContext('2d');
  if(rtoChart) rtoChart.destroy();
  rtoChart=new Chart(ctx,{
    type:'pie',
    data:{labels, datasets:[{data:values, backgroundColor:["#f39c12","#3498db","#2ecc71","#e74c3c","#9b59b6","#1abc9c","#e67e22"]}]},
    options:{responsive:false, plugins:{legend:{position:'right', labels:{boxWidth:20,padding:15}}, tooltip:{callbacks:{label:ctx=>ctx.label+": "+ctx.parsed+"%"}}}}
  });
}
function clearAllData(){
  if(confirm("Are you sure you want to clear all orders? This cannot be undone.")){
    data = [];
    localStorage.removeItem("orders");
    updateAll();
  }
}

// Initial render
updateAll();