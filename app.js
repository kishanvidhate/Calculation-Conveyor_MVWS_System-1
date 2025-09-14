
// === Auto-generated from Excel ===
const SHEET_DATA = [["Water Demand and BOM Calculations for Conveyor MVWS System ", null, null, null, null, null], ["A", "Editable Inputs", null, null, null, null], [1, "Conveyor Length in meter", 150, null, null, null], [2, "Conveyor Width  in meter", 0.8, null, null, null], [3, "Conveyor No of Belt", 2, null, null, null], ["B", "Advanced inputs (if applicable)", null, null, null, null], [1, "The spacing between two spray nozzles shall not exceed 3 m, as per IS:15325:2020 clause 8.7.4.1", 3, null, null, null], [2, "No of spray nozzles shall consdridied at one location", 2, null, null, null], [3, "The system pressure shall be 1.4 to 3.5 bar as per IS:15325 clause 8.7.3. For better penetration, 2.1 bar pressure shall be considered", 2.1, null, null, null], [4, "No of Linear Heat Sensing Cable is considered for the conveyor, recommended on three sides: top, left, and right.", 3, null, null, null], [5, "Distance from conveyor to deluge valve control panel in meter", 18, null, null, null], [6, "Distance from deluge valve to main hydrant line in meter", 18, null, null, null], ["C", "Calculation Results", null, null, null, null], [1, "Total Area (In sq. meter) Formula = L x W x Number of Belts", "=(C3*C4*C5)", null, null, null], [2, "Design Density 10.2 lpm/m2 as per IS:15325 Clause No 8.7.2", 10.2, null, null, null], [3, "Theoretical Water Requirement [L/Min]", "=(C14*C15)", null, null, null], [4, "Theoretical Water Requirement [m^3/h]", "=(C16*0.06)", null, null, null], [5, "MW Water spray nozzle quantity with 5% extra.", "=((C3/C7)*C8)*1.05", null, null, null], [6, "Linear Heat Sensing Cable quantity in meter with 15% extra.", "=((C3*C10)+C11)*1.15", null, null, null], [7, "Theoretical Flow per Nozzle (L/Min):", "=(C16/C18)", null, null, null], [8, "MVWS Nozzle K-Factor Calculation", "=(C20)/(SQRT(C9))", null, null, null], [9, "MVWS Nozzle K-Factor Selected ", "=IF(C21<=18,18,IF(C21<=22,22,IF(C21<=26,26,IF(C21<=30,30,IF(C21<=35,35,IF(C21<=41,41,IF(C21<=51,51,IF(C21<=64,64,IF(C21<=79,79,91)))))))))", null, null, null], [10, "MVWS Nozzle Spray Angle Selected", "120 Deg", null, null, null], [11, "Actual Flow Based on K-Factor (L/min) = K-Factor × √ Pressure × Number of Nozzles", "=(C22*(SQRT(C9)*C18))", null, null, null], [12, "Actual Flow Provided Based on K-Factor (m^3/h)", "=(C24*0.06)", null, null, null], [13, "Deluge Valve Selection", "=IF(OR(C25>1150,C25<10.01),\"Not Found\",IF(C25>=550.01,\"200 mm\",IF(C25>=200.01,\"150 mm\",IF(C25>=100.01,\"100 mm\",IF(C25>=50.01,\"80 mm\",\"50 mm\")))))", null, null, null], ["D", "Bill of Materials (BOM)", null, null, null, null], ["Sr. No", "Item Description", "Qty", "Unit", "Item No", null], [1, "=C26 & \" Cast Iron Deluge Valve with Wet Pilot Basic Trim Assembly\"", "=IF(C26<>\"\",1,\"\")", "Nos", "=SWITCH(--LEFT(B29,FIND(\" \",B29)-1),200,\"DVCI0006\",150,\"DVCI0001\",100,\"DVCI0002\",80,\"DVCI0003\",50,\"DVCI0004\",\"Not Found\")", null], [2, "=C26 & \" M.S Pipes Heavy C Class As Per IS: 1239 \"", "=(C12)", "Meter", "=SWITCH(--LEFT(B30,FIND(\" \",B30)-1),200,\"PICC005\",150,\"PICC0006\",125,\"PICC00XX\",100,\"PICC0007\",80,\"PICC0008\",65,\"PICC0009\",50,\"PICC0010\",40,\"PICC0011\",32,\"PICC0012\",25,\"PICC0013\",\"Not Found\")", null], [3, "=C26& \" G.I. Pipes Heavy C Class As Per IS: 1239\"", "=CEILING(((C3*70%)+C11)*1.1,6)", "Meter", "=SWITCH(--LEFT(B31,FIND(\" \",B31)-1),200,\"PICC0119\",150,\"PICC0016\",125,\"PICC0089\",100,\"PICC0017\",80,\"PICC0018\",65,\"PICC0019\",50,\"PICC0020\",40,\"PICC0021\",32,\"PICC0022\",25,\"PICC0023\",\"Not Found\")", null], [4, "=IF(B31=\"200 mm G.I. Pipes Heavy C Class As Per IS: 1239\",\"150 mm G.I. Pipes Heavy C Class As Per IS: 1239\",\n   IF(B31=\"150 mm G.I. Pipes Heavy C Class As Per IS: 1239\",\"100 mm G.I. Pipes Heavy C Class As Per IS: 1239\",\n   IF(B31=\"100 mm G.I. Pipes Heavy C Class As Per IS: 1239\",\"80 mm G.I. Pipes Heavy C Class As Per IS: 1239\",\n   IF(B31=\"80 mm G.I. Pipes Heavy C Class As Per IS: 1239\",\"65 mm G.I. Pipes Heavy C Class As Per IS: 1239\",\n   IF(B31=\"50 mm G.I. Pipes Heavy C Class As Per IS: 1239\",\"50 mm G.I. Pipes Heavy C Class As Per IS: 1239\",\"\")))))", "=CEILING(((C3*30%))*1.1,6)", "Meter", "=SWITCH(--LEFT(B32,FIND(\" \",B32)-1),200,\"PICC0119\",150,\"PICC0016\",100,\"PICC0017\",80,\"PICC0018\",65,\"PICC0019\",50,\"PICC0020\",40,\"PICC0021\",32,\"PICC0022\",25,\"PICC0023\",\"Not Found\")", null], [5, "32 mm G.I. Pipes Heavy C Class As Per IS: 1239", "=CEILING(IF(OR(C8=1,C8=2),0,IF(C8=3,0.4,\"\"))*(C3/C7)*1.1,6)", "Meter", "=SWITCH(--LEFT(B33,FIND(\" \",B33)-1),200,\"PICC0119\",150,\"PICC0016\",100,\"PICC0017\",80,\"PICC0018\",65,\"PICC0019\",50,\"PICC0020\",40,\"PICC0021\",32,\"PICC0022\",25,\"PICC0023\",\"Not Found\")", null], [6, "25 mm G.I. Pipes Heavy C Class As Per IS: 1239", "=CEILING((C3/C7)*(C8*1.8)*1.1,6)", "Meter", "=SWITCH(--LEFT(B34,FIND(\" \",B34)-1),200,\"PICC0119\",150,\"PICC0016\",100,\"PICC0017\",80,\"PICC0018\",65,\"PICC0019\",50,\"PICC0020\",40,\"PICC0021\",32,\"PICC0022\",25,\"PICC0023\",\"Not Found\")", null], [7, "=C26 &\" Cast Iron Wafer Type Butterfly Valve\"", "=(C29*4)", "Nos", "=SWITCH(--LEFT(B35,FIND(\" \",B35)-1),200,\"BFCI0003\",150,\"BFCI0004\",100,\"BFCI0005\",80,\"BFCI0006\",65,\"BFCI0007\",50,\"BFCI0008\",\"Not Found\")", null], [8, "=C26&\" MS Y Type Strainers \"", "=(C29*1)", "Nos", "=SWITCH(--LEFT(B36,FIND(\" \",B36)-1),200,\"YSMS0003\",150,\"YSMS0004\",100,\"YSMS0005\",80,\"YSMS0006\",65,\"YSMS0009\",50,\"YSMS0007\",\"Not Found\")", null], [9, "Medium Velocity Water Spray Nozzle Nickel Chrome Plated Brass 1/2\" BSPT", "=(C18)", "Nos", "MVWS0001", null], [10, "Digital Linear Heat Detection Cable Alarm Temperature 70⁰C", "=(C19)", "Nos", "DLHS0002", null], [11, "Deluge Valve Control Panel Outdoor with Canopy and IP65 Protection ", "=(C29*1)", "Nos", "DVCP0013", null], [12, "Pressure Switch with All Accessories. Range : 2-14 kg.", "=(C29*2)", "Nos", "PSKO0001", null], [13, "24 VDC Solenoid Valve, Operating Pressure: 1 - 20 Bar, 1/2\" BSPT", "=(C29*1)", "Nos", "SVDC0013", null], [14, "Monitor Module, if applicable", "=(C29*2)", "Nos", "FAMM0017", null], [15, "Control  Nodule, if applicable", "=(C29*1)", "Nos", "FACM0022", null], [16, "12V - 10 AMPS Battery ", "=(C29*2)", "Nos", "B12V0010", null], [17, "Cables and Accessories", "=(C29*1)", "Lot", "OTHW0010", null], [18, "Other Hardware Like Nut Bolts, U Clamps Anchor Fastener, Flanges & Green Gasket Etc.", "=(C29*1)", "Lot", "OTHW0031", null], [19, "MS Support Made of L Angle, C Channel, & MS Plate Etc.", "=((C30+C31+C32+C33+C34)/3)*3.5", "Kg", "SUPP0060", null]];
const INPUTS = [{"row": 3, "label": "Conveyor Length in meter", "address": "C3", "type": "number", "initial": 150}, {"row": 4, "label": "Conveyor Width  in meter", "address": "C4", "type": "number", "initial": 0.8}, {"row": 5, "label": "Conveyor No of Belt", "address": "C5", "type": "number", "initial": 2}, {"row": 7, "label": "The spacing between two spray nozzles shall not exceed 3 m, as per IS:15325:2020 clause 8.7.4.1", "address": "C7", "type": "number", "initial": 3}, {"row": 8, "label": "No of spray nozzles shall consdridied at one location", "address": "C8", "type": "number", "initial": 2}, {"row": 9, "label": "The system pressure shall be 1.4 to 3.5 bar as per IS:15325 clause 8.7.3. For better penetration, 2.1 bar pressure shall be considered", "address": "C9", "type": "number", "initial": 2.1}, {"row": 10, "label": "No of Linear Heat Sensing Cable is considered for the conveyor, recommended on three sides: top, left, and right.", "address": "C10", "type": "number", "initial": 3}, {"row": 11, "label": "Distance from conveyor to deluge valve control panel in meter", "address": "C11", "type": "number", "initial": 18}, {"row": 12, "label": "Distance from deluge valve to main hydrant line in meter", "address": "C12", "type": "number", "initial": 18}];
const RESULTS = [{"row": 14, "label": "Total Area (In sq. meter) Formula = L x W x Number of Belts", "address": "C14"}, {"row": 15, "label": "Design Density 10.2 lpm/m2 as per IS:15325 Clause No 8.7.2", "address": "C15"}, {"row": 16, "label": "Theoretical Water Requirement [L/Min]", "address": "C16"}, {"row": 17, "label": "Theoretical Water Requirement [m^3/h]", "address": "C17"}, {"row": 18, "label": "MW Water spray nozzle quantity with 5% extra.", "address": "C18"}, {"row": 19, "label": "Linear Heat Sensing Cable quantity in meter with 15% extra.", "address": "C19"}, {"row": 20, "label": "Theoretical Flow per Nozzle (L/Min):", "address": "C20"}, {"row": 21, "label": "MVWS Nozzle K-Factor Calculation", "address": "C21"}, {"row": 22, "label": "MVWS Nozzle K-Factor Selected ", "address": "C22"}, {"row": 23, "label": "MVWS Nozzle Spray Angle Selected", "address": "C23"}, {"row": 24, "label": "Actual Flow Based on K-Factor (L/min) = K-Factor × √ Pressure × Number of Nozzles", "address": "C24"}, {"row": 25, "label": "Actual Flow Provided Based on K-Factor (m^3/h)", "address": "C25"}, {"row": 26, "label": "Deluge Valve Selection", "address": "C26"}];
const BOM = {"header": {"A": "Sr. No", "B": "Item Description", "C": "Qty", "D": "Unit", "E": "Item No", "row": 28}, "items": [{"row": 29, "A": {"address": "A29"}, "B": {"address": "B29"}, "C": {"address": "C29"}, "D": {"address": "D29"}, "E": {"address": "E29"}}, {"row": 30, "A": {"address": "A30"}, "B": {"address": "B30"}, "C": {"address": "C30"}, "D": {"address": "D30"}, "E": {"address": "E30"}}, {"row": 31, "A": {"address": "A31"}, "B": {"address": "B31"}, "C": {"address": "C31"}, "D": {"address": "D31"}, "E": {"address": "E31"}}, {"row": 32, "A": {"address": "A32"}, "B": {"address": "B32"}, "C": {"address": "C32"}, "D": {"address": "D32"}, "E": {"address": "E32"}}, {"row": 33, "A": {"address": "A33"}, "B": {"address": "B33"}, "C": {"address": "C33"}, "D": {"address": "D33"}, "E": {"address": "E33"}}, {"row": 34, "A": {"address": "A34"}, "B": {"address": "B34"}, "C": {"address": "C34"}, "D": {"address": "D34"}, "E": {"address": "E34"}}, {"row": 35, "A": {"address": "A35"}, "B": {"address": "B35"}, "C": {"address": "C35"}, "D": {"address": "D35"}, "E": {"address": "E35"}}, {"row": 36, "A": {"address": "A36"}, "B": {"address": "B36"}, "C": {"address": "C36"}, "D": {"address": "D36"}, "E": {"address": "E36"}}, {"row": 37, "A": {"address": "A37"}, "B": {"address": "B37"}, "C": {"address": "C37"}, "D": {"address": "D37"}, "E": {"address": "E37"}}, {"row": 38, "A": {"address": "A38"}, "B": {"address": "B38"}, "C": {"address": "C38"}, "D": {"address": "D38"}, "E": {"address": "E38"}}, {"row": 39, "A": {"address": "A39"}, "B": {"address": "B39"}, "C": {"address": "C39"}, "D": {"address": "D39"}, "E": {"address": "E39"}}, {"row": 40, "A": {"address": "A40"}, "B": {"address": "B40"}, "C": {"address": "C40"}, "D": {"address": "D40"}, "E": {"address": "E40"}}, {"row": 41, "A": {"address": "A41"}, "B": {"address": "B41"}, "C": {"address": "C41"}, "D": {"address": "D41"}, "E": {"address": "E41"}}, {"row": 42, "A": {"address": "A42"}, "B": {"address": "B42"}, "C": {"address": "C42"}, "D": {"address": "D42"}, "E": {"address": "E42"}}, {"row": 43, "A": {"address": "A43"}, "B": {"address": "B43"}, "C": {"address": "C43"}, "D": {"address": "D43"}, "E": {"address": "E43"}}, {"row": 44, "A": {"address": "A44"}, "B": {"address": "B44"}, "C": {"address": "C44"}, "D": {"address": "D44"}, "E": {"address": "E44"}}, {"row": 45, "A": {"address": "A45"}, "B": {"address": "B45"}, "C": {"address": "C45"}, "D": {"address": "D45"}, "E": {"address": "E45"}}, {"row": 46, "A": {"address": "A46"}, "B": {"address": "B46"}, "C": {"address": "C46"}, "D": {"address": "D46"}, "E": {"address": "E46"}}, {"row": 47, "A": {"address": "A47"}, "B": {"address": "B47"}, "C": {"address": "C47"}, "D": {"address": "D47"}, "E": {"address": "E47"}}]};

// Convert "A1" to zero-based row/col
function a1ToRC(a1) {
  const m = a1.match(/([A-Z]+)(\d+)/i);
  if (!m) return null;
  const letters = m[1].toUpperCase();
  const row = parseInt(m[2], 10) - 1;
  let col = 0;
  for (let i = 0; i < letters.length; i++) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }
  return { row, col: col - 1 };
}

// Build HF config and instance
const hfConfig = {
  licenseKey: 'gpl-v3', // HyperFormula is GPL - OK for evaluation/open-source contexts
};

const hf = HyperFormula.HyperFormula.buildFromArray(SHEET_DATA, hfConfig);
const SHEET_ID = 0; // first (and only) sheet

// Helpers to get/set values
function getValue(addr) {
  const rc = a1ToRC(addr);
  if (!rc) return null;
  const val = hf.getCellValue({ sheet: SHEET_ID, row: rc.row, col: rc.col });
  return val;
}

function setValue(addr, value) {
  const rc = a1ToRC(addr);
  if (!rc) return;
  // numeric if parseable
  let v = value;
  if (typeof v === 'string' && v.trim() !== '' && !isNaN(Number(v))) {
    v = Number(v);
  }
  hf.setCellContents({ sheet: SHEET_ID, row: rc.row, col: rc.col }, [[v]]);
}

// Initialize UI
function mountInputs() {
  const inputsForm = document.getElementById('inputs-form');
  const advForm = document.getElementById('advanced-form');

  const B_HEADER_ROW = 6;
  const itemsA = [];
  const itemsB = [];
  INPUTS.forEach(function(it) {
    if (B_HEADER_ROW && it.row > B_HEADER_ROW) itemsB.push(it); else itemsA.push(it);
  });

  function renderSet(formEl, arr) {
    formEl.innerHTML = '';
    arr.forEach(function(it) {
      const div = document.createElement('div');
      div.className = 'row';
      const id = `inp_${it.address}`;
      const label = document.createElement('label');
      label.setAttribute('for', id);
      label.textContent = it.label;

      const inp = document.createElement('input');
      inp.id = id;
      inp.type = it.type === 'number' ? 'number' : 'text';
      const curr = getValue(it.address);
      inp.value = (curr !== null && curr !== undefined) ? curr : (it.initial ?? '');
      inp.addEventListener('input', function() {
        setValue(it.address, inp.value);
        redraw();
      });
      div.appendChild(label);
      div.appendChild(inp);
      formEl.appendChild(div);
    });
  }

  renderSet(inputsForm, itemsA);
  renderSet(advForm, itemsB);
}

function mountResults() {
  const tbody = document.querySelector('#results-table tbody');
  tbody.innerHTML = '';
  RESULTS.forEach(function(r) {
    const tr = document.createElement('tr');
    const tdLabel = document.createElement('td');
    tdLabel.textContent = r.label;
    const tdVal = document.createElement('td');
    const span = document.createElement('span');
    span.className = 'kv';
    const v = getValue(r.address);
    span.textContent = (v === null || v === undefined) ? '' : v;
    tdVal.appendChild(span);
    tr.appendChild(tdLabel);
    tr.appendChild(tdVal);
    tbody.appendChild(tr);
  });
}

function mountBOM() {
  const headerRow = document.getElementById('bom-header-row');
  headerRow.innerHTML = '';
  ['A','B','C','D','E'].forEach(function(k) {
    const th = document.createElement('th');
    th.textContent = BOM.header[k] ?? k;
    headerRow.appendChild(th);
  });

  const tbody = document.querySelector('#bom-table tbody');
  tbody.innerHTML = '';
  BOM.items.forEach(function(item) {
    const tr = document.createElement('tr');
    ['A','B','C','D','E'].forEach(function(k) {
      const td = document.createElement('td');
      const addr = item[k].address;
      const v = getValue(addr);
      td.textContent = (v === null || v === undefined) ? '' : v;
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
}

// Redraw dynamic sections (results + BOM) after any change
function redraw() {
  mountResults();
  mountBOM();
}

// Set page title from A1, if present
(function initTitle() {
  const titleEl = document.getElementById('title');
  const a1 = SHEET_DATA[0]?.[0];
  if (typeof a1 === 'string' && !a1.startsWith('=')) {
    titleEl.textContent = a1.trim();
  }
})();

// Boot
mountInputs();
redraw();
