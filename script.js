const currency = v => (isNaN(v) ? 0 : v).toFixed(2);

function rowTpl(type) {
    if (type === 'service') {
        return `<tr>
      <td class="center index"></td>
      <td><input placeholder="Description" /></td>
      <td><input class="unit" placeholder="Unit" /></td>
      <td><input class="qty" type="number" min="0" value="1" oninput="recalc()" /></td>
      <td><input class="price" type="number" min="0" value="0" oninput="recalc()" /></td>
      <td class="right amt">0.00</td>
      <td class="no-print center"><button class="danger" onclick="delRow(this)">Delete</button></td>
    </tr>`;
    } else {
        return `<tr>
      <td class="center index"></td>
      <td><input placeholder="Medicine Name" /></td>
      <td><input class="unit" placeholder="Unit" /></td>
      <td class="no-print"><input class="totalPrice" type="number" min="0" value="0" oninput="recalc()" /></td>
      <td class="no-print"><input class="totalUnits" type="number" min="1" value="1" oninput="recalc()" /></td>
      <td><input class="qty" type="number" min="1" value="1" oninput="recalc()" /></td>
      <td class="right priceUnit">0.00</td>
      <td class="right amt">0.00</td>
      <td class="no-print center"><button class="danger" onclick="delRow(this)">Delete</button></td>
    </tr>`;
    }
}

function renumber(tbody) {
    [...tbody.querySelectorAll("tr")].forEach((tr, idx) => { tr.querySelector(".index").textContent = idx + 1; });
}

function addServiceRow(desc) {
    const tb = document.getElementById('servicesBody');
    const row = document.createElement('tr'); row.innerHTML = rowTpl('service'); tb.appendChild(row);
    if (desc) row.querySelector('td:nth-child(2) input').value = desc;
    renumber(tb); recalc();
}

function addMedicineRow() {
    const tb = document.getElementById('medsBody');
    const row = document.createElement('tr'); row.innerHTML = rowTpl('med'); tb.appendChild(row);
    renumber(tb); recalc();
}

function delRow(btn) { const tr = btn.closest('tr'); const tb = tr.parentElement; tr.remove(); renumber(tb); recalc(); }

function recalc() {
    // services
    let srv = 0;
    document.querySelectorAll('#servicesBody tr').forEach(tr => {
        const qty = parseFloat(tr.querySelector('.qty')?.value || 0);
        const price = parseFloat(tr.querySelector('.price')?.value || 0);
        const amt = qty * price;
        tr.querySelector('.amt').textContent = currency(amt); srv += amt;
    });
    document.getElementById('srvSubtotal').textContent = currency(srv);

    // medicines
    let med = 0;
    document.querySelectorAll('#medsBody tr').forEach(tr => {
        const totalPrice = parseFloat(tr.querySelector('.totalPrice')?.value || 0);
        const totalUnits = parseFloat(tr.querySelector('.totalUnits')?.value || 1);
        const unitPrice = totalUnits > 0 ? (totalPrice / totalUnits) : 0;
        tr.querySelector('.priceUnit').textContent = currency(unitPrice);
        const qty = parseFloat(tr.querySelector('.qty')?.value || 1);
        const amt = qty * unitPrice;
        tr.querySelector('.amt').textContent = currency(amt); med += amt;
    });
    document.getElementById('medSubtotal').textContent = currency(med);

    document.getElementById('grandTotal').textContent = currency(srv + med);
}

function resetForm() {
    ['patientName', 'age', 'gender', 'mobile', 'email', 'address'].forEach(id => document.getElementById(id).value = '');
    document.getElementById('servicesBody').innerHTML = '';
    document.getElementById('medsBody').innerHTML = '';
    addServiceRow('Registration Charges'); addServiceRow('Room Rent'); addServiceRow('Consultant Charges'); addServiceRow('OT Charges');
    addMedicineRow(); recalc();
}

function uniqueId() {
    const d = new Date(), pad = n => n.toString().padStart(2, '0');
    return `PAT-${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}-${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}-${Math.random().toString(36).substr(2, 4).toUpperCase()}`;
}

function gatherData() {
    const services = [...document.querySelectorAll('#servicesBody tr')].map(tr => ({
        description: tr.querySelector('td:nth-child(2) input').value.trim(),
        unit: tr.querySelector('.unit').value.trim(),
        quantity: parseFloat(tr.querySelector('.qty')?.value || 0),
        price: parseFloat(tr.querySelector('.price')?.value || 0),
        amount: parseFloat(tr.querySelector('.amt').textContent || 0)
    }));
    const medicines = [...document.querySelectorAll('#medsBody tr')].map(tr => ({
        name: tr.querySelector('td:nth-child(2) input').value.trim(),
        unit: tr.querySelector('.unit').value.trim(),
        totalPrice: parseFloat(tr.querySelector('.totalPrice')?.value || 0),
        totalUnits: parseFloat(tr.querySelector('.totalUnits')?.value || 0),
        pricePerUnit: parseFloat(tr.querySelector('.priceUnit').textContent || 0),
        quantity: parseFloat(tr.querySelector('.qty')?.value || 0),
        amount: parseFloat(tr.querySelector('.amt').textContent || 0)
    }));
    return {
        id: uniqueId(),
        patient: {
            name: document.getElementById('patientName').value.trim(),
            age: document.getElementById('age').value.trim(),
            gender: document.getElementById('gender').value,
            mobile: document.getElementById('mobile').value.trim(),
            email: document.getElementById('email').value.trim(),
            address: document.getElementById('address').value.trim()
        },
        services, medicines,
        totals: {
            section1: parseFloat(document.getElementById('srvSubtotal').textContent),
            section2: parseFloat(document.getElementById('medSubtotal').textContent),
            grand: parseFloat(document.getElementById('grandTotal').textContent)
        },
        createdAt: new Date().toISOString()
    };
}

// File System Access API
// let dirHandle = null;
// async function getDirHandle() {
//     if (!dirHandle) { dirHandle = await window.showDirectoryPicker(); }
//     return dirHandle;
// }
// async function saveAsJSON() {
//     const data = gatherData();
//     const filename = `${data.id}.json`;
//     const dir = await getDirHandle();
//     const fileHandle = await dir.getFileHandle(filename, { create: true });
//     const writable = await fileHandle.createWritable();
//     await writable.write(JSON.stringify(data, null, 2));
//     await writable.close();
//     alert("Saved successfully: " + filename);
// }

// ...existing code...

// ...existing code...

let excelDirHandle = null;

async function getExcelDirHandle() {
    if (!excelDirHandle) {
        excelDirHandle = await window.showDirectoryPicker();
    }
    return excelDirHandle;
}

async function saveAsExcel() {
    const data = gatherData();

    // Patient Info as a single row
    const patientSheet = [
        [
            "Name", "Age", "Gender", "Mobile", "Email", "Address", "BP", "Pulse", "Temp", "SpO2", "Height", "Weight", "RBS", "Date"
        ],
        [
            data.patient.name, data.patient.age, data.patient.gender, data.patient.mobile, data.patient.email, data.patient.address,
            data.patient.bp, data.patient.pulse, data.patient.temp, data.patient.spo2, data.patient.height, data.patient.weight, data.patient.rbs, data.patient.date
        ]
    ];

    // Services
    const servicesSheet = [
        ["Description", "Unit", "Quantity", "Price", "Amount"],
        ...data.services.map(s => [s.description, s.unit, s.quantity, s.price, s.amount])
    ];

    // Medicines
    const medicinesSheet = [
        ["Name", "Unit", "Total Price", "Total Units", "Price Per Unit", "Quantity", "Amount"],
        ...data.medicines.map(m => [m.name, m.unit, m.totalPrice, m.totalUnits, m.pricePerUnit, m.quantity, m.amount])
    ];

    // Totals
    const totalsSheet = [
        ["Section 1 Total", data.totals.section1],
        ["Section 2 Total", data.totals.section2],
        ["Grand Total", data.totals.grand]
    ];

    // Create workbook and add sheets
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(patientSheet), "Patient");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(servicesSheet), "Services");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(medicinesSheet), "Medicines");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(totalsSheet), "Totals");

    // Generate Excel file as Blob
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbout], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });

    // Prepare filename: PatientName-details.xlsx (sanitize filename)
    let patientName = data.patient.name ? data.patient.name.trim().replace(/[^a-z0-9_\-]/gi, "_") : "Patient";
    let filename = `${patientName}-details.xlsx`;

    // Save to selected folder
    try {
        const dirHandle = await getExcelDirHandle();
        const fileHandle = await dirHandle.getFileHandle(filename, { create: true });
        const writable = await fileHandle.createWritable();
        await writable.write(blob);
        await writable.close();
        alert("Excel file saved: " + filename);
    } catch (e) {
        alert("Failed to save Excel file: " + e.message);
    }
}
// ...existing code...
// ...existing code...

resetForm();
document.addEventListener('input', () => recalc());

// ...existing code...

function checkSectionVisibility() {
    // Services section
    const servicesBody = document.getElementById('servicesBody');
    const serviceRows = Array.from(servicesBody.children);
    const hasServiceData = serviceRows.length > 0 && serviceRows.some(row =>
        Array.from(row.querySelectorAll('input, select')).some(input => input.value.trim() !== '')
    );
    document.querySelectorAll('.section-title')[0].classList.toggle('hide-print', !hasServiceData);
    document.getElementById('servicesTable').classList.toggle('hide-print', !hasServiceData);
    document.getElementById('srvSubtotal').closest('.totals').classList.toggle('hide-print', !hasServiceData);

    // Medicines section
    const medsBody = document.getElementById('medsBody');
    const medsRows = Array.from(medsBody.children);
    const hasMedsData = medsRows.length > 0 && medsRows.some(row =>
        Array.from(row.querySelectorAll('input, select')).some(input => input.value.trim() !== '')
    );
    document.querySelectorAll('.section-title')[1].classList.toggle('hide-print', !hasMedsData);
    document.getElementById('medsTable').classList.toggle('hide-print', !hasMedsData);
    document.getElementById('medSubtotal').closest('.totals').classList.toggle('hide-print', !hasMedsData);
}

// Update add, delete, and recalc functions to call checkSectionVisibility

function addServiceRow(desc) {
    const tb = document.getElementById('servicesBody');
    const row = document.createElement('tr'); row.innerHTML = rowTpl('service'); tb.appendChild(row);
    if (desc) row.querySelector('td:nth-child(2) input').value = desc;
    renumber(tb); recalc(); checkSectionVisibility();
}

function addMedicineRow() {
    const tb = document.getElementById('medsBody');
    const row = document.createElement('tr'); row.innerHTML = rowTpl('med'); tb.appendChild(row);
    renumber(tb); recalc(); checkSectionVisibility();
}

function delRow(btn) {
    const tr = btn.closest('tr');
    const tb = tr.parentElement;
    tr.remove();
    renumber(tb);
    recalc();
    checkSectionVisibility();
}

function recalc() {
    // services
    let srv = 0;
    document.querySelectorAll('#servicesBody tr').forEach(tr => {
        const qty = parseFloat(tr.querySelector('.qty')?.value || 0);
        const price = parseFloat(tr.querySelector('.price')?.value || 0);
        const amt = qty * price;
        tr.querySelector('.amt').textContent = currency(amt); srv += amt;
    });
    document.getElementById('srvSubtotal').textContent = currency(srv);

    // medicines
    let med = 0;
    document.querySelectorAll('#medsBody tr').forEach(tr => {
        const totalPrice = parseFloat(tr.querySelector('.totalPrice')?.value || 0);
        const totalUnits = parseFloat(tr.querySelector('.totalUnits')?.value || 1);
        const unitPrice = totalUnits > 0 ? (totalPrice / totalUnits) : 0;
        tr.querySelector('.priceUnit').textContent = currency(unitPrice);
        const qty = parseFloat(tr.querySelector('.qty')?.value || 1);
        const amt = qty * unitPrice;
        tr.querySelector('.amt').textContent = currency(amt); med += amt;
    });
    document.getElementById('medSubtotal').textContent = currency(med);

    document.getElementById('grandTotal').textContent = currency(srv + med);

    checkSectionVisibility();
}

function resetForm() {
    ['patientName', 'age', 'gender', 'mobile', 'email', 'address'].forEach(id => document.getElementById(id).value = '');
    document.getElementById('servicesBody').innerHTML = '';
    document.getElementById('medsBody').innerHTML = '';
    addServiceRow('Registration Charges'); addServiceRow('Room Rent'); addServiceRow('Consultant Charges'); addServiceRow('OT Charges');
    addMedicineRow(); recalc(); checkSectionVisibility();
}

function deleteServiceSection() {
    if (confirm("Delete all services?")) {
        document.getElementById('servicesBody').innerHTML = '';
        document.getElementById('srvSubtotal').textContent = '0.00';
        document.getElementById('grandTotal').textContent = currency(
            parseFloat(document.getElementById('medSubtotal').textContent) || 0
        );
        checkSectionVisibility();
    }
}

function deleteMedicineSection() {
    if (confirm("Delete all medicines?")) {
        document.getElementById('medsBody').innerHTML = '';
        document.getElementById('medSubtotal').textContent = '0.00';
        document.getElementById('grandTotal').textContent = currency(
            parseFloat(document.getElementById('srvSubtotal').textContent) || 0
        );
        checkSectionVisibility();
    }
}

// ...existing code...

// ...existing code...

function checkSectionVisibility() {
    // Services section
    const servicesBody = document.getElementById('servicesBody');
    const serviceRows = Array.from(servicesBody.children);
    const hasServiceData = serviceRows.length > 0 && serviceRows.some(row =>
        Array.from(row.querySelectorAll('input, select')).some(input => input.value.trim() !== '')
    );
    document.querySelectorAll('.section-title')[0].classList.toggle('hide-print', !hasServiceData);
    document.getElementById('servicesTable').classList.toggle('hide-print', !hasServiceData);
    document.getElementById('srvSubtotal').closest('.totals').classList.toggle('hide-print', !hasServiceData);

    // Medicines section
    const medsBody = document.getElementById('medsBody');
    const medsRows = Array.from(medsBody.children);
    const hasMedsData = medsRows.length > 0 && medsRows.some(row =>
        Array.from(row.querySelectorAll('input, select')).some(input => input.value.trim() !== '')
    );
    document.querySelectorAll('.section-title')[1].classList.toggle('hide-print', !hasMedsData);
    document.getElementById('medsTable').classList.toggle('hide-print', !hasMedsData);
    document.getElementById('medSubtotal').closest('.totals').classList.toggle('hide-print', !hasMedsData);

    // Grand Total (Total Payable) section
    const showGrandTotal = hasServiceData || hasMedsData;
    document.getElementById('grandTotal').closest('.totals').classList.toggle('hide-print', !showGrandTotal);
}

// ...existing code...

resetForm();
document.addEventListener('input', () => recalc());

// ...existing code...
function checkSectionVisibility() {
    // Services section
    const servicesBody = document.getElementById('servicesBody');
    const serviceRows = Array.from(servicesBody.children);
    const hasServiceData = serviceRows.length > 0 && serviceRows.some(row =>
        Array.from(row.querySelectorAll('input, select')).some(input => input.value.trim() !== '')
    );
    document.getElementById('servicesTitle').classList.toggle('hide-print', !hasServiceData);
    document.getElementById('servicesTable').classList.toggle('hide-print', !hasServiceData);
    document.getElementById('srvSubtotal').closest('.totals').classList.toggle('hide-print', !hasServiceData);

    // Medicines section
    const medsBody = document.getElementById('medsBody');
    const medsRows = Array.from(medsBody.children);
    const hasMedsData = medsRows.length > 0 && medsRows.some(row =>
        Array.from(row.querySelectorAll('input, select')).some(input => input.value.trim() !== '')
    );
    document.getElementById('medicinesTitle').classList.toggle('hide-print', !hasMedsData);
    document.getElementById('medsTable').classList.toggle('hide-print', !hasMedsData);
    document.getElementById('medSubtotal').closest('.totals').classList.toggle('hide-print', !hasMedsData);

    // Grand Total (Total Payable) section
    const showGrandTotal = hasServiceData || hasMedsData;
    document.getElementById('grandTotal').closest('.totals').classList.toggle('hide-print', !showGrandTotal);
}
// ...existing code...

// ...existing code...
function gatherData() {
    const services = [...document.querySelectorAll('#servicesBody tr')].map(tr => ({
        description: tr.querySelector('td:nth-child(2) input').value.trim(),
        unit: tr.querySelector('.unit').value.trim(),
        quantity: parseFloat(tr.querySelector('.qty')?.value || 0),
        price: parseFloat(tr.querySelector('.price')?.value || 0),
        amount: parseFloat(tr.querySelector('.amt').textContent || 0)
    }));
    const medicines = [...document.querySelectorAll('#medsBody tr')].map(tr => ({
        name: tr.querySelector('td:nth-child(2) input').value.trim(),
        unit: tr.querySelector('.unit').value.trim(),
        totalPrice: parseFloat(tr.querySelector('.totalPrice')?.value || 0),
        totalUnits: parseFloat(tr.querySelector('.totalUnits')?.value || 0),
        pricePerUnit: parseFloat(tr.querySelector('.priceUnit').textContent || 0),
        quantity: parseFloat(tr.querySelector('.qty')?.value || 0),
        amount: parseFloat(tr.querySelector('.amt').textContent || 0)
    }));
    return {
        id: uniqueId(),
        patient: {
            name: document.getElementById('patientName').value.trim(),
            age: document.getElementById('age').value.trim(),
            gender: document.getElementById('gender').value,
            mobile: document.getElementById('mobile').value.trim(),
            email: document.getElementById('email').value.trim(),
            address: document.getElementById('address').value.trim(),
            bp: document.getElementById('bp').value.trim(),
            pulse: document.getElementById('pulse').value.trim(),
            temp: document.getElementById('temp').value.trim(),
            spo2: document.getElementById('spo2').value.trim(),
            height: document.getElementById('height').value.trim(),
            weight: document.getElementById('weight').value.trim(),
            rbs: document.getElementById('rbs').value.trim(),
            date: document.getElementById('date').value
        },
        services, medicines,
        totals: {
            section1: parseFloat(document.getElementById('srvSubtotal').textContent),
            section2: parseFloat(document.getElementById('medSubtotal').textContent),
            grand: parseFloat(document.getElementById('grandTotal').textContent)
        },
        createdAt: new Date().toISOString()
    };
}
// ...existing code...