const gradeMap = {
  "A+": 4.0, "A": 4.0, "A-": 3.7,
  "B+": 3.3, "B": 3.0, "B-": 2.7,
  "C+": 2.3, "C": 2.0, "C-": 1.7,
  "D": 1.0, "E": 0.0, "F": 0.0
};

const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const exportButtons = document.getElementById('exportButtons');

// Drag & Drop
dropZone.addEventListener('click', () => fileInput.click());
fileInput.addEventListener('change', handleFile);
dropZone.addEventListener('dragover', e => {
  e.preventDefault();
  dropZone.classList.add('dragover');
});
dropZone.addEventListener('dragleave', () => {
  dropZone.classList.remove('dragover');
});
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('dragover');
  if (e.dataTransfer.files.length) {
    handleFile({ target: { files: e.dataTransfer.files } });
  }
});

function handleFile(e) {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
    calculateGPAByYear(rows);
  };
  reader.readAsArrayBuffer(file);
}

function calculateGPAByYear(rows) {
  const grouped = {};
  let totalCredits = 0, totalPoints = 0;

  rows.forEach(row => {
    const courseKey = Object.keys(row).find(k => row[k]?.toUpperCase?.().includes("BECS"));
    const gradeKey = Object.keys(row).find(k => gradeMap[row[k]?.toUpperCase?.().trim()] !== undefined);
    if (!courseKey || !gradeKey) return;

    const code = row[courseKey].toUpperCase().trim();
    const grade = row[gradeKey].toUpperCase().trim();
    const point = gradeMap[grade];
    const credit = parseInt(code.slice(-1));
    const year = code[5];

    if (isNaN(credit) || point === undefined || isNaN(year)) return;

    const yearKey = `Year ${year}`;
    if (!grouped[yearKey]) grouped[yearKey] = { courses: [], credits: 0, points: 0 };

    grouped[yearKey].courses.push({ code, grade, credit, points: point * credit });
    grouped[yearKey].credits += credit;
    grouped[yearKey].points += point * credit;
    totalCredits += credit;
    totalPoints += point * credit;
  });

  let html = "";
  if (Object.keys(grouped).length === 0) {
    document.getElementById('output').innerHTML = "<p>No valid data found.</p>";
    exportButtons.classList.add("hidden");
    return;
  }

  for (const year of Object.keys(grouped).sort()) {
    const y = grouped[year];
    const gpa = (y.points / y.credits).toFixed(3);
    html += `<h3>${year}</h3>`;
    html += `<table><thead><tr><th>Course Code</th><th>Grade</th><th>Credits</th><th>Points</th></tr></thead><tbody>`;
    y.courses.forEach(course => {
      html += `<tr><td>${course.code}</td><td>${course.grade}</td><td>${course.credit}</td><td>${course.points.toFixed(2)}</td></tr>`;
    });
    html += `</tbody></table>`;
    html += `<p class="summary">Total Credits: ${y.credits} | GPA: ${gpa}</p>`;
  }

  const finalGPA = (totalPoints / totalCredits).toFixed(3);
  html += `<div class="final-summary"><h3>Overall GPA Summary</h3><p>Total Credits: ${totalCredits}</p><p>Final GPA: ${finalGPA}</p></div>`;

  document.getElementById('output').innerHTML = html;
  exportButtons.classList.remove("hidden");
}

function exportToExcel() {
  const tables = document.querySelectorAll("table");
  if (!tables.length) return alert("No GPA tables to export!");

  const wb = XLSX.utils.book_new();

  tables.forEach((table, index) => {
    const ws = XLSX.utils.table_to_sheet(table);
    const year = table.previousElementSibling?.textContent || `Sheet${index + 1}`;
    XLSX.utils.book_append_sheet(wb, ws, year.replace(/\s+/g, "_"));
  });

  XLSX.writeFile(wb, "GPA_Report.xlsx");
}


function printPDF() {
  const content = document.getElementById("output").innerHTML;
  const win = window.open("", "", "height=700,width=900");
  win.document.write("<html><head><title>GPA Report</title>");
  win.document.write("<style>table{width:100%;border-collapse:collapse}th,td{border:1px solid #000;padding:8px;text-align:center}th{background:#007bff;color:#fff}</style>");
  win.document.write("</head><body>");
  win.document.write(content);
  win.document.write("</body></html>");
  win.document.close();
  win.print();
}

function toggleDarkMode() {
  document.body.classList.toggle("dark");
}
