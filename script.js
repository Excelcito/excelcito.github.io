let datosExcel = null;

document.getElementById('fileInput').addEventListener('change', function (e) {
  const file = e.target.files[0];
  const reader = new FileReader();

  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    datosExcel = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    mostrarTabla(datosExcel);
  };

  reader.readAsArrayBuffer(file);
});

function mostrarTabla(data) {
  const tabla = document.getElementById('tabla');
  tabla.innerHTML = '<h2>Contenido del Excel</h2>';
  const table = document.createElement('table');

  data.forEach(row => {
    const tr = document.createElement('tr');
    row.forEach(cell => {
      const td = document.createElement('td');
      td.textContent = cell ?? '';
      tr.appendChild(td);
    });
    table.appendChild(tr);
  });

  tabla.appendChild(table);
}

function limpiarCaracteres(data, chars) {
  if (!chars) return data;
  const regex = new RegExp(`[${chars.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&')}]`, 'g');
  return data.map(row =>
    row.map(cell =>
      typeof cell === 'string' ? cell.replace(regex, '') : cell
    )
  );
}

function exportarSinCaracteres() {
  if (!datosExcel) {
    alert('Primero subí un archivo Excel.');
    return;
  }

  const chars = document.getElementById('caracteres').value;
  const limpio = limpiarCaracteres(datosExcel, chars);

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(limpio);
  XLSX.utils.book_append_sheet(wb, ws, 'Limpio');

  XLSX.writeFile(wb, 'limpio.xlsx');
}
