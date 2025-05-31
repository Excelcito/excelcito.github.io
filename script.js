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
  const chars = document.getElementById('caracteres').value;
  const cabecerasInput = document.getElementById('cabeceras') ? document.getElementById('cabeceras').value : '';

  // Si no hay cabeceras, comportamiento original
  if (!cabecerasInput) {
    if (!datosExcel) {
      alert('Primero subí un archivo Excel.');
      return;
    }
    const limpio = limpiarCaracteres(datosExcel, chars);
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(limpio);
    XLSX.utils.book_append_sheet(wb, ws, 'Limpio');
    XLSX.writeFile(wb, 'limpio.xlsx');
    return;
  }

  // Si hay cabeceras, unificar todas las hojas
  const fileInput = document.getElementById('fileInput');
  const file = fileInput.files[0];
  if (!file) {
    alert('Primero subí un archivo Excel.');
    return;
  }
  const cabecerasOriginales = cabecerasInput.split(',').map(h => h.trim()); // Para mostrar/exportar
  const cabecerasDeseadas = cabecerasOriginales.map(normalizar);

  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    let filasUnificadas = [cabecerasOriginales];

    workbook.SheetNames.forEach(sheetName => {
      const sheet = workbook.Sheets[sheetName];
      const datos = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      if (!datos.length) return;

      const headers = datos[0].map(normalizar);

      // Detectar máximo sufijo numérico
      let maxGrupo = 1;
      headers.forEach(h => {
        cabecerasDeseadas.forEach(cab => {
          const match = h.match(new RegExp(`^${cab}(\\s*(\\d+))?$`));
          if (match && match[2]) {
            maxGrupo = Math.max(maxGrupo, parseInt(match[2]));
          }
        });
      });

      // Para cada fila de datos (excepto cabecera)
      for (let i = 1; i < datos.length; i++) {
        const fila = datos[i];
        // Para cada grupo posible
        for (let grupo = 0; grupo <= maxGrupo; grupo++) {
          let sufijo = (grupo === 0) ? '' : ' ' + grupo;
          // Buscar los índices de las columnas de este grupo
          let indices = cabecerasDeseadas.map(cab => {

            let idx = headers.findIndex(h => h === normalizar(cab + sufijo));
            return idx;
          });
          // Solo agregar si al menos una columna de este grupo tiene dato
          const nuevaFila = indices.map(idx => (idx !== -1 && fila[idx] !== undefined) ? fila[idx] : '');
          if (nuevaFila.some(v => v !== '' && v !== undefined)) {
            filasUnificadas.push(nuevaFila);
          }
        }
      }
    });

    // Limpiar caracteres en el resultado unificado
    const limpio = limpiarCaracteres(filasUnificadas, chars);

    mostrarTabla(limpio);

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(limpio);
    XLSX.utils.book_append_sheet(wb, ws, 'Unificado');
    XLSX.writeFile(wb, 'unificado.xlsx');
  };

  reader.readAsArrayBuffer(file);
}

// Normaliza texto para comparar (minúsculas y sin espacios extra)
function normalizar(texto) {
  return (texto || '').toString().trim().replace(/\s+/g, ' ').toLowerCase();
}
