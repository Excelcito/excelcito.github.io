let datosExcel = null;

document.getElementById('fileInput').addEventListener('change', function (e) {
  const file = e.target.files[0];
  const reader = new FileReader();

  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    rellenarCeldasCombinadas(sheet);
    
    const datos = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false });
    datosExcel = datos;
    
    mostrarTabla(datosExcel);
  };

  reader.readAsArrayBuffer(file);
});

function mostrarTabla(data) {
  const tabla = document.getElementById('tabla');
  const LIMITE_FILAS = 10000;
  const truncado = Array.isArray(data) && data.length > LIMITE_FILAS;
  const filasVisibles = Array.isArray(data) ? data.slice(0, LIMITE_FILAS) : [];

  tabla.innerHTML = `<h2>Contenido del Excel</h2>${truncado ? '<p>Mostrando solo las primeras 10.000 filas para evitar bloqueos.</p>' : ''}`;
  const table = document.createElement('table');

  filasVisibles.forEach(row => {
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

  // Si no hay cabeceras, comportamiento original sobre la primera hoja visible
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
      rellenarCeldasCombinadas(sheet);
      
      const datos = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false });
      if (!datos.length) return;

      // Buscar la fila de cabeceras real (la que maximiza coincidencias)
      const headerRowIndex = encontrarFilaCabecera(datos, cabecerasDeseadas);
      const headers = datos[headerRowIndex].map(normalizar);

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

      // Para cada fila de datos (desde la fila siguiente a cabecera)
      for (let i = headerRowIndex + 1; i < datos.length; i++) {
        const fila = datos[i];
        // Para cada grupo posible (grupo 1 = sin número, grupo 2 = " 2", ...)
        for (let grupo = 1; grupo <= maxGrupo; grupo++) {
          let sufijo = (grupo === 1) ? '' : ' ' + grupo;
          // Buscar los índices de las columnas de este grupo (solo coincidencias exactas con sufijo)
          let indices = cabecerasDeseadas.map(cab => {
            return headers.findIndex(h => h === normalizar(cab + sufijo));
          });

          // Solo agregar si al menos una columna de este grupo tiene dato en esta fila
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

function dividirEnHojas() {
  if (!datosExcel || !datosExcel.length) {
    alert('Primero subí un archivo Excel.');
    return;
  }

  const filasInput = document.getElementById('filasPorHoja');
  const filasPorHoja = filasInput ? parseInt(filasInput.value, 10) : NaN;

  if (!filasPorHoja || filasPorHoja <= 0) {
    alert('Ingresá una cantidad válida de filas por hoja.');
    return;
  }

  const wb = XLSX.utils.book_new();
  let inicio = 0;
  let indiceHoja = 1;

  while (inicio < datosExcel.length) {
    const chunk = datosExcel.slice(inicio, inicio + filasPorHoja);
    if (!chunk.length) break;
    const nombreHoja = `Hoja ${indiceHoja}`;
    const ws = XLSX.utils.aoa_to_sheet(chunk);
    XLSX.utils.book_append_sheet(wb, ws, nombreHoja);
    inicio += filasPorHoja;
    indiceHoja++;
  }

  XLSX.writeFile(wb, 'dividido.xlsx');
}

// Normaliza texto para comparar (minúsculas y sin espacios extra)
function normalizar(texto) {
  return (texto || '').toString().trim().replace(/\s+/g, ' ').toLowerCase();
}

// Encuentra la fila de cabecera que maximiza coincidencias con las cabecerasDeseadas
function encontrarFilaCabecera(datos, cabecerasDeseadas) {
  let mejorFila = 0;
  let maxMatches = 0;
  const maxSearch = Math.min(20, datos.length); // buscar hasta 20 filas por hoja
  for (let r = 0; r < maxSearch; r++) {
    const row = datos[r] || [];
    const normalizedRow = row.map(normalizar);
    let matches = 0;
    cabecerasDeseadas.forEach(cab => {
      if (normalizedRow.includes(cab)) matches++;
    });
    if (matches > maxMatches) {
      maxMatches = matches;
      mejorFila = r;
    }
  }
  return maxMatches > 0 ? mejorFila : 0;
}

// Rellena en la hoja las celdas vacías que forman parte de rangos combinados
function rellenarCeldasCombinadas(sheet) {
  const merges = sheet['!merges'];
  if (!merges || !merges.length) return;
  let range = sheet['!ref']
    ? XLSX.utils.decode_range(sheet['!ref'])
    : {
        s: { r: merges[0].s.r, c: merges[0].s.c },
        e: { r: merges[0].e.r, c: merges[0].e.c }
      };

  merges.forEach(m => {
    const startAddr = XLSX.utils.encode_cell({ r: m.s.r, c: m.s.c });
    const startCell = sheet[startAddr];
    if (!startCell) return;
    for (let R = m.s.r; R <= m.e.r; ++R) {
      for (let C = m.s.c; C <= m.e.c; ++C) {
        const addr = XLSX.utils.encode_cell({ r: R, c: C });
        if (addr === startAddr) continue;
        if (!sheet[addr] || sheet[addr].v === undefined) {
          // duplica completamente la celda para que sheet_to_json vea el contenido
          sheet[addr] = { ...startCell };
        }
      }
    }

    if (m.s.r < range.s.r) range.s.r = m.s.r;
    if (m.s.c < range.s.c) range.s.c = m.s.c;
    if (m.e.r > range.e.r) range.e.r = m.e.r;
    if (m.e.c > range.e.c) range.e.c = m.e.c;
  });

  sheet['!ref'] = XLSX.utils.encode_range(range);
}
