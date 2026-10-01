/**
 * DocToExcel AI - Client Application Logic
 */

document.addEventListener('DOMContentLoaded', () => {
  // State
  let currentFile = null;
  let currentImageBase64 = '';
  let currentMimeType = '';
  let extractedData = {
    documento: {},
    columnas: ["no", "nombre", "cui", "fecha_nacimiento", "sexo", "telefono", "domicilio", "municipio", "diagnostico", "tratamiento"],
    filas: []
  };

  // Zoom & Pan State
  let zoomLevel = 1.0;
  let rotateDegree = 0;
  let panX = 0;
  let panY = 0;
  let isDragging = false;
  let startX = 0;
  let startY = 0;

  // DOM Elements
  const dropzone = document.getElementById('dropzone');
  const fileInput = document.getElementById('fileInput');
  const uploadSection = document.getElementById('uploadSection');
  const loadingSection = document.getElementById('loadingSection');
  const workspaceSection = document.getElementById('workspaceSection');
  const loadingText = document.getElementById('loadingText');
  const modelSelect = document.getElementById('modelSelect');

  // Image Viewer Elements
  const documentImage = document.getElementById('documentImage');
  const imageViewport = document.getElementById('imageViewport');
  const zoomPercent = document.getElementById('zoomPercent');
  const btnZoomIn = document.getElementById('btnZoomIn');
  const btnZoomOut = document.getElementById('btnZoomOut');
  const btnResetZoom = document.getElementById('btnResetZoom');
  const btnRotate = document.getElementById('btnRotate');
  const fileNameBadge = document.getElementById('fileNameBadge');

  // Table & Metadata Elements
  const tableHead = document.getElementById('tableHead');
  const tableBody = document.getElementById('tableBody');
  const rowCountBadge = document.getElementById('rowCountBadge');
  const metaTitulo = document.getElementById('metaTitulo');
  const metaMedico = document.getElementById('metaMedico');
  const metaLugar = document.getElementById('metaLugar');
  const metaFecha = document.getElementById('metaFecha');

  // Actions
  const btnNewDoc = document.getElementById('btnNewDoc');
  const btnAddRow = document.getElementById('btnAddRow');
  const btnCopyTable = document.getElementById('btnCopyTable');
  const btnExportExcel = document.getElementById('btnExportExcel');

  // Modal API Key Elements
  const apiKeyModal = document.getElementById('apiKeyModal');
  const btnOpenApiKey = document.getElementById('btnOpenApiKey');
  const btnCloseModal = document.getElementById('btnCloseModal');
  const btnSaveApiKey = document.getElementById('btnSaveApiKey');
  const inputApiKey = document.getElementById('inputApiKey');
  const btnToggleApiKeyVisibility = document.getElementById('btnToggleApiKeyVisibility');
  const apiKeyStatusText = document.getElementById('apiKeyStatusText');

  // 1. API Key Setup
  function getStoredApiKey() {
    return localStorage.getItem('GEMINI_API_KEY') || '';
  }

  function updateApiKeyUI() {
    const key = getStoredApiKey();
    if (key) {
      apiKeyStatusText.textContent = 'API Key Configurada';
      btnOpenApiKey.classList.remove('btn-outline');
      btnOpenApiKey.classList.add('btn-secondary');
    } else {
      apiKeyStatusText.textContent = 'Ingresar API Key';
    }
  }

  updateApiKeyUI();

  btnOpenApiKey.addEventListener('click', () => {
    inputApiKey.value = getStoredApiKey();
    apiKeyModal.classList.remove('hidden');
  });

  btnCloseModal.addEventListener('click', () => {
    apiKeyModal.classList.add('hidden');
  });

  btnToggleApiKeyVisibility.addEventListener('click', () => {
    const isPassword = inputApiKey.type === 'password';
    inputApiKey.type = isPassword ? 'text' : 'password';
    btnToggleApiKeyVisibility.querySelector('i').className = isPassword ? 'fa-solid fa-eye-slash' : 'fa-solid fa-eye';
  });

  btnSaveApiKey.addEventListener('click', () => {
    const val = inputApiKey.value.trim();
    if (val) {
      localStorage.setItem('GEMINI_API_KEY', val);
      showToast('API Key guardada correctamente', 'success');
    } else {
      localStorage.removeItem('GEMINI_API_KEY');
      showToast('API Key eliminada', 'info');
    }
    updateApiKeyUI();
    apiKeyModal.classList.add('hidden');
  });

  // 2. Drag & Drop File Handling
  ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
    dropzone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
    });
  });

  ['dragenter', 'dragover'].forEach(eventName => {
    dropzone.addEventListener(eventName, () => dropzone.classList.add('dragover'));
  });

  ['dragleave', 'drop'].forEach(eventName => {
    dropzone.addEventListener(eventName, () => dropzone.classList.remove('dragover'));
  });

  dropzone.addEventListener('drop', (e) => {
    const files = e.dataTransfer.files;
    if (files.length > 0) {
      handleSelectedFile(files[0]);
    }
  });

  fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
      handleSelectedFile(e.target.files[0]);
    }
  });

  function handleSelectedFile(file) {
    if (!file.type.startsWith('image/') && file.type !== 'application/pdf') {
      showToast('Por favor selecciona un archivo de imagen válido (JPG, PNG, WEBP) o PDF.', 'error');
      return;
    }

    currentFile = file;
    fileNameBadge.textContent = file.name;

    const reader = new FileReader();
    reader.onload = (e) => {
      const dataUrl = e.target.result;
      currentImageBase64 = dataUrl.split(',')[1];
      currentMimeType = file.type;

      documentImage.src = dataUrl;
      resetZoom();
      processFileWithGemini(file, currentImageBase64, currentMimeType);
    };
    reader.readAsDataURL(file);
  }

  // 3. Process File with Gemini API
  async function processFileWithGemini(file, base64Data, mimeType) {
    uploadSection.classList.add('hidden');
    workspaceSection.classList.add('hidden');
    loadingSection.classList.remove('hidden');

    const apiKey = getStoredApiKey();
    const model = modelSelect.value;

    loadingText.textContent = `Procesando manuscrito con ${model}...`;

    try {
      let resData = null;
      let isSuccess = false;

      // Intentar primero a través del servidor backend
      try {
        const formData = new FormData();
        formData.append('file', file);
        formData.append('model', model);
        if (apiKey) formData.append('api_key', apiKey);

        const headers = {};
        if (apiKey) headers['X-Gemini-API-Key'] = apiKey;

        const response = await fetch('/api/analyze', {
          method: 'POST',
          headers: headers,
          body: formData
        });

        resData = await response.json();
        if (response.ok && resData.success) {
          isSuccess = true;
        } else if (resData.error && resData.error.includes('NameResolutionError')) {
          console.warn('Backend sin salida de red sandbox. Redirigiendo a llamada directa cliente Gemini API...');
        } else if (!response.ok) {
          throw new Error(resData.error || 'Error al comunicarse con la IA');
        }
      } catch (backendErr) {
        console.warn('Error en backend proxy, usando cliente directo:', backendErr);
      }

      // Si el backend falló por red/DNS o sandbox, llamar a la API de Gemini directamente desde el navegador del cliente
      if (!isSuccess) {
        if (!apiKey) {
          throw new Error('Falta la clave API de Gemini. Haz clic en "API Key" e ingrésala para continuar.');
        }

        const systemPrompt = `Eres un experto en digitalización de documentos manuscritos y médicos.
Tu objetivo es analizar la imagen del documento cargado (como formularios de consulta externa, registros de pacientes, listas o tablas) y extraer toda la información estructurada con máxima precisión.

REGLAS DE EXTRACCIÓN:
1. Extrae los metadatos del encabezado del documento si existen (ej. Médico, Lugar de Atención, Fecha del documento).
2. Extrae cada fila de la tabla principal en un objeto con los siguientes campos estándar:
   "no", "nombre", "cui", "fecha_nacimiento", "sexo", "telefono", "domicilio", "municipio", "diagnostico", "tratamiento".
3. Si una celda está vacía o tachada, coloca "".

RESPONDE ÚNICAMENTE EN FORMATO JSON VÁLIDO CON ESTA ESTRUCTURA EXACTA:
{
  "documento": { "titulo": "...", "medico": "...", "lugar": "...", "fecha": "..." },
  "columnas": ["no", "nombre", "cui", "fecha_nacimiento", "sexo", "telefono", "domicilio", "municipio", "diagnostico", "tratamiento"],
  "filas": [ { "no": "8", "nombre": "...", "cui": "...", "fecha_nacimiento": "...", "sexo": "...", "telefono": "...", "domicilio": "...", "municipio": "...", "diagnostico": "...", "tratamiento": "..." } ]
}`;

        const geminiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
        
        const payload = {
          contents: [
            {
              parts: [
                { text: systemPrompt },
                {
                  inline_data: {
                    mime_type: mimeType.startsWith('image/') ? mimeType : 'image/jpeg',
                    data: base64Data
                  }
                }
              ]
            }
          ],
          generationConfig: {
            temperature: 0.1,
            response_mime_type: "application/json"
          }
        };

        const geminiRes = await fetch(geminiUrl, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload)
        });

        const geminiData = await geminiRes.json();
        if (!geminiRes.ok) {
          const msg = geminiData.error ? geminiData.error.message : 'Error al consultar API de Gemini';
          throw new Error(`Gemini API Error: ${msg}`);
        }

        const candidateText = geminiData.candidates[0].content.parts[0].text;
        let cleanText = candidateText.strip ? candidateText.strip() : candidateText.trim();
        if (cleanText.startsWith('```json')) cleanText = cleanText.substring(7);
        if (cleanText.endsWith('```')) cleanText = cleanText.substring(0, cleanText.length - 3);

        resData = {
          success: true,
          data: JSON.parse(cleanText.trim())
        };
      }

      extractedData = resData.data;
      renderExtractedData(extractedData);

      loadingSection.classList.add('hidden');
      workspaceSection.classList.remove('hidden');
      showToast('Documento digitalizado con éxito.', 'success');


    } catch (err) {
      console.error(err);
      loadingSection.classList.add('hidden');
      uploadSection.classList.remove('hidden');
      
      if (err.message.includes('Falta la clave API')) {
        apiKeyModal.classList.remove('hidden');
      }
      showToast(err.message || 'Ocurrió un error al procesar la imagen.', 'error');
    }
  }

  // 4. Render Extracted Data to UI
  function renderExtractedData(data) {
    const doc = data.documento || {};
    metaTitulo.value = doc.titulo || 'INFORMACIÓN DIARIA DEL PACIENTE';
    metaMedico.value = doc.medico || '';
    metaLugar.value = doc.lugar || '';
    metaFecha.value = doc.fecha || '';

    // Columns
    const colsMap = {
      "no": "No.",
      "nombre": "NOMBRE DEL PACIENTE",
      "cui": "CUI / DPI",
      "fecha_nacimiento": "FECHA DE NACIMIENTO",
      "sexo": "SEXO",
      "telefono": "TELÉFONO",
      "domicilio": "DOMICILIO",
      "municipio": "MUNICIPIO",
      "diagnostico": "DIAGNÓSTICO",
      "tratamiento": "TRATAMIENTO"
    };

    let cols = data.columnas || ["no", "nombre", "cui", "fecha_nacimiento", "sexo", "telefono", "domicilio", "municipio", "diagnostico", "tratamiento"];
    
    // Render Header
    let headHtml = '<tr>';
    cols.forEach(colKey => {
      const label = colsMap[colKey.toLowerCase()] || colKey.toUpperCase();
      headHtml += `<th>${label}</th>`;
    });
    headHtml += '<th style="width: 40px; text-align: center;"><i class="fa-solid fa-gear"></i></th>';
    headHtml += '</tr>';
    tableHead.innerHTML = headHtml;

    // Render Body
    renderTableBody(data.filas || [], cols);
  }

  function renderTableBody(rows, cols) {
    tableBody.innerHTML = '';
    rowCountBadge.textContent = `${rows.length} Filas`;

    rows.forEach((row, rowIndex) => {
      const tr = document.createElement('tr');
      
      cols.forEach(colKey => {
        const td = document.createElement('td');
        const input = document.createElement('input');
        input.type = 'text';
        input.className = 'cell-input';
        input.value = row[colKey] || '';
        
        input.addEventListener('input', (e) => {
          extractedData.filas[rowIndex][colKey] = e.target.value;
        });

        td.appendChild(input);
        tr.appendChild(td);
      });

      // Actions Column (Delete row button)
      const tdActions = document.createElement('td');
      tdActions.style.textAlign = 'center';
      const btnDel = document.createElement('button');
      btnDel.className = 'btn-delete-row';
      btnDel.title = 'Eliminar fila';
      btnDel.innerHTML = '<i class="fa-solid fa-trash-can"></i>';
      btnDel.addEventListener('click', () => {
        extractedData.filas.splice(rowIndex, 1);
        renderTableBody(extractedData.filas, cols);
      });

      tdActions.appendChild(btnDel);
      tr.appendChild(tdActions);
      tableBody.appendChild(tr);
    });
  }

  // Add New Row
  btnAddRow.addEventListener('click', () => {
    const cols = extractedData.columnas || ["no", "nombre", "cui", "fecha_nacimiento", "sexo", "telefono", "domicilio", "municipio", "diagnostico", "tratamiento"];
    const newRow = {};
    cols.forEach(c => newRow[c] = '');
    newRow.no = (extractedData.filas.length + 1).toString();
    
    if (!extractedData.filas) extractedData.filas = [];
    extractedData.filas.push(newRow);
    renderTableBody(extractedData.filas, cols);
  });

  // Reset to Upload Section
  btnNewDoc.addEventListener('click', () => {
    workspaceSection.classList.add('hidden');
    uploadSection.classList.remove('hidden');
    fileInput.value = '';
  });

  // 5. Image Pan & Zoom Math
  function applyImageTransform() {
    documentImage.style.transform = `translate(${panX}px, ${panY}px) scale(${zoomLevel}) rotate(${rotateDegree}deg)`;
    zoomPercent.textContent = `${Math.round(zoomLevel * 100)}%`;
  }

  function resetZoom() {
    zoomLevel = 1.0;
    rotateDegree = 0;
    panX = 0;
    panY = 0;
    applyImageTransform();
  }

  btnZoomIn.addEventListener('click', () => {
    zoomLevel = Math.min(zoomLevel + 0.25, 4.0);
    applyImageTransform();
  });

  btnZoomOut.addEventListener('click', () => {
    zoomLevel = Math.max(zoomLevel - 0.25, 0.5);
    applyImageTransform();
  });

  btnResetZoom.addEventListener('click', resetZoom);

  btnRotate.addEventListener('click', () => {
    rotateDegree = (rotateDegree + 90) % 360;
    applyImageTransform();
  });

  imageViewport.addEventListener('wheel', (e) => {
    e.preventDefault();
    if (e.deltaY < 0) {
      zoomLevel = Math.min(zoomLevel + 0.1, 4.0);
    } else {
      zoomLevel = Math.max(zoomLevel - 0.1, 0.5);
    }
    applyImageTransform();
  });

  imageViewport.addEventListener('mousedown', (e) => {
    isDragging = true;
    startX = e.clientX - panX;
    startY = e.clientY - panY;
  });

  window.addEventListener('mousemove', (e) => {
    if (!isDragging) return;
    panX = e.clientX - startX;
    panY = e.clientY - startY;
    applyImageTransform();
  });

  window.addEventListener('mouseup', () => {
    isDragging = false;
  });

  // 6. Export to Excel (.xlsx) using SheetJS or Backend Endpoint
  btnExportExcel.addEventListener('click', async () => {
    try {
      // Sync metadata from inputs
      extractedData.documento = {
        titulo: metaTitulo.value.trim(),
        medico: metaMedico.value.trim(),
        lugar: metaLugar.value.trim(),
        fecha: metaFecha.value.trim()
      };

      // Attempt server side export
      const response = await fetch('/api/export-excel', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(extractedData)
      });

      if (response.ok) {
        const blob = await response.blob();
        const downloadUrl = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = downloadUrl;
        a.download = `digitalizacion_${new Date().toISOString().slice(0, 10)}.xlsx`;
        document.body.appendChild(a);
        a.click();
        a.remove();
        showToast('Excel descargado con éxito.', 'success');
        return;
      }

    } catch (e) {
      console.log('Fallback a exportación en cliente mediante SheetJS:', e);
    }

    // Fallback: Client side SheetJS export
    if (typeof XLSX !== 'undefined') {
      const colsMap = {
        "no": "No.",
        "nombre": "NOMBRE DEL PACIENTE",
        "cui": "CUI / DPI",
        "fecha_nacimiento": "FECHA DE NACIMIENTO",
        "sexo": "SEXO",
        "telefono": "TELÉFONO",
        "domicilio": "DOMICILIO",
        "municipio": "MUNICIPIO",
        "diagnostico": "DIAGNÓSTICO",
        "tratamiento": "TRATAMIENTO"
      };

      const sheetData = [];
      sheetData.push([metaTitulo.value || "INFORMACIÓN DIARIA DEL PACIENTE"]);
      sheetData.push([]);
      if (metaMedico.value) sheetData.push(["Médico:", metaMedico.value]);
      if (metaLugar.value) sheetData.push(["Lugar de Atención:", metaLugar.value]);
      if (metaFecha.value) sheetData.push(["Fecha:", metaFecha.value]);
      sheetData.push([]);

      const cols = extractedData.columnas || ["no", "nombre", "cui", "fecha_nacimiento", "sexo", "telefono", "domicilio", "municipio", "diagnostico", "tratamiento"];
      const headerRow = cols.map(c => colsMap[c.toLowerCase()] || c.toUpperCase());
      sheetData.push(headerRow);

      extractedData.filas.forEach(row => {
        sheetData.push(cols.map(c => row[c] || ''));
      });

      const ws = XLSX.utils.aoa_to_sheet(sheetData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Pacientes");

      XLSX.writeFile(wb, `pacientes_${new Date().toISOString().slice(0,10)}.xlsx`);
      showToast('Excel generado con éxito.', 'success');
    }
  });

  // Copy Table to Clipboard
  btnCopyTable.addEventListener('click', () => {
    const cols = extractedData.columnas || ["no", "nombre", "cui", "fecha_nacimiento", "sexo", "telefono", "domicilio", "municipio", "diagnostico", "tratamiento"];
    let text = cols.join('\t') + '\n';
    
    extractedData.filas.forEach(r => {
      text += cols.map(c => r[c] || '').join('\t') + '\n';
    });

    navigator.clipboard.writeText(text).then(() => {
      showToast('Tabla copiada al portapapeles en formato Excel.', 'success');
    });
  });

  // Toast System
  function showToast(message, type = 'info') {
    const container = document.getElementById('toastContainer');
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    
    const iconClass = type === 'success' ? 'fa-solid fa-circle-check' : (type === 'error' ? 'fa-solid fa-triangle-exclamation' : 'fa-solid fa-circle-info');
    toast.innerHTML = `<i class="${iconClass}"></i> <span>${message}</span>`;
    
    container.appendChild(toast);
    
    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateY(10px)';
      setTimeout(() => toast.remove(), 300);
    }, 4000);
  }

});
