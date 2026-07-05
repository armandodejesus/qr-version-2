// --- Constantes de Conversión ---
// 96 DPI es una resolución común. 1 pulgada = 25.4 mm. 1 mm = 96/25.4 píxeles.
const PIXELS_PER_MM = 96 / 25.4; // Factor de conversión de mm a px
const MM_PER_CM = 10; // 1 cm = 10 mm
const HTML2CANVAS_SCALE = 2; // Escala para html2canvas (mayor = mejor calidad, mayor tamaño de imagen)

// --- Variables Globales ---
let excelData = [];
let selectedColumns = {};
let generatedLabels = [];
let currentLabelFormat = {}; // Objeto para almacenar las configuraciones de formato aplicadas

// --- Elementos del DOM ---
const fileInput = document.getElementById('excelFile');
const formatDesignDiv = document.getElementById('format-design');
const dataSelectionDiv = document.getElementById('data-selection');
const columnSelectorsDiv = document.getElementById('column-selectors');
const applyFormatBtn = document.getElementById('applyFormatBtn');
const labelPreviewDiv = document.getElementById('label-preview');
const labelsContainer = document.getElementById('labels-container');
const downloadAllJpgBtn = document.getElementById('downloadAllJpgBtn');
const downloadAllPdfBtn = document.getElementById('downloadAllPdfBtn');

// --- Elementos de Diseño ---
const labelWidthCmInput = document.getElementById('labelWidthCm');
const labelHeightCmInput = document.getElementById('labelHeightCm');
const includeQRCheckbox = document.getElementById('includeQR');
const qrPositionSelect = document.getElementById('qrPosition');
const qrSizeCmInput = document.getElementById('qrSizeCm');
const qrColorInput = document.getElementById('qrColor');
const qrBackgroundColorInput = document.getElementById('qrBackgroundColor');

// --- Event Listeners ---
fileInput.addEventListener('change', handleFileSelect);
applyFormatBtn.addEventListener('click', applyAndGenerateLabels);
downloadAllJpgBtn.addEventListener('click', downloadAllAsJpg);
downloadAllPdfBtn.addEventListener('click', downloadAllAsPdf);

// --- Verificación Inicial de la Librería QR ---
console.log("Verificando QRCode globalmente:", typeof QRCode); // Debería ser 'function' o 'object' si se cargó localmente


// --- Funciones de Utilidad ---
function cmToPx(cm) {
    // Convierte centímetros a píxeles usando las constantes definidas
    return cm * MM_PER_CM * PIXELS_PER_MM;
}

function pxToCm(px) {
    // Convierte píxeles a centímetros
    return (px / PIXELS_PER_MM) / MM_PER_CM;
}

// --- Funciones Principales ---

function handleFileSelect(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = e.target.result;
        const workbook = XLSX.read(data, { type: 'binary' });
        const firstSheetName = workbook.SheetNames[0]; // Asume datos en la primera hoja
        const worksheet = workbook.Sheets[firstSheetName];
        excelData = XLSX.utils.sheet_to_json(worksheet);

        console.log("Datos del Excel cargados:", excelData);

        if (excelData.length > 0) {
            displayColumnSelectors(); // Muestra los selectores de columnas
            formatDesignDiv.classList.remove('hidden'); // Muestra la sección de diseño
            dataSelectionDiv.classList.remove('hidden'); // Muestra la sección de selección de datos
            labelPreviewDiv.classList.add('hidden');     // Oculta la vista previa anterior
        } else {
            alert('El archivo Excel está vacío o no contiene datos.');
        }
    };
    reader.readAsBinaryString(file);
}

function displayColumnSelectors() {
    columnSelectorsDiv.innerHTML = ''; // Limpiar selectores anteriores
    if (excelData.length === 0) return;

    const headers = Object.keys(excelData[0]); // Obtiene los nombres de las columnas

    // Selector para el dato del código QR
    const qrDataColumnDiv = document.createElement('div');
    qrDataColumnDiv.innerHTML = `
        <label for="qrDataColumn">Dato para Código QR (Columna Excel):</label>
        <select id="qrDataColumn">
            <option value="">-- Selecciona una columna --</option>
            ${headers.map(header => `<option value="${header}">${header}</option>`).join('')}
        </select>
    `;
    columnSelectorsDiv.appendChild(qrDataColumnDiv);

    // Selectores para el contenido de la etiqueta
    const labelContentDiv = document.createElement('div');
    labelContentDiv.innerHTML = `
        <h4>Contenido de la Etiqueta (Columnas Excel):</h4>
        <div id="label-fields">
            <div class="label-field">
                <label>Dato 1:</label>
                <select class="labelColumn">
                    <option value="">-- Selecciona una columna --</option>
                    ${headers.map(header => `<option value="${header}">${header}</option>`).join('')}
                </select>
            </div>
        </div>
        <button type="button" onclick="addLabelField()">Añadir otro campo a la etiqueta</button>
    `;
    columnSelectorsDiv.appendChild(labelContentDiv);
}

function addLabelField() {
    const labelFieldsDiv = document.getElementById('label-fields');
    const newFieldDiv = document.createElement('div');
    newFieldDiv.classList.add('label-field');
    /*  const headers = Object.keys(excelData[0]);   */
	const headers = excelData && excelData.length > 0 ? Object.keys(excelData[0]) : [];
    newFieldDiv.innerHTML = 
		`
        <label>Dato:</label>
        <select class="labelColumn">
            <option value="">-- Selecciona una columna --</option>
            ${headers.map(header => `<option value="${header}">${header}</option>`).join('')}
        </select>
        <button type="button" onclick="this.parentNode.remove()">Eliminar</button>
    `;
    labelFieldsDiv.appendChild(newFieldDiv);
}

function applyAndGenerateLabels() {
    // 1. Capturar selecciones de columnas
    const qrDataColumnSelect = document.getElementById('qrDataColumn');
    const labelColumnInputs = document.querySelectorAll('.labelColumn');

    selectedColumns = {
        qrDataColumn: qrDataColumnSelect ? qrDataColumnSelect.value : null,
        labelDataColumns: Array.from(labelColumnInputs).map(input => input.value).filter(col => col !== '')
    };

    console.log("Columnas seleccionadas:", selectedColumns);

    if (!selectedColumns.qrDataColumn || selectedColumns.labelDataColumns.length === 0) {
        alert("Por favor, selecciona al menos una columna para el dato del código QR y una para el texto de la etiqueta.");
        return;
    }

    // 2. Capturar configuraciones de formato y convertir a Píxeles
    currentLabelFormat = {
        widthPx: cmToPx(parseFloat(labelWidthCmInput.value) || 5),
        heightPx: cmToPx(parseFloat(labelHeightCmInput.value) || 3),
        includeQR: includeQRCheckbox.checked,
        qrPosition: qrPositionSelect.value,
        qrSizePx: cmToPx(parseFloat(qrSizeCmInput.value) || 2),
        qrColor: qrColorInput.value,
        qrBackgroundColor: qrBackgroundColorInput.value,
    };

    console.log("Formato de etiqueta seleccionado (en PX):", {
        width: currentLabelFormat.widthPx,
        height: currentLabelFormat.heightPx,
        qrSize: currentLabelFormat.qrSizePx
    });

    // 3. Validaciones: Asegurar que el QR no se solape
    const qrPadding = 5; // Margen para el QR dentro de la etiqueta (en píxeles)
    if (currentLabelFormat.includeQR) {
        if (currentLabelFormat.qrSizePx > currentLabelFormat.widthPx - (2 * qrPadding)) {
            alert(`El tamaño del código QR (${pxToCm(currentLabelFormat.qrSizePx).toFixed(1)} cm) es demasiado grande para el ancho de la etiqueta (${pxToCm(currentLabelFormat.widthPx).toFixed(1)} cm).`);
            return;
        }
        if (currentLabelFormat.qrSizePx > currentLabelFormat.heightPx - (2 * qrPadding)) {
            alert(`El tamaño del código QR (${pxToCm(currentLabelFormat.qrSizePx).toFixed(1)} cm) es demasiado grande para el alto de la etiqueta (${pxToCm(currentLabelFormat.heightPx).toFixed(1)} cm).`);
            return;
        }
    }

    // 4. Generar las etiquetas
    generateLabels(currentLabelFormat);
}


function generateLabels(format) {
    if (!format) {
        console.error("No se proporcionó formato para generar etiquetas.");
        return;
    }

    labelsContainer.innerHTML = ''; // Limpiar vista previa
    generatedLabels = []; // Limpiar etiquetas previas

    excelData.forEach((row, index) => {
        const qrData = row[selectedColumns.qrDataColumn] || '';
        const labelContentParts = selectedColumns.labelDataColumns.map(col => row[col] || '');
        const fullLabelContent = labelContentParts.join('\n');

        // --- Crear elemento canvas para QR (si se incluye) ---
        let qrCanvas = null;
        if (format.includeQR) {
            qrCanvas = document.createElement('canvas');
            qrCanvas.id = `qr-canvas-${index}`;
            qrCanvas.width = format.qrSizePx;
            qrCanvas.height = format.qrSizePx;
            console.log(`QR requerido para fila ${index + 1}. Dato: "${qrData}"`);
        } else {
            console.log(`QR omitido por formato para fila ${index + 1}.`);
        }

        // --- Verificación de QRCode ---
        if (format.includeQR && typeof QRCode === 'undefined') {
            console.error("¡ERROR FATAL! QRCode no está definido. Librería no cargada.");
            alert("Error: La librería de generación de QR no está disponible. Revisa la consola del navegador.");
            return; // Detener si QR es requerido y la librería no existe
        }

        // --- Función para dibujar el QR y luego crear la etiqueta ---
        const drawQRAndCreateLabel = (callback) => {
            if (!format.includeQR || !qrCanvas) {
                // Si no se incluye QR, creamos la etiqueta directamente
                createLabel(null, null, callback); // Pasa null para QR y error simulado
                return;
            }

            QRCode.toCanvas(qrCanvas, qrData, {
                width: format.qrSizePx,
                height: format.qrSizePx,
                colorDark: format.qrColor,
                colorLight: format.qrBackgroundColor,
                errorCorrectionLevel: 'H' // Nivel de corrección alto
            }, function (error) {
                if (error) {
                    console.error(`Error al generar QR para fila ${index + 1}:`, error);
                } else {
                    console.log(`QR generado para fila ${index + 1}.`);
                }
                createLabel(qrCanvas, error, callback); // Pasa el canvas y el error
            });
        };

        // --- Función para crear la etiqueta HTML ---
        const createLabel = (qrCanvasElement, qrError, finalCallback) => {
            const labelElement = document.createElement('div');
            labelElement.classList.add('label');
            labelElement.id = `label-${index}`;
            labelElement.style.width = `${format.widthPx}px`;
            labelElement.style.height = `${format.heightPx}px`;
            labelElement.style.border = '1px solid #ccc'; // Borde básico
            labelElement.style.padding = '10px'; // Padding interno para el contenido
            labelElement.style.margin = '10px';
            labelElement.style.display = 'flex';
            labelElement.style.position = 'relative';
            labelElement.style.boxSizing = 'border-box';
            labelElement.style.overflow = 'hidden'; // Crucial para que el QR no se salga
			
		

		
            // Estilos de texto base
            const textElementStyles = {
                textAlign: 'center', // Alineación horizontal del texto
                width: '100%',
                whiteSpace: 'pre-wrap', // Mantiene saltos de línea (\n)
                flexGrow: '1', // Permite que el texto ocupe espacio disponible
                padding: '5px', // Pequeño padding para el texto
                overflowY: 'auto' // Si el texto es muy largo, que se pueda hacer scroll dentro de la etiqueta
            };

            // Posicionamiento y alineación del QR, y ajuste de la alineación del texto
            let textFlexDirection = 'column'; // Texto y QR en columna por defecto
            let textAlignment = 'center';     // Alineación horizontal para todos los elementos internos
            let textJustification = 'center'; // Justificación vertical del contenido principal

            if (format.includeQR && qrCanvasElement) {
                const qrContainer = document.createElement('div');
                qrContainer.style.position = 'absolute';
                qrContainer.style.width = `${format.qrSizePx}px`;
                qrContainer.style.height = `${format.qrSizePx}px`;
                // Añadir margen al QR para que no toque bordes de la etiqueta
                const qrMargin = 5; // Margen en píxeles
                qrContainer.style.margin = `${qrMargin}px`;

                switch (format.qrPosition) {
                    case 'top-left':
                        qrContainer.style.top = `${qrMargin}px`; qrContainer.style.left = `${qrMargin}px`;
                        textFlexDirection = 'column'; textAlignment = 'left'; textJustification = 'flex-start'; // Texto arriba a la izquierda
                        break;
                    case 'top-right':
                        qrContainer.style.top = `${qrMargin}px`; qrContainer.style.right = `${qrMargin}px`;
                        textFlexDirection = 'column'; textAlignment = 'right'; textJustification = 'flex-start'; // Texto arriba a la derecha
                        break;
                    case 'bottom-left':
                        qrContainer.style.bottom = `${qrMargin}px`; qrContainer.style.left = `${qrMargin}px`;
                        textFlexDirection = 'column'; textAlignment = 'left'; textJustification = 'flex-end'; // Texto abajo a la izquierda
                        break;
                    case 'center':
                        qrContainer.style.top = '50%'; qrContainer.style.left = '50%';
                        qrContainer.style.transform = 'translate(-50%, -50%)';
                        textFlexDirection = 'column'; textAlignment = 'center'; textJustification = 'center'; // Todo centrado
                        break;
					// Dentro de createLabel, en el switch de format.qrPosition
					case 'top-center':
    					qrContainer.style.top = `${qrMargin}px`; // Usa el margen definido
    					qrContainer.style.left = '50%';
    					qrContainer.style.transform = 'translateX(-50%)'; // Centra horizontalmente
    					textFlexDirection = 'column'; // Organiza el contenido en columna
    					textAlignment = 'center';     // Centra el texto horizontalmente
    					textJustification = 'flex-start'; // Justifica el texto en la parte superior
    					break;
					case 'bottom-center':
						qrContainer.style.bottom = `${qrMargin}px`; // Usa el margen definido
    					qrContainer.style.left = '50%';
    					qrContainer.style.transform = 'translateX(-50%)'; // Centra horizontalmente
    					textFlexDirection = 'column'; // Organiza el contenido en columna
    					textAlignment = 'center';     // Centra el texto horizontalmente
    					textJustification = 'flex-end'; // Justifica el texto en la parte inferior
						break;
// ... se mantendrían los otros casos existentes [T8](3) [T9](4)

                    case 'bottom-right': // Por defecto
                    default:
                        qrContainer.style.bottom = `${qrMargin}px`; qrContainer.style.right = `${qrMargin}px`;
                        textFlexDirection = 'column'; textAlignment = 'right'; textJustification = 'flex-end'; // Texto abajo a la derecha
                        break;
                }
                qrContainer.appendChild(qrCanvasElement);
                labelElement.appendChild(qrContainer);
            } else if (!format.includeQR) {
                // Si no hay QR, el texto debe centrarse verticalmente en la etiqueta
                textJustification = 'center';
            }

            // Configurar estilos del contenedor principal de la etiqueta
            labelElement.style.flexDirection = textFlexDirection;
            labelElement.style.justifyContent = textJustification; // Alineación vertical del contenido
            labelElement.style.alignItems = textAlignment;       // Alineación horizontal de los elementos internos

            // Crear y configurar el div del texto
            const textDiv = document.createElement('div');
            Object.assign(textDiv.style, textElementStyles);
            textDiv.innerText = fullLabelContent;
            labelElement.prepend(textDiv); // Añadir texto (antes del QR si está abajo)

            labelsContainer.appendChild(labelElement);
            generatedLabels.push(labelElement);
            console.log(`Etiqueta ${index + 1} creada. QR: ${format.includeQR && qrCanvasElement ? 'Sí' : 'No'}.`);

            if (finalCallback) finalCallback(); // Llama al callback final si existe
        };

        // Ejecutar la lógica: dibujar QR (si aplica) y luego crear la etiqueta
        drawQRAndCreateLabel(function() {
            // Este callback se ejecuta después de que la etiqueta se creó y añadió al DOM.
        });

    }); // Fin del forEach

    labelPreviewDiv.classList.remove('hidden');
}


// --- Funciones de Descarga (ajustadas para usar el formato aplicado) ---

function downloadAllAsJpg() {
    if (generatedLabels.length === 0) {
        alert('No hay etiquetas generadas para descargar.');
        return;
    }
    alert('Iniciando descarga de JPGs...');
    generatedLabels.forEach((labelElement, index) => {
        // Usamos la escala global definida
        html2canvas(labelElement, { scale: HTML2CANVAS_SCALE }).then(canvas => {
            const imgData = canvas.toDataURL('image/jpeg');
            const link = document.createElement('a');
            link.href = imgData;
            link.download = `etiqueta_${index + 1}.jpg`;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }).catch(err => {
            console.error(`Error al generar JPG para etiqueta ${index + 1}:`, err);
            alert(`Error al generar JPG para etiqueta ${index + 1}. Revisa la consola.`);
        });
    });
}

function downloadAllAsPdf() {
    if (generatedLabels.length === 0) {
        alert('No hay etiquetas generadas para descargar.');
        return;
    }
    alert('Generando PDF...');
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF({
        orientation: 'portrait', // Orientación 'portrait' o 'landscape'
        unit: 'mm',             // Unidad de medida en milímetros
        format: 'a4'            // Tamaño de página (a4, letter, etc.)
    });

    // Usamos el formato actual aplicado para calcular el tamaño en mm
    const format = currentLabelFormat; // Formato en Píxeles

    // Convertir las dimensiones de píxeles (usadas en html2canvas) a mm para jsPDF
    // El tamaño en Píxeles del canvas renderizado es `format.widthPx / HTML2CANVAS_SCALE`.
    // Este valor se convierte a mm: `(format.widthPx / HTML2CANVAS_SCALE) / PIXELS_PER_MM`
    const labelWidthMm = (format.widthPx / HTML2CANVAS_SCALE) / PIXELS_PER_MM;
    const labelHeightMm = (format.heightPx / HTML2CANVAS_SCALE) / PIXELS_PER_MM;

    const paddingMM = 5; // Margen entre etiquetas y bordes de la página
    let xPos = paddingMM;
    let yPos = paddingMM;
    const pageWidthMm = pdf.internal.pageSize.getWidth();
    const pageHeightMm = pdf.internal.pageSize.getHeight();

    const addLabelToPdf = async (labelElement, index) => {
        return new Promise((resolve) => {
            // Usamos la misma escala que en JPG para consistencia
            html2canvas(labelElement, { scale: HTML2CANVAS_SCALE }).then(canvas => {
                const imgData = canvas.toDataURL('image/png'); // PNG es bueno para conversiones a PDF

                // Comprobar si la etiqueta cabe en la página actual
                if (yPos + labelHeightMm > pageHeightMm - paddingMM) {
                    pdf.addPage(); // Añadir nueva página si no cabe verticalmente
                    yPos = paddingMM;
                    xPos = paddingMM;
                }
                // Comprobar si la etiqueta cabe en la línea actual (columna)
                if (xPos + labelWidthMm > pageWidthMm - paddingMM) {
                    yPos += labelHeightMm + paddingMM; // Pasar a la siguiente línea
                    xPos = paddingMM; // Volver a la primera columna
                    // Si al pasar a la siguiente línea tampoco cabe, añadir nueva página
                    if (yPos + labelHeightMm > pageHeightMm - paddingMM) {
                        pdf.addPage();
                        yPos = paddingMM;
                    }
                }

                // Añadir la imagen de la etiqueta al PDF
                pdf.addImage(imgData, 'PNG', xPos, yPos, labelWidthMm, labelHeightMm);
                yPos += labelHeightMm + paddingMM; // Mover la posición Y para la siguiente etiqueta

                resolve();
            }).catch(err => {
                console.error(`Error al convertir etiqueta ${index + 1} a imagen para PDF:`, err);
                resolve(); // Resolvemos para no bloquear el proceso si una etiqueta falla
            });
        });
    };

    // Bucle asíncrono para añadir todas las etiquetas al PDF
    (async () => {
        for (let i = 0; i < generatedLabels.length; i++) {
            await addLabelToPdf(generatedLabels[i], i);
        }
        pdf.save('etiquetas_generadas.pdf');
        alert('PDF generado y descargado.');
    })();
}
