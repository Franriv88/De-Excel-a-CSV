let csvContent = ""; // Variable para almacenar el contenido CSV

function handleDragOver(event) {
  event.preventDefault();
  event.dataTransfer.dropEffect = "copy";
}

function handleDrop(event) {
  event.preventDefault();
  const file = event.dataTransfer.files[0];
  if (file && file.name.endsWith(".xlsx")) {
    leerExcel(file);
  } else {
    alert("Por favor, sube un archivo Excel válido (.xlsx).");
  }
}

function leerExcel(file) {
  const reader = new FileReader();
  reader.onload = function(event) {
    const data = new Uint8Array(event.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    // Procesar filas y convertirlas en formato CSV
    csvContent = "email;quantity\n"; // Encabezado en la primera fila

    // Agregar cada fila del archivo Excel al contenido CSV, sin omitir ninguna fila
    for (let i = 0; i < rows.length; i++) { // Inicia desde la primera fila del archivo Excel
      const row = rows[i];
      if (row[0] && row[1]) {
        // Concatenar columna A y B con ';' y añadir al CSV
        csvContent += `${row[0]};${row[1]}\n`;
      }
    }

    document.getElementById("download-btn").disabled = false; // Habilitar el botón de descarga
  };

  reader.readAsArrayBuffer(file);
}

function descargarCSV() {
  if (!csvContent) {
    alert("No hay contenido para descargar.");
    return;
  }

  const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = "archivo_convertido.csv";
  link.click();
}
