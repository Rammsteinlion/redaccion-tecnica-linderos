<script setup lang="ts">
import { ref } from 'vue';
import * as XLSX from 'xlsx';
import { Document, Packer, Paragraph, TextRun } from 'docx';

type SheetData = { [key: string]: (string | number)[][] };

const sheetsData = ref<SheetData>({});

const excelExport = (e: Event): void => {
  const input = e.target as HTMLInputElement;

  if (!input.files?.length) return;

  const file = input.files[0];
  const reader = new FileReader();

  reader.onload = (event) => {
    const binaryStr = event.target?.result;
    const workbook = XLSX.read(binaryStr, { type: 'binary' });

    const allSheetsData: SheetData = {};
    workbook.SheetNames.forEach((sheetName) => {
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      allSheetsData[sheetName] = jsonData;
    });

    sheetsData.value = allSheetsData;
  };

  reader.readAsBinaryString(file);
};

const generateDocument = () => {
  const doc = new Document();

  // Suponiendo que la primera hoja se llama "Hoja1" y la segunda "Hoja2"
  const firstTableData = sheetsData.value["Hoja1"];
  const secondTableData = sheetsData.value["Hoja2"];

  if (firstTableData && secondTableData) {
    // Suponiendo que el primer elemento de cada fila en Hoja1 es un ID
    // y que en Hoja2 se relaciona por el mismo ID
    const secondTableMap = new Map<string, string>(); // Mapa para buscar rápidamente por ID

    // Crear un mapa de la segunda tabla
    for (let i = 1; i < secondTableData.length; i++) {
      const id = secondTableData[i][0]; // ID en la primera columna
      const point = secondTableData[i][1]; // Punto relacionado
      secondTableMap.set(id, point);
    }

    // Generar el documento combinando los datos
    for (let i = 1; i < firstTableData.length; i++) {
      const id = firstTableData[i][0]; // ID en la primera columna de Hoja1
      const name = firstTableData[i][1]; // Nombre en la segunda columna de Hoja1
      const relatedPoint = secondTableMap.get(id); // Buscar el punto relacionado

      const paragraph = new Paragraph({
        children: [
          new TextRun(`ID: ${id}. Nombre: ${name}. Punto relacionado: ${relatedPoint ? relatedPoint : 'No encontrado'}.`),
        ],
      });
      doc.addParagraph(paragraph);
    }

    Packer.toBuffer(doc).then((buffer) => {
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
      const link = document.createElement('a');
      link.href = window.URL.createObjectURL(blob);
      link.download = 'documento.docx';
      link.click();
    });
  } else {
    console.log('Datos insuficientes para generar el documento.');
  }
};
</script>

<template>
  <input 
    type="file" 
    name="excelFile" 
    id="excelFile" 
    @change="excelExport" 
    accept=".xlsx"
  />
  <button @click="generateDocument">Generar Documento</button>

  <div v-if="Object.keys(sheetsData).length">
    <h2>Datos del archivo Excel:</h2>
    <div v-for="(data, sheetName) in sheetsData" :key="sheetName">
      <h3>{{ sheetName }}</h3>
      <table>
        <thead>
          <tr>
            <th v-for="(header, index) in data[0]" :key="index">{{ header }}</th>
          </tr>
        </thead>
        <tbody>
          <tr v-for="(row, rowIndex) in data.slice(1)" :key="rowIndex">
            <td v-for="(cell, cellIndex) in row" :key="cellIndex">{{ cell }}</td>
          </tr>
        </tbody>
      </table>
    </div>
  </div>
</template>

<style scoped>
table {
  width: 100%;
  border-collapse: collapse;
}

th, td {
  border: 1px solid #ddd;
  padding: 8px;
}

th {
  background-color: #f2f2f2;
}
</style>
