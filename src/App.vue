<script setup lang="ts">
import { ref } from 'vue'
import * as XLSX from 'xlsx'
import { saveAs } from 'file-saver'
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  AlignmentType,
  Table,
  TableRow,
  TableCell,
  VerticalAlign,
  WidthType,
} from 'docx'

import IconDownload from './assets/icon/IconDownload.vue'
import IconUploadFiles from './assets/icon/IconUploadFiles.vue'
import IconExcel from './assets/icon/IconExcel.vue'
import IconDelete from './assets/icon/IconDelete.vue'
import IconCheck from './assets/icon/IconCheck.vue'

const fileName = ref<string>('')
const fileSize = ref<string>('')
const uploadedFile = ref<File | null>(null)
const excelData = ref<any[]>([])
const progress = ref<number>(0) // Para la barra de progreso
const isProcessing = ref<boolean>(false) // Para mostrar el estado de procesamiento
const downloadReady = ref<boolean>(false) // Para habilitar el botón de descarga
// Inicializar las variables para almacenar los datos
let coordinatesData: { punto: number; norte: number; este: number }[] = []
let colindanciasData: {
  tramo: string
  distancia: number
  colindante: string
}[] = []

const UploadFileExport = (e: Event): void => {
  const input = e.target as HTMLInputElement
  if (input.files && input.files.length > 0) {
    uploadedFile.value = input.files[0]
    fileName.value = uploadedFile.value.name
    fileSize.value = (uploadedFile.value.size / 1024).toFixed(1) + ' KB'

    const reader = new FileReader()
    reader.onload = () => {
      const fileData = reader.result as ArrayBuffer
      const wb = XLSX.read(fileData, { type: 'array' })

      // Recorrer todas las hojas del archivo Excel
      wb.SheetNames.forEach(sheetName => {
        // Obtener la hoja correspondiente a cada nombre
        const sheet = wb.Sheets[sheetName]

        // Leer los datos de la hoja en formato JSON
        const sheetData = XLSX.utils.sheet_to_json(sheet, { header: 1 })

        sheetData.forEach((row: any, index: number) => {
          if (
            row.some(
              (cell: any) =>
                typeof cell === 'string' &&
                cell.includes('CUADRO DE COORDENADAS PLANAS ORIGEN NACIONAL'),
            )
          ) {
            // Si encontramos la fila que contiene las coordenadas, procesamos las filas siguientes
            for (let i = index + 1; i < sheetData.length; i++) {
              const dataRow = sheetData[i]

              // Validamos si la fila tiene al menos 3 columnas y si la primera columna tiene datos
              if (dataRow) {
                const punto = dataRow[1] // Punto (columna 1)
                const norte = parseFloat(dataRow[2]) // Norte (columna 2)
                const este = parseFloat(dataRow[3]) // Este (columna 3)

                // Verificamos si los valores de Norte y Este son números válidos antes de agregar
                if (!isNaN(norte) && !isNaN(este)) {
                  coordinatesData.push({ punto, norte, este })
                }
              }
            }
          }

          // Buscar si alguna celda de la fila contiene 'CUADRO DE COLINDANCIAS'
          if (
            row.some(
              (cell: any) =>
                typeof cell === 'string' &&
                cell.includes('CUADRO DE COLINDANCIAS'),
            )
          ) {
            // Si encontramos la fila que contiene las colindancias, procesamos las filas siguientes
            for (let i = index + 1; i < sheetData.length; i++) {
              const dataRow = sheetData[i]

              // Validamos si la fila tiene al menos 3 columnas y si la primera columna tiene datos
              if (dataRow) {
                const tramo = dataRow[0] // Tramo (columna 1)
                const distancia = parseFloat(dataRow[1]) // Distancia (columna 2)
                const colindante = dataRow[2] // Colindante (columna 3)

                if (!isNaN(distancia)) {
                  colindanciasData.push({ tramo, distancia, colindante })
                }
              }
            }
          }
        })
      })
    }
    reader.readAsArrayBuffer(uploadedFile.value)
  }
}

const convertFile = async (): Promise<void> => {
  isProcessing.value = true
  progress.value = 0

  // Simulación de procesamiento y actualización de la barra de progreso
  const interval = setInterval(() => {
    if (progress.value < 100) {
      progress.value += 5
    } else {
      clearInterval(interval)
      downloadReady.value = true
      isProcessing.value = false
    }
  }, 200)

  // Aquí agregas la lógica para convertir los datos a formato Word (por ejemplo, usando `docxtemplater`, `html-docx-js`, etc.)
  // await createWordDocument()

  // Cuando termine el procesamiento, la barra de progreso debe estar al 100%
}

const createWordDocument = () => {
  const doc = new Document({
    sections: [
      {
        properties: {},
        children: [
          // Título del documento
          new Paragraph({
            children: [
              new TextRun({
                text: 'Descripción Técnica del Bien Inmueble',
                bold: true,
                size: 24,
              }),
            ],
            alignment: AlignmentType.CENTER,
          }),

          // Información del bien inmueble
          new Paragraph({
            children: [
              new TextRun({
                text: 'El bien inmueble identificado con nombre XXXXXXX y catastralmente con NUPRE / Número predial XXXXXXX...',
              }),
            ],
            alignment: AlignmentType.JUSTIFY,
          }),

          // Título de los linderos
          new Paragraph({
            children: [
              new TextRun({
                text: 'Linderos Técnicos:',
                bold: true,
                size: 16,
              }),
            ],
            alignment: AlignmentType.LEFT,
          }),

          // Linderos como tabla
          new Table({
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph('POR EL NORTE:')],
                    verticalAlign: VerticalAlign.CENTER,
                    width: { size: 50, type: WidthType.PERCENTAGE },
                  }),
                  new TableCell({
                    children: [
                      new Paragraph('Lindero 1: Inicia en el punto XX...'),
                      new Paragraph('Lindero 2: Inicia en el punto XX...'),
                    ],
                    verticalAlign: VerticalAlign.CENTER,
                    width: { size: 50, type: WidthType.PERCENTAGE },
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph('POR EL ESTE:')],
                    verticalAlign: VerticalAlign.CENTER,
                    width: { size: 50, type: WidthType.PERCENTAGE },
                  }),
                  new TableCell({
                    children: [
                      new Paragraph('Lindero 3: Inicia en el punto XX...'),
                      new Paragraph('Lindero 4: Inicia en el punto XX...'),
                    ],
                    verticalAlign: VerticalAlign.CENTER,
                    width: { size: 50, type: WidthType.PERCENTAGE },
                  }),
                ],
              }),
              // Añadir más filas para el Sur y Oeste siguiendo el mismo formato
            ],
          }),

          // Agregar las coordenadas de manera organizada
          new Paragraph({
            children: [
              new TextRun('Coordenadas:\n'),
              new TextRun(JSON.stringify(coordinatesData)),
            ],
            alignment: AlignmentType.LEFT,
          }),

          // Agregar las colindancias
          new Paragraph({
            children: [
              new TextRun('Colindancias:\n'),
              new TextRun(JSON.stringify(colindanciasData)),
            ],
            alignment: AlignmentType.LEFT,
          }),
        ],
      },
    ],
  })

  // Convert document to blob
  Packer.toBlob(doc).then(blob => {
    saveAs(blob, 'coordenadas_colindancias.docx')
  })
}

const resetFileUpload = (): void => {
  uploadedFile.value = null
  fileName.value = ''
  fileSize.value = ''
  const input = document.getElementById('uploadFile1') as HTMLInputElement
  if (input) {
    input.value = ''
  }
}
</script>

<template>
  <div class="flex justify-center items-center h-screen bg-blue-200">
    <div
      :class="[
        'card p-2 m-4',
        'rounded-lg bg-white bg-opacity-50 shadow-[0_0_30px_rgba(0,0,0,0.15)]',
        'min-h-[400px] max-h-[500px] min-w-[370px] max-w-[500px]',
        'relative transform-style-3d overflow-hidden',
        'flex flex-col justify-between',
      ]"
    >
      <div class="flex-none m-2 p-0 text-center">
        <span class="text-center font-semibold text-[18px]"
          >Cuadro de áreas</span
        >
      </div>
      <div class="flex-1 flex flex-col gap-[0.75rem] m-2 text-center">
        <!-- Barra de progreso -->
        <div v-if="isProcessing" class="w-full bg-gray-200 h-2 mb-2">
          <div
            class="bg-green-400 h-2"
            :style="{ width: progress + '%' }"
          ></div>
        </div>

        <label
          for="uploadFile1"
          :class="[
            'w-full h-[253px]',
            'cursor-pointer rounded font-[sans-serif]',
            'flex flex-col items-center justify-center',
            'text-gray-300 text-base border-2 border-gray-300 border-dashed',
          ]"
        >
          <IconCheck width="50px" height="50px" v-if="uploadedFile" />
          <IconUploadFiles v-else />

          <input
            type="file"
            id="uploadFile1"
            class="hidden"
            @change="UploadFileExport"
            accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          />
          <p class="text-xs font-medium text-gray-400 mt-2">
            {{
              uploadedFile
                ? 'Documento cargado con éxito'
                : 'Solo puedes cargar documentos Excel'
            }}
          </p>
        </label>
        <div
          v-if="uploadedFile"
          class="file-block h-[3rem] px-5"
          style="display: flex"
        >
          <div class="file-info">
            <IconExcel width="40px" height="100%" />
            <p class="w-full flex items-center gap-[0.40rem] px-2">
              <span class="">{{ fileName }}</span> |
              <span class="">{{ fileSize }}</span>
            </p>
          </div>
          <IconDelete
            width="30px"
            height="30px"
            fill="white"
            class="hover:scale-[1.1] cursor-pointer"
            @click="resetFileUpload"
          />
        </div>
      </div>

      <div class="flex justify-center items-center gap-[0.75rem]">
        <button
          @click="convertFile"
          :disabled="!uploadedFile"
          :class="[
            'text-white bg-green-400',
            'rounded-md inline-flex items-center',
            'font-bold py-2 px-4',
            { 'opacity-50 cursor-not-allowed': !uploadedFile },
          ]"
        >
          <span>Convertir</span>
        </button>
        <!-- Botón de descarga -->
        <div v-if="downloadReady" class="text-center">
          <button
            @click="createWordDocument"
            class="text-white bg-blue-500 py-2 px-4 rounded-md"
          >
            Descargar
          </button>
        </div>
      </div>
    </div>
  </div>
</template>

<style scoped>
.file-block {
  color: #f7fff7;
  background-color: #7b2cbf;
  transition: all 1s;
  width: 85%;
  position: relative;
  flex-direction: row;
  justify-content: space-between;
  align-items: center;
  border-radius: 25px;
  margin: 0 auto;
}

.file-info {
  display: flex;
  align-items: center;
  font-size: 14px;
}
</style>
