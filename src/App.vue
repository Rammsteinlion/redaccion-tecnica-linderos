<script setup lang="ts">
import { ref } from 'vue'
import * as XLSX from 'xlsx'
import IconDownload from './assets/icon/IconDownload.vue'
import IconUploadFiles from './assets/icon/IconUploadFiles.vue'
import IconExcel from './assets/icon/IconExcel.vue'
import IconDelete from './assets/icon/IconDelete.vue'
import IconCheck from './assets/icon/IconCheck.vue'

const fileName = ref<string>('')
const fileSize = ref<string>('')
const uploadedFile = ref<File | null>(null)
const coordinatesMatches = ref<
  { punto: number; norte: number; este: number }[]
>([])
const colindanciasMatches = ref<
  { tramo: string; distancia: number; colindante: string }[]
>([])

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

      wb.SheetNames.forEach(sheetName => {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], {
          header: 1,
        }) as any[]

        // Variables para detectar las secciones
        let isCoordinatesSection = false
        let isColindanciasSection = false

        coordinatesMatches.value = []
        colindanciasMatches.value = []

        rows.forEach(row => {
          // Verificamos si encontramos el encabezado de coordenadas
          if (row[0] === 'CUADRO DE COORDENADAS PLANAS ORIGEN NACIONAL') {
            isCoordinatesSection = true
            isColindanciasSection = false
            return
          }

          // Verificamos si encontramos el encabezado de colindancias
          if (row[0] === 'CUADRO DE COLINDANCIAS') {
            isCoordinatesSection = false
            isColindanciasSection = true
            return
          }

          // Procesamos las coordenadas
          if (isCoordinatesSection && row[0] !== 'PUNTO' && row.length >= 3) {
            const punto = parseInt(row[0])
            const norte = parseFloat(row[1])
            const este = parseFloat(row[2])

            if (!isNaN(punto) && !isNaN(norte) && !isNaN(este)) {
              coordinatesMatches.value.push({ punto, norte, este })
            }
          }

          // Procesamos las colindancias
          if (isColindanciasSection && row[0] !== 'TRAMO' && row.length >= 3) {
            const tramo = row[0]
            const distancia = parseFloat(row[1])
            const colindante = row[2]

            if (tramo && !isNaN(distancia) && colindante) {
              colindanciasMatches.value.push({ tramo, distancia, colindante })
            }
          }
        })

        console.log('Coordenadas:', coordinatesMatches.value)
        console.log('Colindancias:', colindanciasMatches.value)
      })
    }
    reader.readAsArrayBuffer(uploadedFile.value)
  }
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
      class="card p-2 m-4 rounded-lg bg-white bg-opacity-50 shadow-[0_0_30px_rgba(0,0,0,0.15)] min-h-[400px] max-h-[500px] min-w-[370px] max-w-[500px] relative transform-style-3d overflow-hidden flex flex-col justify-between"
    >
      <div class="flex-none m-2 p-0 text-center">
        <span class="text-center font-semibold text-[18px]"
          >Cuadro de áreas</span
        >
      </div>
      <div class="flex-1 flex flex-col gap-[0.75rem] m-2 text-center">
        <label
          for="uploadFile1"
          class="w-full h-[253px] cursor-pointer rounded font-[sans-serif] flex flex-col items-center justify-center text-gray-300 text-base border-2 border-gray-300 border-dashed"
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
              <span>{{ fileName }}</span> | <span>{{ fileSize }}</span>
            </p>
          </div>
          <IconDelete
            width="30px"
            height="30px"
            fill="white"
            class="hover:scale-[1.1] cursor-pointer"
            @click="resetFileUpload"
          />
          <div class="progress-bar" style="width: 0px"></div>
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
