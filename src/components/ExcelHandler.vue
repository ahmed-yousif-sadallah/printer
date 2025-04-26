<template>
  <div class="p-5 max-w-4xl mx-auto">
    <h2 class="text-2xl font-bold mb-6">Excel File Handler</h2>

    <div class="flex items-center gap-5 mb-6">
      <input
        type="file"
        @change="handleFileUpload"
        ref="fileInput"
        class="block w-full text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-full file:border-0 file:text-sm file:font-semibold file:bg-violet-50 file:text-violet-700 hover:file:bg-violet-100"
      />

      <button
        @click="exportToExcel"
        class="px-4 py-2 bg-green-500 text-white rounded-md hover:bg-green-600 transition-colors"
      >
        Export to Excel
      </button>
    </div>


    <div class="flex flex-col gap-2">
            <div class="flex items-center gap-4">
              <input
                v-model="customLabelsInput"
                placeholder="Enter labels (separated by comma or space)"
                class="flex-1 px-4 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
              />
              <button
                @click="printCustomLabels"
                :disabled="!customLabelsInput.trim()"
                class="px-4 py-2 bg-green-500 text-white rounded-md hover:bg-green-600 transition-colors disabled:bg-gray-400 disabled:cursor-not-allowed"
              >
                Print Custom Labels
              </button>
            </div>
            <p class="text-sm text-gray-600">Example: C555C, C556C, C557C</p>
          </div>


    <div v-if="excelData.length" class="mt-6">
      <div class="mb-4">
        <p class="text-lg font-semibold">
          Remaining Labels: {{ remainingLabels }}
        </p>
      </div>

      <div class="mb-6 space-y-4">
        <div class="flex items-center gap-4">
          <button
            @click="connectToPrinter"
            class="px-4 py-2 bg-blue-500 text-white rounded-md hover:bg-blue-600 transition-colors"
          >
            Connect to Printer
          </button>
          <span class="text-sm text-gray-600">Connect your printer first</span>
        </div>

        <div class="flex items-center gap-4">
          <button
            @click="sendSetHeadCommand"
            class="px-4 py-2 bg-purple-500 text-white rounded-md hover:bg-purple-600 transition-colors"
          >
            Set Head
          </button>
          <span class="text-sm text-gray-600">Send SET HEAD command</span>
        </div>

        <div class="flex flex-col gap-4">
          <div class="flex items-center gap-4">
            <button
              @click="printTestLabels"
              :disabled="!labels.length"
              class="px-4 py-2 bg-blue-500 text-white rounded-md hover:bg-blue-600 transition-colors disabled:bg-gray-400 disabled:cursor-not-allowed"
            >
              Print All Labels
            </button>
            <div class="flex items-center gap-2">
              <span class="text-sm font-medium">Status:</span>
              <span v-if="printedLabels.length" class="text-sm text-green-600">
                Printed {{ printedLabels.length }} of {{ labels.length }} labels
              </span>
              <span v-else class="text-sm text-gray-600"> Ready to print </span>
            </div>
          </div>

         
        </div>
      </div>

      <h3 class="text-xl font-semibold mb-4">Preview</h3>
      <div class="overflow-x-auto" ref="tableRef">
        <table class="min-w-full border-collapse">
          <thead>
            <tr>
              <th class="px-4 py-2 bg-gray-100 border border-gray-300">
                Position
              </th>
              <th
                v-for="(header, index) in headers"
                :key="index"
                class="px-4 py-2 bg-gray-100 border border-gray-300"
              >
                {{ header }}
              </th>
            </tr>
          </thead>
          <tbody>
            <tr
              v-for="(row, rowIndex) in excelData"
              :key="rowIndex"
              :class="[
                rowIndex % 2 === 0 ? 'bg-white' : 'bg-gray-50',
                printedLabels.includes(`${row[0]}${row[1]}${row[2]}`)
                  ? 'opacity-50'
                  : '',
                currentPrintingRow === rowIndex
                  ? 'ring-2 ring-blue-500 bg-blue-50'
                  : '',
              ]"
            >
              <td class="px-4 py-2 border border-gray-300 font-medium">
                {{ excelData.length - rowIndex }}
              </td>
              <td
                v-for="(cell, cellIndex) in row"
                :key="cellIndex"
                class="px-4 py-2 border border-gray-300"
              >
                {{ cell }}
              </td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  </div>
</template>

<script setup>
import { ref, computed } from "vue";
import * as XLSX from "xlsx";

const fileInput = ref(null);
const excelData = ref([]);
const headers = ref([]);
const labels = ref([]);
const printedLabels = ref([]);
const currentPrintingRow = ref(null);
const tableRef = ref(null);
const customLabelsInput = ref("");

const remainingLabels = computed(() => {
  return labels.value.length - printedLabels.value.length;
});

const handleFileUpload = (event) => {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const firstSheet = workbook.Sheets[workbook.SheetNames[1]];
    const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

    if (jsonData.length > 0) {
      headers.value = jsonData[0];
      excelData.value = jsonData.slice(1);

      console.log(`total rows: ${excelData.value.length}`);
      // Create string from first three columns of each row
      excelData.value.forEach((row, index) => {
        const rowString = `${row[0]}${row[1]}${row[2]}`;
        console.log(`${rowString} ${index + 1}`);
        labels.value.push(rowString);
      });
    }
  };
  reader.readAsArrayBuffer(file);
};

const exportToExcel = () => {
  if (excelData.value.length === 0) return;

  const ws = XLSX.utils.aoa_to_sheet([headers.value, ...excelData.value]);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

  XLSX.writeFile(wb, "exported_data.xlsx");
};

const connectToPrinter = async () => {
  try {
    const device = await navigator.usb.requestDevice({
      filters: [{}], // Allow all USB devices for selection
    });

    console.log(`Vendor ID: ${device.vendorId}`);
    console.log(`Product ID: ${device.productId}`);
    console.log(device);
  } catch (err) {
    console.error("Device selection failed:", err);
  }
};

const sendSetHeadCommand = async () => {
  try {
    const device = await navigator.usb.requestDevice({
      filters: [{ vendorId: 0x20d1 }],
    });

    await device.open();
    if (device.configuration === null) {
      await device.selectConfiguration(1);
    }
    await device.claimInterface(0);

    const iface = device.configuration.interfaces[0].alternate;
    const outEndpoint = iface.endpoints.find((e) => e.direction === "out");

    const cmds = [
      "SET HEAD 1",
      "PRINT 1"
    ];
    const data = new TextEncoder().encode(cmds.join("\r\n") + "\r\n");
    await device.transferOut(outEndpoint.endpointNumber, data);
    console.log("✅ SET HEAD command sent successfully");
    await device.close();
  } catch (err) {
    console.error("SET HEAD command error:", err);
  }
};

const testLabels = ["M52R"];

const printTestLabels = async () => {
  try {
    // Ask for the device only once
    const device = await navigator.usb.requestDevice({
      filters: [{ vendorId: 0x20d1 }],
    });

    await device.open();
    if (device.configuration === null) {
      await device.selectConfiguration(1);
    }
    await device.claimInterface(0);

    const iface = device.configuration.interfaces[0].alternate;
    const outEndpoint = iface.endpoints.find((e) => e.direction === "out");

    for (const label of labels.value) {
      if (printedLabels.value.includes(label)) continue;

      // Find and set current printing row
      const rowIndex = excelData.value.findIndex(
        (row) => `${row[0]}${row[1]}${row[2]}` === label
      );
      currentPrintingRow.value = rowIndex;

      // Scroll to the current row
      if (tableRef.value) {
        const rowElement = tableRef.value.querySelector(
          `tr:nth-child(${rowIndex + 2})`
        );
        if (rowElement) {
          rowElement.scrollIntoView({ behavior: "smooth", block: "center" });
        }
      }

      const cmds = [
        "DENSITY 20",
        "SIZE 30 mm,20 mm",
        "CLS",
        `TEXT 40,60,"7",0,3,3,"${label}"`,
        "PRINT 1",
      ];
      const data = new TextEncoder().encode(cmds.join("\r\n") + "\r\n");
      await device.transferOut(outEndpoint.endpointNumber, data);
      printedLabels.value.push(label);
      console.log(`✅ Printed: ${label}`);
      console.log(`Remaining labels: ${remainingLabels.value}`);
      await new Promise((resolve) => setTimeout(resolve, 400));
    }

    currentPrintingRow.value = null;
    await device.close();
  } catch (err) {
    console.error("Print error:", err);
    currentPrintingRow.value = null;
  }
};

const printCustomLabels = async () => {
  try {
    if (!customLabelsInput.value.trim()) return;

    // Split input by commas or spaces and clean up
    const customLabels = customLabelsInput.value
      .split(/[, ]+/)
      .map((label) => label.trim())
      .filter((label) => label);

    // Ask for the device only once
    const device = await navigator.usb.requestDevice({
      filters: [{ vendorId: 0x20d1 }],
    });

    await device.open();
    if (device.configuration === null) {
      await device.selectConfiguration(1);
    }
    await device.claimInterface(0);

    const iface = device.configuration.interfaces[0].alternate;
    const outEndpoint = iface.endpoints.find((e) => e.direction === "out");

    for (const label of customLabels) {
      // Find and set current printing row
      const rowIndex = excelData.value.findIndex(
        (row) => `${row[0]}${row[1]}${row[2]}` === label
      );
      currentPrintingRow.value = rowIndex;

      // Scroll to the current row
      if (tableRef.value && rowIndex !== -1) {
        const rowElement = tableRef.value.querySelector(
          `tr:nth-child(${rowIndex + 2})`
        );
        if (rowElement) {
          rowElement.scrollIntoView({ behavior: "smooth", block: "center" });
        }
      }

      const cmds = [
        "DENSITY 20",
        "SIZE 30 mm,20 mm",
        "CLS",
        `TEXT 40,60,"7",0,3,3,"${label}"`,
        "PRINT 1",
      ];
      const data = new TextEncoder().encode(cmds.join("\r\n") + "\r\n");
      await device.transferOut(outEndpoint.endpointNumber, data);
      printedLabels.value.push(label);
      console.log(`✅ Printed: ${label}`);
      await new Promise((resolve) => setTimeout(resolve, 400));
    }

    currentPrintingRow.value = null;
    await device.close();
    customLabelsInput.value = ""; // Clear input after printing
  } catch (err) {
    console.error("Print error:", err);
    currentPrintingRow.value = null;
  }
};
</script>
