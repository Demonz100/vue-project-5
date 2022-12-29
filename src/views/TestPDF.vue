<template>
    <br>
    <button @click="exportToPDF">Export to PDF</button>
    <br><br><br>
    <div ref="document">
        <div id="element-to-convert">
            <ReportA4 :report-datas="tempData"></ReportA4>
        </div>
    </div>
</template>

<script setup lang="ts">
import html2pdf from 'html2pdf.js'
import { defineAsyncComponent, ref } from 'vue';
const ReportA4 = defineAsyncComponent(()=>import('@/components/ReportA4.vue'))

const exportToPDF = () => {
    html2pdf(document.getElementById("element-to-convert"), {
        filename: "i-was-html.pdf",
        image: { 
          type: "jpeg", 
          quality: 1
        },
        html2canvas: {
          scale: 3,
          letterRendering: true
        },
        jsPDF: { 
          unit: "cm", 
          format: "a4", 
          orientation: "portrait"
        }
    });
}

const tempData = ref({
    name: 'Ponna Vanda',
    dob: '18-04-1996',
    nationality: 'Cambodian',
    national_id: '1234567889999',
    home_tel: '',
    mobile_tel: '096-531-4142',
    office: '',
    email: '',
    mobile_banking: false,
    internet_banking: true,
    subscribe: true,
    unsubscribe: false,
    lock: false,
    unlock: true,
    reset_password: false,
    re_active: false,
    amend: false
})
</script>

<style scoped>
table{
    border-top-width: 1px;
    border-right-width: 0px;
    border-bottom-width: 0px;
    border-left-width: 1px;
    border-color: black;
    border-style: solid;
    border-collapse: collapse;
}
th, td {
    border-top-width: 0px;
    border-right-width: 1px;
    border-bottom-width: 1px;
    border-left-width: 0px;
    border-color: black;
    border-style: solid;
}
</style>