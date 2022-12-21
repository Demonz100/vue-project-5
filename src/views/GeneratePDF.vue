<template>
    <div>
    <br><br>
    <!-- <ReportA4 :reportDatas="tempData"/> -->
    <button @click="generatePDF()">Download Excel</button>
    
    <vue-html2pdf
        :show-layout="false"
        :float-layout="true"
        :enable-download="true"
        :preview-modal="true"
        :paginate-elements-by-height="1400"
        filename="nightprogrammerpdf"
        :pdf-quality="2"
        :manual-pagination="false"
        pdf-format="a4"
        pdf-orientation="portrait"
        ref="html2Pdf"
    >
      <section slot="pdf-content">
        <ReportA4 :reportDatas="tempData"/>
      </section>
    </vue-html2pdf>
    </div>
</template>

<script setup lang="ts">
import { defineAsyncComponent, ref } from 'vue';
import VueHtml2pdf from "vue-html2pdf";
const ReportA4 = defineAsyncComponent(()=>import('@/components/ReportA4.vue'))

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

const html2Pdf = ref(null)
const generatePDF = () => {
  console.log(html2Pdf.value);
  
}
</script>

<style scoped>

</style>