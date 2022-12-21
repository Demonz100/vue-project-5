<template>
    <main>
        <p>Export Excel</p>
        <button @click="exportToExcel">Download Excel</button>

        <br><br><br>
        <p v-if="isLoading">Loading...</p>
        <ul v-else>
            <li v-for="data in datas">{{data}}</li>
        </ul>
        
    </main>
</template>
  
<script setup lang="ts">
import { exportSheet } from '@/exportExcel'
import { ref } from 'vue';

const datas = ref([])

const isLoading = ref(false)

const getDatas = async () => {
    isLoading.value = true

    const res = await fetch('https://reqres.in/api/users?delay=2')
    .then(response => response.json())
    .then(data => {datas.value = data.data})
    .catch(error => console.log(error))

    isLoading.value = false
}

getDatas()

const exportToExcel = () => {
    exportSheet(datas.value, 'users'); 
}

</script>
  