<template>
    <main>
        <p>Export Excel</p>
        <button @click="exportToExcel">Download Excel</button>

        <br><br><br>
        <ul>
            <li v-for="data in table1">{{ data }}</li>
        </ul>

        <br><br><br>
        <ul>
            <li v-for="data in table2">{{ data }}</li>
        </ul>
    </main>
</template>
  
<script setup lang="ts">
import { exportSheet } from '@/exportExcel'
import { ref } from 'vue';

const table1 = ref([{ "id": 1, "email": "george.bluth@reqres.in", "first_name": "George", "last_name": "Bluth", "avatar": "https://reqres.in/img/faces/1-image.jpg", "age": 18 },
{ "id": 2, "email": "janet.weaver@reqres.in", "first_name": "Janet", "last_name": "Weaver", "avatar": "https://reqres.in/img/faces/2-image.jpg", "age": 23 },
{ "id": 3, "email": "emma.wong@reqres.in", "first_name": "Emma", "last_name": "Wong", "avatar": "https://reqres.in/img/faces/3-image.jpg", "age": 27 },
{ "id": 4, "email": "eve.holt@reqres.in", "first_name": "Eve", "last_name": "Holt", "avatar": "https://reqres.in/img/faces/4-image.jpg", "age": 26 },
{ "id": 5, "email": "charles.morris@reqres.in", "first_name": "Charles", "last_name": "Morris", "avatar": "https://reqres.in/img/faces/5-image.jpg", "age": 30 },
{ "id": 6, "email": "tracey.ramos@reqres.in", "first_name": "Tracey", "last_name": "Ramos", "avatar": "https://reqres.in/img/faces/6-image.jpg", "age": 26 }])

const table2 = ref([{ "no": 1, "type": "Saving Account", "account_no": "2-120-00000135-5", "account_nickname": "sothon", "balance": 999044.33, "currency": 'USD', "last_transaction_date_and_time": '14 Dec 2022, 04:31:04 PM', "acc.status": 'active', "account_holder": 'Thoeurn Sothon' },
{ "no": 2, "type": "Saving Account", "account_no": "2-120-00000135-5", "account_nickname": "sothon", "balance": 999044.33, "currency": 'USD', "last_transaction_date_and_time": '14 Dec 2022, 04:31:04 PM', "acc.status": 'active', "account_holder": 'Thoeurn Sothon' },
{ "no": 3, "type": "Saving Account", "account_no": "2-120-00000135-5", "account_nickname": "sothon", "balance": 999044.33, "currency": 'USD', "last_transaction_date_and_time": '14 Dec 2022, 04:31:04 PM', "acc.status": 'active', "account_holder": 'Thoeurn Sothon' },
{ "no": 4, "type": "Saving Account", "account_no": "2-120-00000135-5", "account_nickname": "sothon", "balance": 999044.33, "currency": 'USD', "last_transaction_date_and_time": '14 Dec 2022, 04:31:04 PM', "acc.status": 'active', "account_holder": 'Thoeurn Sothon' },
{ "no": 5, "type": "Saving Account", "account_no": "2-120-00000135-5", "account_nickname": "sothon", "balance": 999044.33, "currency": 'USD', "last_transaction_date_and_time": '14 Dec 2022, 04:31:04 PM', "acc.status": 'active', "account_holder": 'Thoeurn Sothon' }])

const exportData = ref([
    {
        data: table1.value,
        sheetName: 'users',
        calcName: 'Average Age: ',
        calcCol: 'F8',
        calcFormula: 'AVERAGE(F2:F7)',
        calcResult: 180
    },
    {
        data: table2.value,
        calcName: 'Total: ',
        calcResult: 4995221.65
    },
    {
        data: table1.value,
        sheetName: 'users',
        calcName: 'Average Age: ',
        calcCol: 'F8',
        calcFormula: 'AVERAGE(F2:F7)',
        calcResult: 180
    }
])

const exportToExcel = () => {
    exportSheet(exportData.value);
}

</script>
  