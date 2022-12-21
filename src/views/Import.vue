<template>
  <main>
    <input type="file" accept=".xls,.xlsx" @change="importFromExcel">
    <br><br>

    <table v-if="importDatas.length != 0">
      <tr>
        <th v-for="tableHead in tableHeads">{{ tableHead }}</th>
      </tr>
      <tr v-for="data in importDatas" :key="data.id">
        <td>{{ data.id }}</td>
        <td>{{ data.email }}</td>
        <td>{{ data.first_name }}</td>
        <td>{{ data.last_name }}</td>
        <td>{{ data.avatar }}</td>
      </tr>
    </table>
  </main>
</template>

<script setup lang="ts">
import { ref } from 'vue';
import { importSheet } from '@/importExcel'

// interface ImportedDatas {
//   id: number,
//   email: string,
//   first_name: string,
//   last_name: string,
//   avatar: string
// }

const importDatas = ref<object[]>([])
const tableHeads = ref<string[]>([])

const importFromExcel = async (event: Event) => {
  await importSheet(event).then((res) => { importDatas.value = res })
  tableHeads.value = Object.keys(importDatas.value[0])
}




</script>

<style scoped>
table {
  border-collapse: collapse;
  border-bottom: 1px solid rgb(0, 0, 0, 0.2);

  max-width: 2000px;
  width: 100%;

  margin: 10px auto 20px auto;
}

th {
  background-color: #aecbeb;
  border: 1px solid black;
  text-align: center;
  font-size: 15px;
}

td {
  border-bottom: 1px solid black;
  border: 1px solid rgb(0, 0, 0, 0.2);
}
</style>
