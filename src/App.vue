<template>
  <div id="app">
    <excel-demo />
  </div>
</template>

<script>
import ExcelDemo from "./components/excelDemo";

export default {
  name: 'App',
  components: {
    ExcelDemo,
  }
}
</script>

<style>
#app {
  font-family: Avenir, Helvetica, Arial, sans-serif;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  text-align: center;
  color: #2c3e50;
  margin-top: 60px;
}
</style>
