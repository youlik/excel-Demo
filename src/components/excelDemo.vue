<template>
    <div>
        <button class="btn" @click="toExcel">导出</button>
        <table ref="table" border="1" style="width: 800px">
            <tr>
                <th>班级</th>
                <th>组别</th>
                <th>姓名</th>
            </tr>
            <tr>
                <td rowspan="3">一班</td>
                <td rowspan="2">第一组</td>
                <td>小明</td>
            </tr>
            <tr>
                <td>小吴</td>
            </tr>
            <tr>
                <td>第二组</td>
                <td>小赵</td>
            </tr>
        </table>
    </div>
</template>

<script>
    import XLSX from 'xlsx'
    // import XLSXS from 'xlsx-style'
    import FileSaver from 'file-saver'
    export default {
        name: "excelDemo",
        methods:{
            toExcel(){
                let wb = XLSX.utils.table_to_book(this.$refs.table,{sheet:'分组表'})
                console.log(wb)
                var ws = XLSX.write(wb, {
                    bookType: "xlsx",
                    bookSST: true,
                    type: "array"
                });
                console.log(ws)
                try {
                    FileSaver.saveAs(new Blob([ws],{type: "application/octet-stream"}),'翻斗花园')
                } catch (e) {
                    if(typeof console !== 'undefined') console.log(e,ws)
                }
                return ws
            }
        }
    }
</script>

<style scoped>
    .btn{
        width: 100px;
        line-height: 50px;

    }
</style>