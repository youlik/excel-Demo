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
    import XLSXS from 'xlsx-style'
    import FileSaver from 'file-saver'
    export default {
        name: "excelDemo",
        methods:{
            toExcel(){
                let wb = XLSX.utils.table_to_book(this.$refs.table,{sheet:'分组表'})
                this.setExlStyle(wb['Sheets']['分组表'])
                this.addRangeBorder(wb['Sheets']['分组表']['!merges'],wb['Sheets']['分组表'])
                var ws = XLSXS.write(wb, {
                    type: "buffer"
                });
                try {
                    FileSaver.saveAs(
                        new Blob([ws], { type: "application/octet-stream" }),
                        `5.xlsx`
                    );
                } catch (e) {
                    if (typeof console !== "undefined") console.log(e, ws);
                }
                return ws;
            },
            setExlStyle(data) {
                let borderAll = {  //单元格外侧框线
                    top: {
                        style: 'thin',
                    },
                    bottom: {
                        style: 'thin'
                    },
                    left: {
                        style: 'thin'
                    },
                    right: {
                        style: 'thin'
                    }
                };
                data['!cols'] = [];
                for (let key in data) {
                    if (data[key] instanceof Object) {
                        data[key].s = {
                            border: borderAll,
                            alignment: {
                                horizontal: 'center',   //水平居中对齐
                                vertical:'center'
                            },
                            font:{
                                sz:11
                            },
                            bold:true,
                            numFmt: 0
                        }
                        data['!cols'].push({wpx: 115});
                    }
                }
                return data;
            },
            addRangeBorder(range,ws){
                let arr = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"];
                range.forEach(item=>{
                    console.log(item)
                    let startRowNumber = Number(item.s.c),
                        endRowNumber = Number(item.e.c);
                    for(let i = startRowNumber;i<= endRowNumber;i++){
                        for( let j = Number(item.s.r)+1;j<=Number(item.e.r);j++){
                            ws[`${arr[i]}${j+1}`] = {s:{border:{top:{style:'thin'}, left:{style:'thin'},bottom:{style:'thin'},right:{style:'thin'}}}};
                        }
                    }

                })
                console.log(ws)
                return ws;

            },
        }
    }
</script>

<style scoped>
    .btn{
        width: 100px;
        line-height: 50px;

    }
</style>