
/*
    万能流  application/octet-stream
    word文件  application/msword
    excel文件  application/vnd.ms-excel
    txt文件  text/plain
    图片文件  image/png、jpeg、gif、bmp
 */
function downloadByBlob(fileName, text) {
    let a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([text], { type: "application/octet-stream" }));
    a.download = fileName || 'Blob导出测试.txt';
    a.click();
    a.remove();
    URL.revokeObjectURL(a.href);
}


// fileName 文件名称
// columns 列名属性 主要用来匹配后台数据比如 name ,遍历tableDatas的时候 item.name 获取对应列的值
// columnNames 列标题名
// tableDatas 传递过来的参数值
function exportExcel(fileName, columns,columnNames, tableDatas) {
    //列名
    let columnHtml = "";
    columnHtml += "<tr style=\"text-align: center;\">\n";
    for (let key in columnNames) {
        columnHtml += "<td style=\"font-weight:bold;background-color:#bad5fd\">" + columnNames[key] + "</td>\n";
    }
    columnHtml += "</tr>\n";

    //数据
    let dataHtml = "";
    for (let data of tableDatas) {
        dataHtml += "<tr style=\"text-align: center;\">\n";
        for (let key in columns) {
            dataHtml += "<td>" + data[columns[key]] + "</td>\n";
        }
        dataHtml += "</tr>\n";
    }

    //完整拼接Excel内容的html
    let excelHtml = "<html xmlns:o=\"urn:schemas-microsoft-com:office:office\"\n" +
        "      xmlns:x=\"urn:schemas-microsoft-com:office:excel\"\n" +
        "      xmlns=\"http://www.w3.org/TR/REC-html40\">\n" +
        "<head>\n" +
        "   <!-- 加这个，其他单元格带边框 -->" +
        "   <xml>\n" +
        "        <x:ExcelWorkbook>\n" +
        "            <x:ExcelWorksheets>\n" +
        "                <x:ExcelWorksheet>\n" +
        "                    <x:Name></x:Name>\n" +
        "                    <x:WorksheetOptions>\n" +
        "                        <x:DisplayGridlines/>\n" +
        "                    </x:WorksheetOptions>\n" +
        "                </x:ExcelWorksheet>\n" +
        "            </x:ExcelWorksheets>\n" +
        "        </x:ExcelWorkbook>\n" +
        "   </xml>\n" +
        "   <style>td{font-family: \"宋体\";}</style>\n" +
        "</head>\n" +
        "<body>\n" +
        "<table border=\"1\">\n" +
        "    <thead>\n" +
        columnHtml +
        "    </thead>\n" +
        "    <tbody>\n" +
        dataHtml +
        "    </tbody>\n" +
        "</table>\n" +
        "</body>\n" +
        "</html>";

    //浏览器下载为本地文件
    downloadByBlob((fileName || "导出Excel") + ".xls", excelHtml);
}