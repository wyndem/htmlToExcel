function toExcel(tableId,names,title) {

    let Worksheet = '';
    let htmlList = '';
    let HRefList = '';


    for (let i = 0; i < tableId.length; i++) {
        let tableHtml = document.getElementById(tableId[i]);
        // 使用outerHTML属性获取整个table元素的HTML代码（包括<table>标签），然后包装成一个完整的HTML文档，设置charset为urf-8以防止中文乱码
        let appendHtml = tableHtml.innerHTML;
        if (appendHtml.length < 600) {
            return;
        }
        let info = {
            title :names[i] || 'sheet' + i + 1,
            className: tableId[i]
        }
        Worksheet += '<x:ExcelWorksheet><x:Name>' + info.title + '</x:Name><x:WorksheetSource HRef="' + info.className + '.htm"/></x:ExcelWorksheet>';
        htmlList += sheetExcelTable(info.className, appendHtml);
        HRefList += '<o:File HRef="' + info.className + '.htm"/>';
    }


    let txt = 'MIME-Version: 1.0\n' + 'X-Document-Type: Workbook\n' + 'Content-Type: multipart/related; boundary="----=_NextPart_dummy"\n' + '\n' + '------=_NextPart_dummy\n' + 'Content-Location: WorkBook.htm\n' + 'Content-Type: text/html; charset=utf-8\n' + '\n' + '<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">\n' + '<head>\n' + '<meta name="Excel Workbook Frameset">\n' + '<meta http-equiv="Content-Type" content="text/html; charset=utf-8">\n' + '<link rel="File-List" href="filelist.xml">\n' + '<!--[if gte mso 9]><xml>\n' + ' <x:ExcelWorkbook>\n' + '<x:ExcelWorksheets>' + Worksheet + '</x:ExcelWorksheets>\n' + '    <x:ActiveSheet>0</x:ActiveSheet>\n' + ' </x:ExcelWorkbook>\n' + '</xml><![endif]-->\n' + '</head>\n' + '<frameset>\n' + '    <frame src="sheet0.htm" name="frSheet">\n' + '    <noframes><body><p>This page uses frames, but your browser does not support them.</p></body></noframes>\n' + '</frameset>\n' + '</html>'
    txt += htmlList;
    txt += 'Content-Location: filelist.xml\n' + 'Content-Type: text/xml; charset="utf-8"\n' + '\n' + '<xml xmlns:o="urn:schemas-microsoft-com:office:office">\n' + '    <o:MainFile HRef="../WorkBook.htm"/>\n' + HRefList + '<o:File HRef="filelist.xml"/>\n' + '</xml>\n' + '------=_NextPart_dummy--';

    // // 实例化一个Blob对象，其构造函数的第一个参数是包含文件内容的数组，第二个参数是包含文件类型属性的对象
    let blob = new Blob([txt], {
        type: "text/plain;charset=utf-8",
    }); //application/octet-stream
    //也可以用js创建一个a标签
    let a = document.createElement('a');
    // 利用URL.createObjectURL()方法为a元素生成blob URL
    a.href = URL.createObjectURL(blob);
    // 设置文件名
    a.download = title  + ".xls";
    a.click();


}

function sheetExcelTable(tableId,html){
    var template = '\n------=_NextPart_dummy\n'+ 'Content-Location: '+ tableId +'.htm\n' + 'Content-Type: text/html; charset="utf-8"\n' + '\n' +  '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><meta http-equiv="Content-Type" charset="utf-8"><!--[if gte mso 9]><xml><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></xml><![endif]--></head><body><table  border="1" cellpadding="0" cellspacing="0">{table}</table></body></html>' + '\n------=_NextPart_dummy\n' ;
    function format(s, c) {
        return s.replace(/{(\w+)}/g, function (m, p) {
            return c[p];
        });
    }
    var ctx = {
        worksheet: name || 'worksheet',
        table: html
    };
    return  format(template, ctx)
}