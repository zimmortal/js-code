// 将页面中的table转成EXCEL 下载下来
class ExportExcel {
  constructor () {
    this.idTmr = null
    this.uri = 'data:application/vnd.ms-excel;base64,'
    this.template = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><meta charset="UTF-8"><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body><table>{table}</table></body></html>'
  }

  getBrowser () {
    const explorer = window.navigator.userAgent
    //ie
    if (explorer.indexOf('MSIE') >= 0) {
      return 'ie'
    }
    //firefox
    else if (explorer.indexOf('Firefox') >= 0) {
      return 'Firefox'
    }
    //Chrome
    else if (explorer.indexOf('Chrome') >= 0) {
      return 'Chrome'
    }
    //Opera
    else if (explorer.indexOf('Opera') >= 0) {
      return 'Opera'
    }
    //Safari
    else if (explorer.indexOf('Safari') >= 0) {
      return 'Safari'
    }
  };

  exports (tableid) {
    if (this.getBrowser() === 'ie') {
      const curTbl = document.getElementById(tableid)
      let oXL = new ActiveXObject('Excel.Application')
      const oWB = oXL.Workbooks.Add()
      const xlSheet = oWB.Worksheets(1)
      const sel = document.body.createTextRange()
      sel.moveToElementText(curTbl)
      sel.select()
      sel.execCommand('Copy')
      xlSheet.Paste()
      oXL.Visible = true

      try {
        var fname = oXL.Application.GetSaveAsFilename('Excel.xls', 'Excel Spreadsheets (*.xls), *.xls')
      } catch (e) {
        alert(e)
      } finally {
        oWB.SaveAs(fname)
        oWB.Close(savechanges = false)
        oXL.Quit()
        oXL = null
        this.idTmr = window.setInterval('Cleanup();', 1)
      }
    } else {
      this.openExport(tableid)
    }
  };

  openExport (table, name) {
    if (!table.nodeType) {
      table = document.getElementById(table)
    }
    const ctx = {
      worksheet: name || 'Worksheet',
      table: table.innerHTML
    }
    window.location.href = this.uri + this.base64(this.format(this.template, ctx))
  };

  base64 (s) {
    return window.btoa(unescape(encodeURIComponent(s)))
  };

  format (s, c) {
    return s.replace(/{(\w+)}/g, function (m, p) {
      return c[p]
    })
  };
}

const exportExcel = new ExportExcel()
exportExcel.exports('对应的ID')
