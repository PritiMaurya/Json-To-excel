import { Component } from '@angular/core';
import * as XLSX from 'xlsx';
import * as $ from 'jquery';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent {
  title = 'convertJsonToExcelDemo';

  /* XLS Head Columns */
  xlsHeader = ["EmployeeID", "Full Name"];

  /* XLS Rows Data */
  xlsRows = [{
    "EmployeeID": "EMP001",
    "FullName": "Jolly"
  },
  {
    "EmployeeID": "EMP002",
    "FullName": "Macias"
  },
  {
    "EmployeeID": "EMP003",
    "FullName": "Lucian"
  },
  {
    "EmployeeID": "EMP004",
    "FullName": "Blaze"
  },
  {
    "EmployeeID": "EMP005",
    "FullName": "Blossom"
  },
  {
    "EmployeeID": "EMP006",
    "FullName": "Kerry"
  },
  {
    "EmployeeID": "EMP007",
    "FullName": "Adele"
  },
  {
    "EmployeeID": "EMP008",
    "FullName": "Freaky"
  },
  {
    "EmployeeID": "EMP009",
    "FullName": "Brooke"
  },
  {
    "EmployeeID": "EMP010",
    "FullName": "FreakyJolly.Com"
  }
  ];

  convertJsonToExcel() {
    var createXLSLFormatObj = [];

    createXLSLFormatObj.push(this.xlsHeader);

    this.xlsRows.forEach(obj => {
      console.log('obj ', obj);
      var innerRowData = [];
      $("tbody").append('<tr><td>' + obj.EmployeeID + '</td><td>' + obj.FullName + '</td></tr>');
      $.each(obj, function (ind, val) {
        console.log('val ', val);
        innerRowData.push(val);
      });
      createXLSLFormatObj.push(innerRowData);
    });


    /* File Name */
    var filename = "test.xlsx";

    /* Sheet Name */
    var ws_name = "TestSheet";

    if (typeof console !== 'undefined') console.log(new Date());
    var wb = XLSX.utils.book_new(),
      ws = XLSX.utils.aoa_to_sheet(createXLSLFormatObj);

    /* Add worksheet to workbook */
    XLSX.utils.book_append_sheet(wb, ws, ws_name);

    /* Write workbook and Download */
    if (typeof console !== 'undefined') console.log(new Date());
    XLSX.writeFile(wb, filename);
  }


}
