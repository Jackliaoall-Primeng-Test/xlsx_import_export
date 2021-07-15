import { Component } from '@angular/core';
import * as XLSX from 'xlsx';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {

  data = [
    ['1', 'a', 'aa'],
    ['2', 'b', 'bb'],
    ['3', 'c', 'cc']
  ];

  inputdata = [];

  constructor() {}

  ngOnInit(): void {
  }

  // 匯出
  daochu(){
    /* generate worksheet */
    const ws: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(this.data);
    const ws2: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(this.data);

    /* generate workbook and add the worksheet */
    const wb: XLSX.WorkBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.utils.book_append_sheet(wb, ws2, 'Sheet2');

    console.log(wb)
    /* save to file */
    XLSX.writeFile(wb, 'SheetJS.xlsx');
  }

  // 匯入
  public daoru(evt: any): void {
    /* wire up file reader */
    const target: DataTransfer = <DataTransfer>(evt.target);
    if (target.files.length !== 1) throw new Error('Cannot use multiple files');
    const reader: FileReader = new FileReader();
    reader.onload = (e: any) => {
      /* read workbook */
      const bstr: string = e.target.result;
      const wb: XLSX.WorkBook = XLSX.read(bstr, {type: 'binary'});

      /* grab first sheet */
      const wsname: string = wb.SheetNames[0];
      const ws: XLSX.WorkSheet = wb.Sheets[wsname];

      /* save data */
      this.inputdata = (XLSX.utils.sheet_to_json(ws, {header: 1}));
      console.log(this.inputdata);

      evt.target.value = ""; // 清空
    };
    reader.readAsBinaryString(target.files[0]);

    console.log(reader.result);
  }
}
