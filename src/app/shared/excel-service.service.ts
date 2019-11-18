import { Injectable } from '@angular/core';
import * as FileSaver from 'file-saver';
import * as XLSL from 'xlsx';


const EXCEL_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8'
const EXCEL_EXTENTION = '.xlsx';

@Injectable({
  providedIn: 'root'
})
export class ExcelService {

  constructor() { }

  public exportAsExcelFile(json:any[],excelFileName: string):void{
    const workSheet: XLSL.WorkSheet = XLSL.utils.json_to_sheet(json);

    const workBook: XLSL.WorkBook = {
      Sheets: {
        'data': workSheet
      },
      SheetNames: ['data']
    };

    const excelBuffer: any = XLSL.write(workBook, {
      bookType: 'xlsx',
      type: 'array'
    });
    this.saveAsExcelFile(excelBuffer,excelFileName);
  }
  private saveAsExcelFile(buffer: any, fileName: string): void{
    const data : Blob = new Blob([buffer],{type:EXCEL_TYPE});

    FileSaver.saveAs(data,fileName+'_export_'+new Date().getTime()+EXCEL_EXTENTION);
  }


}
