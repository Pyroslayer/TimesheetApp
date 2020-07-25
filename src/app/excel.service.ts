import { Injectable } from '@angular/core';
import {saveAs} from 'file-saver';
import { Workbook } from 'exceljs';
import { FormGroup, FormArray } from '@angular/forms';
import { DatePipe } from '@angular/common';
import { from } from 'rxjs/internal/observable/from';
import { map } from 'rxjs/operators';
import { EmailServiceService } from './email-service.service';

@Injectable({
  providedIn: 'root'
})
export class ExcelService {

  workbook;
  wts;
  rowArray;
  colArray;
  othersColArray;
  othersRowArray;
  finalRow;
  totalRowCol;
  othersTotalRowCol;
  finalTotal;
  dark = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: {argb: 'C0C0C0'}
  };
  
  days = ['saturday', 'sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday'];
  constructor(private datePipe: DatePipe,
    private email: EmailServiceService) {
    this.workbook = new Workbook();
    this.wts = this.workbook.addWorksheet('WeeklyTimeSheet');
   }

  generate(formGroup: FormGroup, list: any[]){
    this.rowArray = list['rowArray'];
    this.colArray = list['colArray'];
    this.othersRowArray = list['othersRowArray'];
    this.othersColArray = list['othersColArray'];
    this.finalRow = list['finalRow'];
    this.totalRowCol = list['totalRowCol'];
    this.othersTotalRowCol = list['othersTotalRowCol'];
    this.finalTotal = list['finalTotal'];
    
    this.generateRows(15);
    this.mergeCells();
    this.generateLayout(formGroup);
    this.setColumnWidth();
    this.setRowHeight();
    this.generateExcel(formGroup.get('dateTo').value, formGroup.get('empInitials').value);
    

    
  }

  generateBlob(formGroup: FormGroup, list: any[]){
    this.rowArray = list['rowArray'];
    this.colArray = list['colArray'];
    this.othersRowArray = list['othersRowArray'];
    this.othersColArray = list['othersColArray'];
    this.finalRow = list['finalRow'];
    this.totalRowCol = list['totalRowCol'];
    this.othersTotalRowCol = list['othersTotalRowCol'];
    this.finalTotal = list['finalTotal'];
    
    this.generateRows(15);
    this.mergeCells();
    this.generateLayout(formGroup);
    this.setColumnWidth();
    this.setRowHeight();
    this.blobExcel(formGroup.get('dateTo').value, formGroup.get('empInitials').value);

    
    
  }

  generateLayout(formGroup: FormGroup){

    let titleRow = this.wts.getRow(2).getCell(1);
    titleRow.value = 'TeraSystem Inc.'
    titleRow.font = { name: 'Denmark', family: 4, size: 14  };
    titleRow.alignment = { vertical: 'middle', horizontal: 'center' };

    let subtitleRow = this.wts.getRow(3).getCell(1);
    subtitleRow.value = 'WeeklyTimeSheet';    
    subtitleRow.font = { name: 'Arial', size: 10};
    subtitleRow.alignment = { vertical: 'middle', horizontal: 'center' };

    let row5 = this.wts.getRow(5);    
    row5.getCell(1).value = 'Employee No.:';
    row5.getCell(1).font = { name: 'Arial', size: 10, italic: true };
    row5.getCell(1).alignment = {vertical: 'bottom', horizontal: 'right'};
    row5.getCell(3).value = formGroup.get('empNo').value;
    row5.getCell(3).font = { name: 'Arial', size: 10, bold: true };
    row5.getCell(3).alignment = {vertical: 'bottom', horizontal: 'right'};
    row5.getCell(4).font = { name: 'Arial', size: 10, italic: true };
    row5.getCell(4).alignment = {vertical: 'bottom', horizontal: 'right'};
    row5.getCell(4).value = 'Group:';
    row5.getCell(5).font = { name: 'Arial', size: 10, bold: true };
    row5.getCell(5).alignment = {vertical: 'top', horizontal: 'left'};
    row5.getCell(5).value = formGroup.get('group').value;

    let row6 = this.wts.getRow(6);   
    row6.getCell(1).value = 'Employee Name:';
    row6.getCell(1).font = { name: 'Arial', size: 10, italic: true };
    row6.getCell(1).alignment = {vertical: 'bottom', horizontal: 'right'};
    row6.getCell(3).font = { name: 'Arial', size: 10, bold: true };
    row6.getCell(3).value = formGroup.get('empName').value;
    row6.getCell(6).value = 'Month:';
    row6.getCell(6).font = { name: 'Arial', size: 10};
    row6.getCell(7).font = { name: 'Arial', size: 10, bold: true };
    row6.getCell(7).alignment = {vertical: 'middle', horizontal: 'center'};
    row6.getCell(7).value = new Date(formGroup.get('dateTo').value).toLocaleString('en', { month: 'long'  });

    let row7 = this.wts.getRow(7);
    row7.getCell(1).font = { name: 'Arial', size: 10, italic: true };
    row7.getCell(1).alignment = {vertical: 'bottom', horizontal: 'right'};
    row7.getCell(1).value = 'Signature:';
    row7.getCell(6).value = 'Year:';
    row7.getCell(6).font = { name: 'Arial', size: 10};
    row7.getCell(7).font = { name: 'Arial', size: 10, bold: true };
    row7.getCell(7).alignment = {vertical: 'middle', horizontal: 'center'};
    row7.getCell(7).value = new Date(formGroup.get('dateTo').value).getFullYear();

    let row9 = this.wts.getRow(9);
    var dateCounter = 0;
    for (var x = 6; x <= 12 ; x++){
      row9.getCell(x).font = { name: 'Arial', size: 10, bold: true };
      row9.getCell(x).alignment = {vertical: 'middle', horizontal: 'center'};

      var newDate = new Date(formGroup.get('dateFrom').value);
      newDate.setDate(newDate.getDate() + dateCounter);
      row9.getCell(x).value = this.datePipe.transform(newDate, 'MM/dd/yyyy');
      dateCounter++;
    }

    var days = ['Saturday', 'Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Total'];
    // Add table Border
    for (var row = 10; row <= 15; row++ ){
      for (var col = 1; col <= 13; col++){
        this.wts.getRow(row).getCell(col).border = {
          top: {style: 'thin'},
          left: {style: 'thin'},
          bottom: {style: 'thin'},
          right: {style: 'thin'}
        }
      }
    };
    
    for (var col = 1; col <= 13; col++){
      this.wts.getRow(14).getCell(col).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: {argb: '1FB714'}
      };
    }

    let row10 = this.wts.getRow(10);
    row10.font = { name: 'Arial', size: 10, italic: true };
    row10.alignment = {vertical: 'bottom', horizontal: 'center'};
    row10.getCell(1).value = 'Period Covered :';
    row10.getCell(6).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: {argb: 'C0C0C0'}
    };
    row10.getCell(7).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: {argb: 'C0C0C0'}
    };

    var counter = 6;
    days.forEach(day => {
      row10.getCell(counter).value = day;
      counter++;
    });

    let row11 = this.wts.getRow(11);
    row11.font = { name: 'Arial', size: 10, italic: true };
    row11.alignment = {vertical: 'bottom', horizontal: 'right'};
    row11.getCell(1).value = 'From :';
    row11.getCell(2).value = this.datePipe.transform(new Date(formGroup.get('dateFrom').value), 'dd-MMM-yy');
    row11.getCell(5).value = 'Time In (hrs):';
    row11.getCell(6).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: {argb: 'C0C0C0'}
    };
    row11.getCell(7).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: {argb: 'C0C0C0'}
    };

    
      
 

    let row12 = this.wts.getRow(12);
    row12.font = { name: 'Arial', size: 10, italic: true };
    row12.alignment = {vertical: 'bottom', horizontal: 'right'};
    row12.getCell(1).value = 'To :';
    row12.getCell(2).value = this.datePipe.transform(new Date(formGroup.get('dateTo').value), 'dd-MMM-yy');
    row12.getCell(5).value = 'Time Out (hrs):';
    row12.getCell(6).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: {argb: 'C0C0C0'}
    };
    row12.getCell(7).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: {argb: 'C0C0C0'}
    };

    let row13 = this.wts.getRow(13);
    row13.font = { name: 'Arial', size: 10, italic: true };
    row13.alignment = {vertical: 'bottom', horizontal: 'right'};
    row13.getCell(5).value = '# of Hours:';
    row13.getCell(6).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: {argb: 'C0C0C0'}
    };
    row13.getCell(7).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: {argb: 'C0C0C0'}
    };

    var cellCounter = 6; 
    var totalHours = 0;
    (formGroup.get('timeInOut') as FormArray).controls.forEach(element => {
      row11.getCell(cellCounter).value = element.get('timeIn').value;     
      row12.getCell(cellCounter).value = element.get('timeOut').value;
      totalHours += element.get('hours').value;
      var date = new Date(0, 0);
      date.setSeconds(element.get('hours').value * 60 * 60);
      row13.getCell(cellCounter).value = element.get('hours').value == 0 ? null : date.toTimeString().slice(0, 8);
      cellCounter++;
    });
    
    
    row13.getCell(13).value = ('00' + (totalHours - (totalHours % 1))).slice(-2) + ':' + ('00' + Math.round((totalHours % 1) * 60)).slice(-2) + ':00';

    this.generateTableTask(formGroup.get('timeSheet') as FormArray);
    this.generateTableOthers(formGroup.get('others') as FormArray);

    this.wts.addRow([]);
    for (var x = 6; x <= 13; x++){
      this.wts.lastRow.getCell(x).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: {argb: '90713A'}
      };
      
    }
    
    this.wts.lastRow.getCell(5).value = 'Total Hours';
    this.wts.lastRow.getCell(5).font = { name: 'Arial', size: 10, bold: true };
    this.wts.lastRow.getCell(5).alignment = {vertical: 'bottom', horizontal: 'right'};
    var counter = 6
    this.days.forEach(day => {
      this.wts.lastRow.getCell(counter).value = this.finalRow[day];
      this.wts.lastRow.getCell(counter).numFmt = '0.00';
      counter++;
    });
    this.wts.lastRow.getCell(counter).value = this.finalTotal;
    this.wts.lastRow.getCell(counter).numFmt = '0.00';

    this.addBorderRow();
    this.wts.addRow([]);
    this.wts.addRow([]);
    this.wts.addRow([]);
    this.wts.addRow([]);
    this.wts.mergeCells('A' + this.wts.rowCount + ':B' + this.wts.rowCount);
    this.wts.lastRow.getCell(1).value = 'Approved by:';
    this.wts.lastRow.getCell(1).font = { name: 'Arial', size: 10, bold: true, italic: true };
    this.wts.lastRow.getCell(1).alignment = {vertical: 'bottom', horizontal: 'right'};
    this.wts.addRow([]);
    this.wts.addRow([]);
    this.wts.lastRow.getCell(5).font = { name: 'Arial', size: 10, bold: true };
    this.wts.lastRow.getCell(5).alignment = {vertical: 'bottom', horizontal: 'center'};
    this.wts.lastRow.getCell(5).value = formGroup.get('empSupervisor').value;


  }

  generateTableTask(formArray: FormArray){
   

    let row14 = this.wts.getRow(14);
    row14.font = { name: 'Arial', size: 10, italic: true, bold: true };
    row14.alignment = {vertical: 'top', horizontal: 'left'};
    row14.getCell(1).value = 'Productive Hours';

    let row15 = this.wts.getRow(15);
    row15.font = { name: 'Arial', size: 10, italic: true, bold: true };
    row15.alignment = {vertical: 'bottom', horizontal: 'center'};
    row15.getCell(1).value = 'Client Name';
    row15.getCell(3).value = 'Project Name';
    row15.getCell(4).value = 'Task';
    row15.getCell(5).value = 'Activities:';
    row15.getCell(6).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: {argb: 'C0C0C0'}
    };
    row15.getCell(7).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: {argb: 'C0C0C0'}
    };

    this.addNewRowTask(formArray);

    this.wts.addRow([]);
    for (var x = 6; x <= 13; x++){
      this.wts.lastRow.getCell(x).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: {argb: 'FFFF99'}
      };
      
    }
    
    this.wts.lastRow.getCell(5).value = 'Total Hours Work';
    this.wts.lastRow.getCell(5).font = { name: 'Arial', size: 10, bold: true };
    this.wts.lastRow.getCell(5).alignment = {vertical: 'bottom', horizontal: 'right'};
    var counter = 6
    this.days.forEach(day => {
      this.wts.lastRow.getCell(counter).value = this.colArray[day];
      this.wts.lastRow.getCell(counter).numFmt = '0.00';
      counter++;
    });
    this.wts.lastRow.getCell(counter).value = this.totalRowCol;
    this.wts.lastRow.getCell(counter).numFmt = '0.00';

    this.addBorderRow();

    



  }

  generateTableOthers(formArray: FormArray){
    let rowOthers = this.wts.addRow([]);
    for (var x = 1; x <= 13; x++){
      rowOthers.getCell(x).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: {argb: 'FFCC00'}
      };
      
    }
    this.addBorderRow();
    this.wts.unMergeCells('A' + this.wts.rowCount);
    this.wts.mergeCells('A' + this.wts.rowCount + ':C' + this.wts.rowCount);
    rowOthers.font = { name: 'Arial', size: 10, italic: true, bold: true };
    rowOthers.alignment = {vertical: 'top', horizontal: 'left'};
    rowOthers.getCell(1).value = 'Others';

    let othersHeader = this.wts.addRow([]);
    othersHeader.font = { name: 'Arial', size: 10, italic: true, bold: true };
    othersHeader.alignment = {vertical: 'bottom', horizontal: 'center'};
    othersHeader.getCell(3).value = 'Type';
    othersHeader.getCell(4).value = 'Activity';
    othersHeader.getCell(5).value = 'Description';
    othersHeader.getCell(6).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: {argb: 'C0C0C0'}
    };
    othersHeader.getCell(7).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: {argb: 'C0C0C0'}
    };
    this.addBorderRow();
    this.addNewRowOthers(formArray);

    this.wts.addRow([]);
    for (var x = 6; x <= 13; x++){
      this.wts.lastRow.getCell(x).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: {argb: 'FFFF99'}
      };
      
    }
    
    
    var counter = 6
    this.days.forEach(day => {
      this.wts.lastRow.getCell(counter).value = this.othersColArray[day];
      this.wts.lastRow.getCell(counter).numFmt = '0.00';
      counter++;
    });
    this.wts.lastRow.getCell(counter).value = this.othersTotalRowCol;
    this.wts.lastRow.getCell(counter).numFmt = '0.00';

    this.addBorderRow();

  }

  addNewRowOthers(formArray: FormArray){
    var counter = 0;
    formArray.controls.forEach(element => {
      var row = this.wts.addRow([]);
      row.font = { name: 'Arial', size: 10};
      row.alignment = {vertical: 'middle'};
      row.getCell(3).value = element.get('type').value;
      row.getCell(4).value = element.get('activity').value;
      row.getCell(5).value = element.get('description').value;
      row.getCell(6).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: {argb: 'C0C0C0'}
      };
      row.getCell(7).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: {argb: 'C0C0C0'}
      };

      var dayCounter = 6;
      this.days.forEach(day => {
        row.getCell(dayCounter).value = element.get(day).value == 0 ? null : element.get(day).value;
        row.getCell(dayCounter).numFmt = '0.00'
        dayCounter++;
      });
      row.getCell(dayCounter).value = this.othersRowArray[counter];
      row.getCell(dayCounter).numFmt = '0.00';
      
      
      counter++;
      
      this.addBorderRow();
      
    });
  }

  addNewRowTask(formArray: FormArray){
    var counter = 0;
    formArray.controls.forEach(element => {

      var row = this.wts.addRow([]);
      row.font = { name: 'Arial', size: 10};
      row.alignment = {vertical: 'middle'};
      row.getCell(1).value = element.get('cName').value;
      row.getCell(3).value = element.get('pName').value;
      row.getCell(4).value = element.get('task').value;
      row.getCell(5).value = element.get('activity').value;
      row.getCell(6).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: {argb: 'C0C0C0'}
      };
      row.getCell(7).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: {argb: 'C0C0C0'}
      };

      var dayCounter = 6;
      this.days.forEach(day => {
        row.getCell(dayCounter).value = element.get(day).value == 0 ? null : element.get(day).value;
        row.getCell(dayCounter).numFmt = '0.00'
        dayCounter++;
      });
      row.getCell(dayCounter).value = this.rowArray[counter];
      row.getCell(dayCounter).numFmt = '0.00';
      counter++;
      
      this.addBorderRow();
      
    });

  }

  addBorderRow(){
    var lastRow = this.wts.rowCount;
    this.wts.mergeCells('A' + lastRow + ':B' + lastRow);
    for (var col = 1; col <= 13; col++){
      this.wts.lastRow.getCell(col).border = {
        top: {style: 'thin'},
        left: {style: 'thin'},
        bottom: {style: 'thin'},
        right: {style: 'thin'}
      }
    }
  }

  generateRows(numRows: number){
    for (var x = 0; x < numRows; x++){
      this.wts.addRow([]);
    }
  }

  

  mergeCells(){

    this.wts.mergeCells('A2:M2');
    this.wts.mergeCells('A3:M3');
    this.wts.mergeCells('A5:B5');
    this.wts.mergeCells('A6:B6');
    this.wts.mergeCells('C6:E6');
    this.wts.mergeCells('G6:H6');
    this.wts.mergeCells('A7:B7');
    this.wts.mergeCells('C7:E8');
    this.wts.mergeCells('G7:H7');
    this.wts.mergeCells('A10:B10');
    this.wts.mergeCells('C10:E10');
    this.wts.mergeCells('A14:C14');
    this.wts.mergeCells('A15:B15');

    
  }

  generateExcel(dateTo,empInitials){
    this.workbook.xlsx.writeBuffer().then((data) => {
      let blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      saveAs(blob, this.datePipe.transform(dateTo, 'yyyyMMdd') + empInitials.toUpperCase() + '.xlsx');
    });

    this.workbook = new Workbook();
    this.wts = this.workbook.addWorksheet('WeeklyTimeSheet');

  }

  blobExcel(dateTo,empInitials):any{
    this.workbook.xlsx.writeBuffer().then((data) => {  
    var file = this.email.attachFile(this.arrayBufferToBase64(data), this.datePipe.transform(dateTo, 'yyyyMMdd') + empInitials.toUpperCase());
      
    });
    this.workbook = new Workbook();
    this.wts = this.workbook.addWorksheet('WeeklyTimeSheet');
  }

  arrayBufferToBase64( buffer ) {
    var binary = '';
    var bytes = new Uint8Array( buffer );
    var len = bytes.byteLength;
    for (var i = 0; i < len; i++) {
        binary += String.fromCharCode( bytes[ i ] );
    }
    return window.btoa( binary );
}



  setColumnWidth(){
    this.wts.getColumn(1).width = 5.71 + 0.71;
    this.wts.getColumn(2).width = 12.71 + 0.71;
    this.wts.getColumn(3).width = 18.71;
    this.wts.getColumn(4).width = 21.71;    
    this.wts.getColumn(5).width = 56.14 + 0.71;
    this.wts.getColumn(6).width = 10 + 0.71;
    this.wts.getColumn(7).width = 10 + 0.71;
    this.wts.getColumn(8).width = 10 + 0.71;
    this.wts.getColumn(9).width = 10 + 0.71;
    this.wts.getColumn(10).width = 10 + 0.71;
    this.wts.getColumn(11).width = 10 + 0.71;
    this.wts.getColumn(12).width = 10 + 0.71;
    this.wts.getColumn(13).width = 10 + 0.71;
  }

  setRowHeight(){
    for (var x = 1; x <= this.wts.rowCount; x++ ){
      this.wts.getRow(x).height = 12.75;

    }
    this.wts.getRow(2).height = 18;
  }

}
