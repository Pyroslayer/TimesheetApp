import { Component, OnInit } from '@angular/core';
import { FormBuilder, Validators, FormGroup, FormArray } from '@angular/forms';
import {Constants } from '../../assets/constants';
import { Time, DatePipe } from '@angular/common';
import { ExcelService } from '../excel.service';
import { FormService } from '../form.service';


@Component({
  selector: 'app-home',
  templateUrl: './home.component.html',
  styleUrls: ['./home.component.css']
})
export class HomeComponent implements OnInit {

  group: FormGroup;

  array: FormArray;
  colArray: number[]=[];
  rowArray: number[]=[0];
  othersColArray: number[]=[];
  othersRowArray: number[]=[0];
  cName;
  pName;
  tasks;
  empGroup;
  othersType;
  othersActivity;
  finalRow: number[] = [0];
  totalClass = [];

  
  
  days = ['saturday','sunday','monday','tuesday','wednesday','thursday','friday'];
  totalRowCol = 0;
  othersTotalRowCol = 0;
  finalTotal = 0;
  

  get timeSheet(): FormArray { return this.group.get('timeSheet') as FormArray; }
  get timeInOut(): FormArray { return this.group.get('timeInOut') as FormArray; }
  get others(): FormArray { return this.group.get('others') as FormArray; }

  constructor(  private formBuilder: FormBuilder, 
                private datePipe: DatePipe, 
                private excel: ExcelService,
                private form: FormService ) {
    this.group = this.formBuilder.group({
      empNo: ['', Validators.required],
      empName: ['', Validators.required],
      empSupervisor: ['', Validators.required],
      empInitials: ['', Validators.required],
      group: ['', Validators.required],
      dateFrom: ['', Validators.required],
      dateTo: ['', Validators.required],
      timeSheet: this.formBuilder.array([]),
      timeInOut: this.formBuilder.array([]),
      others: this.formBuilder.array([]),

    });

    
    // this.form.syncWithLocalStorage(this.group.get('timeSheet').value, 'timeSheet');
    // this.form.syncWithLocalStorage(this.group.get('timeInOut').value, 'timeInOut');
    // this.form.syncWithLocalStorage(this.group.get('others').value, 'others');

    
  }

  ngOnInit() {
    
    this.cName = Constants.CLIENT_NAME;
    this.pName = Constants.PROJECT_NAME;
    this.tasks = Constants.TASK;
    this.empGroup = Constants.GROUP;
    this.othersType = Constants.TYPE;
    this.othersActivity = Constants.ACTIVITIES;
    
    this.fillTimeInOut();
    this.mapRows();

    this.getTotal();
    
    
  }

  mapRows(){
    if(localStorage.getItem('taskCount')!=null){
      for(var x= 0; x < +localStorage.getItem('taskCount');x++){
        this.newRow();
      }
    }
    else {
      this.newRow();
    }

    if(localStorage.getItem('othersCount')!=null){
      for(var x= 0; x < +localStorage.getItem('othersCount');x++){
        this.newOthersRow();
      }
    }
    else {
      this.newOthersRow();
    }

    
    this.form.getFromLocalStorage(this.group,'group');
  }

  setDate(){
    var dateFrom = new Date(this.group.get('dateTo').value );
    if(dateFrom.getDay()==5){
      dateFrom.setDate(dateFrom.getDate() - 6);
      this.group.get('dateFrom').setValue(this.datePipe.transform(dateFrom,'yyyy-MM-dd'));
    }else {
      alert('Date is not a Friday!');
      this.group.get('dateTo').setValue('');
      this.group.get('dateFrom').setValue('');
    }
    
  }
    
  


  fillTimeInOut(){
    
    this.days.forEach(day => {
      if(day == 'saturday' || day == 'sunday'){
        this.timeInOut.push(
          this.formBuilder.group({
            timeIn: null,
            timeOut: null,
            hours: 0
          })
        );
      } else {
        this.timeInOut.push(
          this.formBuilder.group({
            timeIn: '08:30',
            timeOut: '17:30',
            hours: 9
          })
        );
      }
    });

    this.getHoursPerDay();
  }



  newRow(){
    this.timeSheet.push(
      this.formBuilder.group({
        cName: ['', Validators.required],
        pName: ['', Validators.required],
        task: ['', Validators.required],
        activity: ['', Validators.required],
        saturday: 0,
        sunday: 0,
        monday: 0,
        tuesday: 0,
        wednesday: 0,
        thursday: 0,
        friday: 0
      })
    );
    
    this.getTotal();
  }

  newOthersRow(){
    this.others.push(
      this.formBuilder.group({
        type: ['', Validators.required],
        activity: ['', Validators.required],
        description: ['', Validators.required],
        saturday: 0,
        sunday: 0,
        monday: 0,
        tuesday: 0,
        wednesday: 0,
        thursday: 0,
        friday: 0
      })
    );
    
    this.getTotal();
    
  }

  deleteRow(rowValue){
    if(confirm('Are you sure you want to delete TASK row '+ (rowValue+1))){
      this.timeSheet.controls.splice(rowValue,1);
    }
    
  }

  validate(){
    this.group.markAllAsTouched();
    var counter = 0;
    this.days.forEach(day => {
      console.log(this.timeInOut.controls[counter].get('hours').value);
      
      if(this.finalRow[day] != this.timeInOut.controls[counter].get('hours').value){
        this.timeInOut.controls[counter].get('hours').setErrors({'incorrect':true});
      } else {
        this.timeInOut.controls[counter].get('hours').setErrors(null);
      }
      counter++;
    });

  }

  deleteOthersRow(rowValue){
    if(confirm('Are you sure you want to delete OTHER row '+ (rowValue+1))){
      this.others.controls.splice(rowValue,1);
    }
  }

  getHoursPerDay(){
    var counter = 0;
    this.days.forEach(element => {
      var rawTimeIn = this.timeInOut.controls[counter].get('timeIn').value;
      var rawTimeOut = this.timeInOut.controls[counter].get('timeOut').value;
      var hours = this.timeInOut.controls[counter].get('hours');
     
      var timeIn = new Date( '01/01/2001 ' + (rawTimeIn == null ? '00:00:00' : rawTimeIn));
      var timeOut = new Date( '01/01/2001 ' + (rawTimeOut == null ? '00:00:00' : rawTimeOut));

      var hoursPerDay = timeOut.getHours() - timeIn.getHours();
      hoursPerDay += (timeOut.getMinutes() - timeIn.getMinutes()) / 60;

      if(hoursPerDay<0 || rawTimeOut === '' || rawTimeIn === ''){
        hoursPerDay = 0;
      }

      hours.setValue(hoursPerDay);
      
      counter++;
    });

  }

  validateInput(i,day,type){

    var input;
    if(type == 'timeSheet'){
      input = this.timeSheet.controls[i].get(day);
    } 
    else if(type == 'others'){
      input = this.others.controls[i].get(day);
    }
    

    if(input.value<0){
      input.setValue(0);
      
    }
    this.getTotal();
  }

  


  getTotal(){
    this.totalRowCol = 0;
    this.othersTotalRowCol = 0;
    this.finalTotal = 0;

    console.log(this.group);
    
    //initialize total of days to zero
    this.days.forEach(day => {
      this.colArray[day] = 0;
      this.othersColArray[day] = 0;
      
    });

    
    let counter = 0;
    this.timeSheet.controls.forEach(element => {
      this.rowArray[counter] = 0;
      this.days.forEach(day => {

        this.rowArray[counter] += element.get(day).value;
        this.colArray[day] += element.get(day).value;
      });

      counter++;
    });

    counter = 0;
    this.others.controls.forEach(element => {
      this.othersRowArray[counter] = 0;
      this.days.forEach(day => {

        this.othersRowArray[counter] += element.get(day).value;
        this.othersColArray[day] += element.get(day).value;
      });

      counter++;
    });

    this.days.forEach(day => {
      this.finalRow[day] = this.othersColArray[day] + this.colArray[day];
      this.finalTotal += this.finalRow[day];
      
    });

    //add total based on row
    this.rowArray.forEach(element => {
      this.totalRowCol += element;
    });

    this.othersRowArray.forEach(element => {
      this.othersTotalRowCol += element;
      
      
    });
  }

   setToLocal(){
    var taskCount = 0;
    var othersCount = 0;
    this.timeSheet.controls.forEach(element => {
      taskCount++
    });
    this.others.controls.forEach(element => {
      othersCount++
    });
    localStorage.setItem('taskCount',taskCount.toString());
    localStorage.setItem('othersCount',othersCount.toString());

    this.form.setToLocalStorage(this.group,'group');
    alert('Saved!');
  
  }

  generate(){
    this.validate();

    if(this.group.valid){
      let list = [];
      list['rowArray']= this.rowArray;
      list['colArray']= this.colArray;
      list['othersColArray'] = this.othersColArray;
      list['othersRowArray'] = this.othersRowArray;
      list['finalRow'] = this.finalRow;
      list['totalRowCol'] = this.totalRowCol;
      list['othersTotalRowCol'] = this.othersTotalRowCol;
      list['finalTotal'] = this.finalTotal;
    
    this.excel.generate(this.group, list);
    }

  }


}
