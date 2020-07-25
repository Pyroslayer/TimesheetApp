import { Component, OnInit } from '@angular/core';
import { FormBuilder, Validators, FormGroup, FormArray, FormControl } from '@angular/forms';
import {Constants } from '../../assets/constants';
import { Time, DatePipe } from '@angular/common';
import { ExcelService } from '../excel.service';
import { FormService } from '../form.service';
import { EmailServiceService } from '../email-service.service';

@Component({
  selector: 'app-home',
  templateUrl: './home.component.html',
  styleUrls: ['./home.component.css']
})
export class HomeComponent implements OnInit {

  group: FormGroup;
  emailGroup: FormGroup;
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
  headerToggle = true;
  tableToggle = true;
  wfhDay;
  dateArray = [];
  wfhTasks = [];
  wfhQueue = [];
  showEditWFH = true;
  emailType = '';
  filename = '';
  days = ['saturday','sunday','monday','tuesday','wednesday','thursday','friday'];
  totalRowCol = 0;
  othersTotalRowCol = 0;
  finalTotal = 0;
  showWFH = false;
  showEmail = false;
  wfhArray = []

  get timeSheet(): FormArray { return this.group.get('timeSheet') as FormArray; }
  get timeInOut(): FormArray { return this.group.get('timeInOut') as FormArray; }
  get others(): FormArray { return this.group.get('others') as FormArray; }
  // get wfhArray(): FormArray { return this.group.get('wfhArray') as FormArray; }
  get toEmail():FormArray { return this.emailGroup.get('toEmail') as FormArray}
  get ccEmail():FormArray { return this.emailGroup.get('ccEmail') as FormArray}
  get bccEmail():FormArray { return this.emailGroup.get('bccEmail') as FormArray}
  constructor(  private formBuilder: FormBuilder, 
                private datePipe: DatePipe, 
                private excel: ExcelService,
                private form: FormService,
                private email: EmailServiceService,
                ) {
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
      wfh: [0],
      // wfhArray: this.formBuilder.array([]),
    });
    this.emailGroup = this.formBuilder.group({
      userEmail: ['',Validators.required],
      userPassword: ['',Validators.required],
      toEmail: this.formBuilder.array([]),
      ccEmail: this.formBuilder.array([]),
      bccEmail: this.formBuilder.array([]),
      subject: ['',Validators.required],
      message: ['',Validators.required],
    });
    // if(!localStorage.getItem('group')){
    //   this.form.setToLocalStorage(this.group, 'group');
    // }
    // if(!localStorage.getItem('timeSheet')){
    //   this.form.setToLocalStorage(this.group.get('timeSheet') as FormGroup, 'timeSheet');
    // }
    // if(!localStorage.getItem('timeInOut')){
    //   this.form.setToLocalStorage(this.group.get('timeInOut') as FormGroup, 'timeInOut');
    // }
    // if(!localStorage.getItem('others')){
    //   this.form.setToLocalStorage(this.group.get('others') as FormGroup, 'others');
    // }

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
    this.checkErrors();
    
    // if(localStorage.getItem('group')){
    //   this.setDate();
    // }

    // this.getTasks();
    this.group.get('wfh').valueChanges.subscribe(data=>{
      this.getTasks();
    })

    console.log(this.group);
    
  }

  timeIn(){
    var date = new Date().getDay() + 1;
    if(date === 7) { date = 0; }
    this.timeInOut.controls[date].get('timeIn').setValue((new Date().getHours()).toString() + ':'+ (this.minTwoDigits(new Date().getMinutes())).toString() );
    
  }

  timeOut(){
    var date = new Date().getDay() + 1;
    if(date === 7) { date = 0; }
    this.timeInOut.controls[date].get('timeOut').setValue((new Date().getHours()).toString() + ':'+ (this.minTwoDigits(new Date().getMinutes())).toString() );

  }

  checkErrors() {
    if ( this.group.get('empNo').valid && 
        this.group.get('empName').valid &&
        this.group.get('empSupervisor').valid && 
        this.group.get('empInitials').valid && 
        this.group.get('group').valid){
          this.headerToggle = false;
    }
    else {
      this.headerToggle = true;
    }
    

  }

  tableButtonToggle(){
    this.tableToggle = !this.tableToggle;
  }

  headerButtonToggle(){
    this.headerToggle = !this.headerToggle;
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

    if(localStorage.getItem('toEmailCount')!=null){
      for(var x= 0; x < +localStorage.getItem('toEmailCount');x++){
        this.newTo();
      }
    }
    else {
      this.newTo();
    }

    if(localStorage.getItem('ccEmailCount')!=null){
      for(var x= 0; x < +localStorage.getItem('ccEmailCount');x++){
        this.newCc();
      }
    }
    else {
      this.newCc();
    }

    if(localStorage.getItem('bccEmailCount')!=null){
      for(var x= 0; x < +localStorage.getItem('bccEmailCount');x++){
        this.newBcc();
      }
    }
    else {
      this.newBcc();
    }

    if(localStorage.getItem('othersCount')!=null){
      for(var x= 0; x < +localStorage.getItem('othersCount');x++){
        this.newOthersRow();
      }
    }
    else {
      this.newOthersRow();
    }

    if(localStorage.getItem('group')){
      this.form.getFromLocalStorage(this.group,'group');
    }
    if(localStorage.getItem('emailGroup')){
      this.form.getFromLocalStorage(this.emailGroup,'emailGroup');
    }
    

    this.group.valueChanges.subscribe(data => {
      this.setToLocal();
    });
    this.emailGroup.valueChanges.subscribe(data => {
      this.setToLocal();
    });
    // this.form.syncWithLocalStorage(this.group.get('timeSheet') as FormGroup, 'timeSheet');
    // this.group.get('timeInOut').valueChanges.subscribe(data => {
    //   this.form.setToLocalStorage(this.group.get('timeInOut') as FormGroup, 'timeInOut');
    // })
    // this.form.syncWithLocalStorage(this.group.get('timeInOut') as FormGroup, 'timeInOut');
    // this.form.syncWithLocalStorage(this.group.get('others') as FormGroup, 'others');
    // this.form.syncWithLocalStorage(this.group, 'group');
  }

  setDate(){
    var dateFrom = new Date(this.group.get('dateTo').value );
    if(dateFrom.getDay()==5){
      dateFrom.setDate(dateFrom.getDate() - 6);
      this.group.get('dateFrom').setValue(this.datePipe.transform(dateFrom,'yyyy-MM-dd'));
    }else if(this.group.get('dateTo').value !== '') {
      alert('Date is not a Friday!');
      this.group.get('dateTo').setValue('');
      this.group.get('dateFrom').setValue('');
    }

    for(var x=6;x>=0;x--){
      var tempDate = new Date(this.group.get('dateTo').value );
      tempDate.setDate(tempDate.getDate() - x);
      this.dateArray[6-x] = this.datePipe.transform(tempDate,'MM/dd/yyyy');
    }

    
    
  }
    
  getTasks(){
    this.wfhTasks = [];
    this.wfhQueue = [];
    this.wfhArray = [];
    Object.keys((this.group.controls['timeSheet'] as FormGroup).controls).forEach(key => {
      if(((this.group.controls['timeSheet'] as FormGroup).controls[key].get(this.days[this.group.get('wfh').value]).value) > 0){
        this.wfhTasks.push((this.group.controls['timeSheet'] as FormGroup).controls[key].get('activity').value);
        this.wfhArray.push({
          status: "",
          remarks: ""
        });
  
      } else {
        var checker = false;
        for(var x = 0; x < this.group.get('wfh').value; x++){
          if(((this.group.controls['timeSheet'] as FormGroup).controls[key].get(this.days[x]).value) > 0) {
            checker = true;
          }
        }
        if(!checker){
          this.wfhQueue.push((this.group.controls['timeSheet'] as FormGroup).controls[key].get('activity').value);
        }
      }
    });
    
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


  newTo(){
    this.toEmail.push(
      this.formBuilder.group({
      email:['',Validators.required]
      })
    )
  }

  deleteTo(index){
    this.toEmail.removeAt(index);
  }

  deleteCc(index){
    this.ccEmail.removeAt(index);
  }

  deleteBcc(index){
    this.bccEmail.removeAt(index);
  }

  newCc(){
    this.ccEmail.push(
      this.formBuilder.group({
        email:['']
        })
    )
  }

  newBcc(){
    this.bccEmail.push(
      this.formBuilder.group({
        email:['']
        })
    )
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
      this.timeSheet.removeAt(rowValue);
      this.getTotal();
    }
    
    
  }

  validate(){
    this.group.markAllAsTouched();
    var counter = 0;
    this.days.forEach(day => {
      
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
      this.others.removeAt(rowValue);
      this.getTotal();
      
      
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

      hours.setValue(Math.round(hoursPerDay*100)/100);
      
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

    input.setValue(Math.round((input.value + Number.EPSILON) * 100) / 100)
    this.getTotal();
  }

  


  getTotal(){
    this.totalRowCol = 0;
    this.othersTotalRowCol = 0;
    this.finalTotal = 0;
    this.rowArray = [];
    this.othersRowArray = [];

    
    //initialize total of days to zero
    this.days.forEach(day => {
      this.colArray[day] = 0;
      this.othersColArray[day] = 0;
      
    });

    
    let counter = 0;
    this.timeSheet.controls.forEach(element => {
      this.rowArray[counter] = 0;
      this.days.forEach(day => {
        element.get(day).setValue(element.get(day).value);
        this.rowArray[counter] += element.get(day).value ;
        this.colArray[day] += element.get(day).value;

      });
      counter++;
    });

    counter = 0;
    this.timeSheet.controls.forEach(element => {
      this.days.forEach(day => {
        this.rowArray[counter] = Math.round((this.rowArray[counter] + Number.EPSILON) * 100) / 100;
        this.colArray[day] = Math.round((this.colArray[day] + Number.EPSILON) * 100) / 100;
      });
      counter++;
    });

    counter = 0;
    this.others.controls.forEach(element => {
      this.othersRowArray[counter] = 0;
      this.days.forEach(day => {
        element.get(day).setValue(element.get(day).value);
        this.othersRowArray[counter] += element.get(day).value;
        this.othersColArray[day] += element.get(day).value;
      });

      counter++;
    });

    counter = 0;
    this.others.controls.forEach(element => {
      this.days.forEach(day => {
        this.othersRowArray[counter] = Math.round((this.othersRowArray[counter] + Number.EPSILON) * 100) / 100;
        this.othersColArray[day] = Math.round((this.othersColArray[day] + Number.EPSILON) * 100) / 100;
      });

      counter++;
    });

    this.days.forEach(day => {
      this.finalRow[day] = this.othersColArray[day] + this.colArray[day];
      this.finalRow[day] = Math.round((this.finalRow[day] + Number.EPSILON) * 100) / 100
      this.finalTotal += this.finalRow[day];
    });
    this.finalTotal = Math.round((this.finalTotal + Number.EPSILON) * 100) / 100

    //add total based on row
    this.rowArray.forEach(element => {
      this.totalRowCol += element;
    });
    this.totalRowCol = Math.round((this.totalRowCol + Number.EPSILON) * 100) / 100

    this.othersRowArray.forEach(element => {
      this.othersTotalRowCol += element;
    });
    this.othersTotalRowCol = Math.round((this.othersTotalRowCol + Number.EPSILON) * 100) / 100


  }

   setToLocal(){
    var taskCount = 0;
    var othersCount = 0;
    var toEmailCount = 0;
    var ccEmailCount = 0;
    var bccEmailCount = 0;
    this.timeSheet.controls.forEach(element => {
      taskCount++
    });
    this.others.controls.forEach(element => {
      othersCount++
    });
    this.toEmail.controls.forEach(element => {
      toEmailCount++
    });
    this.ccEmail.controls.forEach(element => {
      ccEmailCount++
    });
    this.bccEmail.controls.forEach(element => {
      bccEmailCount++
    });
    localStorage.setItem('taskCount',taskCount.toString());
    localStorage.setItem('othersCount',othersCount.toString());
    localStorage.setItem('toEmailCount',toEmailCount.toString());
    localStorage.setItem('ccEmailCount',ccEmailCount.toString());
    localStorage.setItem('bccEmailCount',bccEmailCount.toString());

    this.form.setToLocalStorage(this.group,'group');
    this.form.setToLocalStorage(this.emailGroup,'emailGroup');
    // alert('Saved!');
  
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


  setToTimeSheet(){
    this.validate();

    if(this.group.valid){
      let list = [];
      this.emailType="timesheet";
      this.filename = this.datePipe.transform(this.group.get('dateTo').value, 'yyyyMMdd') + this.group.get('empInitials').value.toUpperCase()
      this.emailGroup.get('message').setValue('Good day sir, I have attached in this email my time sheet for this week. Thank you.');
      this.emailGroup.get('subject').setValue('TIMESHEET '+this.datePipe.transform(this.group.get('dateTo').value, 'yyyyMMdd') + this.group.get('empInitials').value.toUpperCase()),
      this.emailGroup.get('subject').disable();
      this.ccEmail.controls[0].get('email').setValue('');
      this.ccEmail.enable();

      list['rowArray']= this.rowArray;
      list['colArray']= this.colArray;
      list['othersColArray'] = this.othersColArray;
      list['othersRowArray'] = this.othersRowArray;
      list['finalRow'] = this.finalRow;
      list['totalRowCol'] = this.totalRowCol;
      list['othersTotalRowCol'] = this.othersTotalRowCol;
      list['finalTotal'] = this.finalTotal;
    
    this.excel.generateBlob(this.group, list);  
    } else {
      alert('Please check all the fields in your timesheet!');
    }
  }

  toggleWFH(){
    this.getTasks();
    console.log(this.wfhArray);
    
    this.showWFH = !this.showWFH;
    if(this.showWFH){
      this.showEditWFH = true;
      
      for(var x=6;x>=0;x--){
        var tempDate = new Date(this.group.get('dateTo').value );
        tempDate.setDate(tempDate.getDate() - x);
        this.dateArray[6-x] = this.datePipe.transform(tempDate,'MM/dd/yyyy');
      }
      
    }
  }

  toggleEditWFH(){
    this.showEditWFH = !this.showEditWFH;
  }

  toggleEmail(){
    this.showEmail = !this.showEmail;
    if(!this.showEmail){
      this.emailGroup.markAsUntouched();
      this.emailType = "";
      this.emailGroup.get('subject').enable();
      this.emailGroup.get('subject').setValue('');
      this.ccEmail.controls[0].get('email').enable();
      this.ccEmail.controls[0].get('email').setValue('');
      this.email.clearFile();
      this.emailGroup.get('message').setValue('');    
    
    }
  }

  setToWfh() {
    this.getTasks();
    this.emailType = "WFH";
    this.email.clearFile();
    var date = new Date().getDay() + 1;
    if(date === 7){
      date = 0;
    }
    this.group.get('wfh').setValue(date);
    
    
    for(var x=6;x>=0;x--){
      var tempDate = new Date(this.group.get('dateTo').value );
      tempDate.setDate(tempDate.getDate() - x);
      this.dateArray[6-x] = this.datePipe.transform(tempDate,'MM/dd/yyyy');
    }

    this.emailGroup.get('subject').setValue('WFH - TA');
    this.emailGroup.get('subject').disable();
    this.ccEmail.controls[0].get('email').setValue('tera.leaves@gmail.com');
    this.ccEmail.controls[0].get('email').disable();
    this.emailGroup.get('message').setValue('test')
    
  }


  sendEmail(){
    this.emailGroup.markAllAsTouched();
    if(!this.emailGroup.valid){
      return;
    }
    if(!confirm("Are you sure you want to submit this email?")){
      return;
    }
    var message = '';
      if(this.emailType!== 'WFH'){
        message = this.emailGroup.get('message').value
      }else if(this.emailType = "WFH"){
        message = this.message();
      }
      var data = {
        message: message,
        userEmail: this.emailGroup.get('userEmail').value,
        userName: this.group.get('empName').value,
        userPassword: this.emailGroup.get('userPassword').value,
        toEmail: this.toEmail.getRawValue(),
        ccEmail: this.ccEmail.getRawValue(),
        bccEmail: this.bccEmail.getRawValue(),
        subject: this.emailGroup.get('subject').value,
      }
      this.email.sendEmail(data).subscribe((data:any)=>{
        alert(data.message);
        
      },(err)=>{
        alert(err.error.message);
      }
      
      ) 

  }

  convertTo12(timeString){
    var dt = new Date('1970-01-01T' + timeString);
    var hours = dt.getHours() ; // gives the value in 24 hours format
    var AmOrPm = hours >= 12 ? 'PM' : 'AM';
    hours = (hours % 12) || 12;
    var minutes = dt.getMinutes() ;
    var finalTime = hours + ":" + this.minTwoDigits(minutes) + " " + AmOrPm; 
    return finalTime;
  }

  minTwoDigits(n) {
    return (n < 10 ? '0' : '') + n;
  }
  capitalize(word){
    return word.charAt(0).toUpperCase() + word.substring(1);
  }

  message(){
    var message = `<p class="MsoNormal" style="margin:0cm 0cm 8pt;line-height:15.6933px;font-size:11pt;font-family:Calibri,sans-serif">Hi Sir!</p> <p class="MsoNormal" style="margin:0cm 0cm 8pt;line-height:15.6933px;font-size:11pt;font-family:Calibri,sans-serif">Please see work from home details for ${this.dateArray[this.group.get('wfh').value]} (${this.capitalize(this.days[this.group.get('wfh').value])})</p> <p class="MsoNormal" style="margin:0cm 0cm 8pt;line-height:15.6933px;font-size:11pt;font-family:Calibri,sans-serif">Attendance Detail:</p>`;
    message+= `<table border="1" cellspacing="0" cellpadding="0" align="left" style="border-collapse:collapse;border:none;margin-left:6.75pt;margin-right:6.75pt"><tbody><tr style="height:17.1pt"><td width="103" valign="top" style="width:77.3pt;border:1pt solid windowtext;padding:0cm 5.4pt;height:17.1pt"><p class="MsoNormal" align="center" style="margin:0cm 0cm 0.0001pt;text-align:center;line-height:normal;font-size:11pt;font-family:Calibri,sans-serif"><b>LOG IN</b></p></td><td width="103" valign="top" style="width:77.3pt;border-top:1pt solid windowtext;border-right:1pt solid windowtext;border-bottom:1pt solid windowtext;border-left:none;padding:0cm 5.4pt;height:17.1pt"><p class="MsoNormal" align="center" style="margin:0cm 0cm 0.0001pt;text-align:center;line-height:normal;font-size:11pt;font-family:Calibri,sans-serif"><b>LOG OUT</b></p></td></tr><tr style="height:17.1pt"><td width="103" valign="top" style="width:77.3pt;border-right:1pt solid windowtext;border-bottom:1pt solid windowtext;border-left:1pt solid windowtext;border-top:none;padding:0cm 5.4pt;height:17.1pt"><p class="MsoNormal" style="margin:0cm 0cm 0.0001pt;line-height:normal;font-size:11pt;font-family:Calibri,sans-serif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ${ this.convertTo12(this.group.get('timeInOut')['controls'][this.group.get('wfh').value].get('timeIn').value)}</p></td><td width="103" valign="top" style="width:77.3pt;border-top:none;border-left:none;border-bottom:1pt solid windowtext;border-right:1pt solid windowtext;padding:0cm 5.4pt;height:17.1pt"><p class="MsoNormal" style="margin:0cm 0cm 0.0001pt;line-height:normal;font-size:11pt;font-family:Calibri,sans-serif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ${ this.convertTo12(this.group.get('timeInOut')['controls'][this.group.get('wfh').value].get('timeOut').value)}</p></td></tr></tbody></table><p class="MsoNormal" style="margin:0cm 0cm 8pt;line-height:15.6933px;font-size:11pt;font-family:Calibri,sans-serif">&nbsp;</p><p class="MsoNormal" style="margin:0cm 0cm 8pt;line-height:15.6933px;font-size:11pt;font-family:Calibri,sans-serif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</p><p class="MsoNormal" style="margin:0cm 0cm 8pt;line-height:15.6933px;font-size:11pt;font-family:Calibri,sans-serif">&nbsp;</p><p class="MsoNormal" style="margin:0cm 0cm 8pt;line-height:15.6933px;font-size:11pt;font-family:Calibri,sans-serif">Task Accomplishment Summary for the day</p>`;
    message += `<table border="0" cellspacing="0" cellpadding="0" width="601" style="width:450.8pt;border-collapse:collapse"> <tbody> <tr> <td width="200" valign="top" style="width:150.25pt;border:1pt solid windowtext;padding:0cm 5.4pt"> <p class="MsoNormal" align="center" style="margin:0cm 0cm 0.0001pt;text-align:center;line-height:normal;font-size:11pt;font-family:Calibri,sans-serif"><b>TASK</b></p> </td> <td width="200" valign="top" style="width:150.25pt;border-top:1pt solid windowtext;border-right:1pt solid windowtext;border-bottom:1pt solid windowtext;border-left:none;padding:0cm 5.4pt"> <p class="MsoNormal" align="center" style="margin:0cm 0cm 0.0001pt;text-align:center;line-height:normal;font-size:11pt;font-family:Calibri,sans-serif"><b>STATUS</b></p> </td> <td width="200" valign="top" style="width:150.3pt;border-top:1pt solid windowtext;border-right:1pt solid windowtext;border-bottom:1pt solid windowtext;border-left:none;padding:0cm 5.4pt"> <p class="MsoNormal" align="center" style="margin:0cm 0cm 0.0001pt;text-align:center;line-height:normal;font-size:11pt;font-family:Calibri,sans-serif"><b>REMARKS</b></p> </td> </tr>`;
    if(this.wfhTasks.length === 0){
      message += `<tr> <td width="200" valign="top" style="width:150.25pt;border-right:1pt solid windowtext;border-bottom:1pt solid windowtext;border-left:1pt solid windowtext;border-top:none;padding:0cm 5.4pt"> <p class="MsoNormal" style="margin:0cm 0cm 0.0001pt;line-height:normal;font-size:11pt;font-family:Calibri,sans-serif">No task for the day</p> </td> <td width="200" valign="top" style="width:150.25pt;border-top:none;border-left:none;border-bottom:1pt solid windowtext;border-right:1pt solid windowtext;padding:0cm 5.4pt"></td> <td width="200" valign="top" style="width:150.3pt;border-top:none;border-left:none;border-bottom:1pt solid windowtext;border-right:1pt solid windowtext;padding:0cm 5.4pt"></td> </tr>`;
    }
    this.wfhArray.forEach((task,i) => {
      message+= `<tr> <td width="200" valign="top" style="width:150.25pt;border-right:1pt solid windowtext;border-bottom:1pt solid windowtext;border-left:1pt solid windowtext;border-top:none;padding:0cm 5.4pt"> <p class="MsoNormal" style="margin:0cm 0cm 0.0001pt;line-height:normal;font-size:11pt;font-family:Calibri,sans-serif">${this.wfhTasks[i]}</p> </td> <td width="200" valign="top" style="width:150.25pt;border-top:none;border-left:none;border-bottom:1pt solid windowtext;border-right:1pt solid windowtext;padding:0cm 5.4pt"> ${this.wfhArray[i]['status']} </td> <td width="200" valign="top" style="width:150.3pt;border-top:none;border-left:none;border-bottom:1pt solid windowtext;border-right:1pt solid windowtext;padding:0cm 5.4pt"> ${this.wfhArray[i]['remarks']} </td> </tr>`;
 
    });
    message+= `<tr> <td width="200" valign="top" style="width:150.25pt;border-right:1pt solid windowtext;border-bottom:1pt solid windowtext;border-left:1pt solid windowtext;border-top:none;padding:0cm 5.4pt"> <p class="MsoNormal" style="margin:0cm 0cm 0.0001pt;line-height:normal;font-size:11pt;font-family:Calibri,sans-serif">&nbsp; &nbsp; </p> </td> <td width="200" valign="top" style="width:150.25pt;border-top:none;border-left:none;border-bottom:1pt solid windowtext;border-right:1pt solid windowtext;padding:0cm 5.4pt"> <p class="MsoNormal" style="margin:0cm 0cm 0.0001pt;line-height:normal;font-size:11pt;font-family:Calibri,sans-serif">&nbsp;</p> </td> <td width="200" valign="top" style="width:150.3pt;border-top:none;border-left:none;border-bottom:1pt solid windowtext;border-right:1pt solid windowtext;padding:0cm 5.4pt"> <p class="MsoNormal" style="margin:0cm 0cm 0.0001pt;line-height:normal;font-size:11pt;font-family:Calibri,sans-serif">&nbsp;</p> </td> </tr> <tr> <td width="200" valign="top" style="width:150.25pt;border-right:1pt solid windowtext;border-bottom:1pt solid windowtext;border-left:1pt solid windowtext;border-top:none;padding:0cm 5.4pt"> <p class="MsoNormal" style="margin:0cm 0cm 0.0001pt;line-height:normal;font-size:11pt;font-family:Calibri,sans-serif">&nbsp; &nbsp; </p> </td> <td width="200" valign="top" style="width:150.25pt;border-top:none;border-left:none;border-bottom:1pt solid windowtext;border-right:1pt solid windowtext;padding:0cm 5.4pt"> <p class="MsoNormal" style="margin:0cm 0cm 0.0001pt;line-height:normal;font-size:11pt;font-family:Calibri,sans-serif">&nbsp;</p> </td> <td width="200" valign="top" style="width:150.3pt;border-top:none;border-left:none;border-bottom:1pt solid windowtext;border-right:1pt solid windowtext;padding:0cm 5.4pt"> <p class="MsoNormal" style="margin:0cm 0cm 0.0001pt;line-height:normal;font-size:11pt;font-family:Calibri,sans-serif">&nbsp;</p> </td> </tr>`;
    message += ` <tr> <td width="601" colspan="3" valign="top" style="width:450.8pt;border-right:1pt solid windowtext;border-bottom:1pt solid windowtext;border-left:1pt solid windowtext;border-top:none;padding:0cm 5.4pt"> <p class="MsoNormal" style="margin:0cm 0cm 0.0001pt;line-height:normal;font-size:11pt;font-family:Calibri,sans-serif"><b>Other tasks on Queue:</b></p> </td> </tr>`;
    if( this.wfhQueue.length === 0){
      message+=`<tr><td width="200" valign="top" style="width:150.25pt;border-right:1pt solid windowtext;border-bottom:1pt solid windowtext;border-left:1pt solid windowtext;border-top:none;padding:0cm 5.4pt"> <p class="MsoNormal" style="margin:0cm 0cm 0.0001pt;line-height:normal;font-size:11pt;font-family:Calibri,sans-serif">No pending task</p> </td> <td width="200" valign="top" style="width:150.25pt;border-top:none;border-left:none;border-bottom:1pt solid windowtext;border-right:1pt solid windowtext;padding:0cm 5.4pt"></td> <td width="200" valign="top" style="width:150.3pt;border-top:none;border-left:none;border-bottom:1pt solid windowtext;border-right:1pt solid windowtext;padding:0cm 5.4pt"></td></tr>`;
    }
    
    this.wfhQueue.forEach(task => {
      message+= `<tr> <td width="200" valign="top" style="width:150.25pt;border-right:1pt solid windowtext;border-bottom:1pt solid windowtext;border-left:1pt solid windowtext;border-top:none;padding:0cm 5.4pt"> <p class="MsoNormal" style="margin:0cm 0cm 0.0001pt;line-height:normal;font-size:11pt;font-family:Calibri,sans-serif">${task}</p> </td> <td width="200" valign="top" style="width:150.25pt;border-top:none;border-left:none;border-bottom:1pt solid windowtext;border-right:1pt solid windowtext;padding:0cm 5.4pt"> </td> <td width="200" valign="top" style="width:150.3pt;border-top:none;border-left:none;border-bottom:1pt solid windowtext;border-right:1pt solid windowtext;padding:0cm 5.4pt"></td> </tr>`;
    });
    message+= ` <tr> <td width="200" valign="top" style="width:150.25pt;border-right:1pt solid windowtext;border-bottom:1pt solid windowtext;border-left:1pt solid windowtext;border-top:none;padding:0cm 5.4pt"> <p class="MsoNormal" style="margin:0cm 0cm 0.0001pt;line-height:normal;font-size:11pt;font-family:Calibri,sans-serif">&nbsp; &nbsp; </p> </td> <td width="200" valign="top" style="width:150.25pt;border-top:none;border-left:none;border-bottom:1pt solid windowtext;border-right:1pt solid windowtext;padding:0cm 5.4pt"> <p class="MsoNormal" style="margin:0cm 0cm 0.0001pt;line-height:normal;font-size:11pt;font-family:Calibri,sans-serif">&nbsp;</p> </td> <td width="200" valign="top" style="width:150.3pt;border-top:none;border-left:none;border-bottom:1pt solid windowtext;border-right:1pt solid windowtext;padding:0cm 5.4pt"> <p class="MsoNormal" style="margin:0cm 0cm 0.0001pt;line-height:normal;font-size:11pt;font-family:Calibri,sans-serif">&nbsp;</p> </td> </tr> <tr> <td width="200" valign="top" style="width:150.25pt;border-right:1pt solid windowtext;border-bottom:1pt solid windowtext;border-left:1pt solid windowtext;border-top:none;padding:0cm 5.4pt"> <p class="MsoNormal" style="margin:0cm 0cm 0.0001pt;line-height:normal;font-size:11pt;font-family:Calibri,sans-serif">&nbsp; &nbsp; </p> </td> <td width="200" valign="top" style="width:150.25pt;border-top:none;border-left:none;border-bottom:1pt solid windowtext;border-right:1pt solid windowtext;padding:0cm 5.4pt"> <p class="MsoNormal" style="margin:0cm 0cm 0.0001pt;line-height:normal;font-size:11pt;font-family:Calibri,sans-serif">&nbsp;</p> </td> <td width="200" valign="top" style="width:150.3pt;border-top:none;border-left:none;border-bottom:1pt solid windowtext;border-right:1pt solid windowtext;padding:0cm 5.4pt"> <p class="MsoNormal" style="margin:0cm 0cm 0.0001pt;line-height:normal;font-size:11pt;font-family:Calibri,sans-serif">&nbsp;</p> </td> </tr> </tbody> </table>`;
                

    return message;           
    
  }
}
