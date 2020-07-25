import { Injectable } from '@angular/core';
import { HttpClient, HttpHeaders } from '@angular/common/http';
import config from '../assets/properties/config.json';
import { map } from 'rxjs/operators';
// import 'rxjs/add/operator/map';

@Injectable({
  providedIn: 'root'
})
export class EmailServiceService {

  apiUrl = "https://timesheettera.000webhostapp.com/api";
  // apiUrl = "http://localhost/WebMailer";
  file = '';
  filename = '';
  constructor(
    private http: HttpClient
  ) { }

  sendEmail(data){
    
    data.toEmail.forEach((element, index) => {
      data.toEmail[index] = element.email
      
    });
    data.ccEmail.forEach((element, index) => {
      data.ccEmail[index] = element.email
      
    });
    data.bccEmail.forEach((element, index) => {
      data.bccEmail[index] = element.email
      
    });
    if(this.file){
      data.attachment = this.file;
      data.filename = this.filename;
    }
    
    console.log(data);
    
    
    return this.http.post(this.apiUrl+'/api',JSON.stringify(data));
  }


  attachFile(data,name){
      this.file = data;
      this.filename = name;
      return this.file
  }

  clearFile(){
    this.file = '';
      this.filename = '';
  }

}
