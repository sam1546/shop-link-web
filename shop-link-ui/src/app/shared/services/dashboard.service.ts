import { HttpClient } from '@angular/common/http';
import { Injectable } from '@angular/core';
import { Response, Headers, RequestOptions, Http } from '@angular/http';
import { map, catchError, tap } from 'rxjs/operators';
import { Observable } from 'rxjs/Rx'
import * as glob from '../models/global';

@Injectable({
  providedIn: 'root'
})

export class DashboardService {
  constructor(private http: HttpClient, private httpd: Http) {
  }
  public findDB(txtAckMirno: string): Observable<any> {
    return this.httpd.get(glob.apiUrl + "Dashboard/findDB", { params: { mirono: txtAckMirno } })
      .pipe(map((response: any) => {
        return response.json();
      }), catchError((error: Response) => {
        return "error"
      }));
  }


  bindDataGridALLRecords(mirno: string, plantCode: string, poType:string): Observable<any> {
    return this.http.get(glob.apiUrl + "Dashboard/bindDataGridALLRecords", { params: { mirno: mirno, plantCode: plantCode, poType:poType } })
      .pipe(map((response: any) => {
        return response;
      }), catchError((error: Response) => {
        return "error"
      }));
  }

  public GetWorkCenter(group: any): Observable<any> {
    return this.httpd.get(glob.apiUrl + "Dashboard/loadWorkCenters", { params: { group: group } })
      .pipe(map((response: any) => {
        return response.json();
      }), catchError((error: Response) => {
        return "error"
      }));
  }

  public GetGroups(): Observable<any> {
    return this.httpd.get(glob.apiUrl + "Dashboard/loadGroups")
      .pipe(map((response: any) => {
        return response.json();
      }));
  }
  public getBpByMirno(mirno: string): Observable<any> {
    return this.httpd.get(glob.apiUrl + "Dashboard/getBpByMirno", { params: { mirno: mirno } })
      .pipe(map((response: any) => {
        return response.json();
      }));
  }

  public getCalculations(mirno: string, plantcode:string, poType:string): Observable<any> {
    return this.httpd.get(glob.apiUrl + "Dashboard/getCalculations", { params: { mirno: mirno, plantcode:plantcode, poType:poType } })
      .pipe(map((response: any) => {
        return response.json();
      }));
  }

  public getOperationsByMirno(mirno: string): Observable<any> {
    return this.httpd.get(glob.apiUrl + "Dashboard/getOperationsByMirno", { params: { mirno: mirno } })
      .pipe(map((response: any) => {
        return response.json();
      }));
  }

  public gettotalWO_Totalreleased(mirno: string): Observable<any> {
    return this.httpd.get(glob.apiUrl + "Dashboard/gettotalWO_Totalreleased", { params: { mirno: mirno } })
      .pipe(map((response: any) => {
        return response.json();
      }));
  }

  public releasePO(mirno: string, plantCode:string): Observable<any> {
    return this.httpd.get(glob.apiUrl + "Dashboard/releasePO", { params: { mirno: mirno, plantCode:plantCode } })
      .pipe(map((response: any) => {
        return response.json();
      }));
  }

  public ackPO(mirno: string): Observable<any> {
    return this.httpd.get(glob.apiUrl + "Dashboard/ackPO", { params: { mirno: mirno} })
      .pipe(map((response: any) => {
        return response.json();
      }));
  }
  
  public allocate(mirno: string, plantCode : string, comboBox_MachineName:string, cmb_Group: string, cmbShift:string, txtRack:string): Observable<any> {
    return this.httpd.get(glob.apiUrl + "Dashboard/allocate", { params: { mirno: mirno, plantCode:plantCode,comboBox_MachineName:comboBox_MachineName,cmb_Group:cmb_Group, cmbShift:cmbShift, txtRack:txtRack} })
      .pipe(map((response: any) => {
        return response.json();
      }));
  }
  
  bindDataGridAfterAllocate(mirno: string, plantCode: string): Observable<any> {
    return this.http.get(glob.apiUrl + "Dashboard/bindDataGridAfterAllocate", { params: { mirno: mirno, plantCode: plantCode} })
      .pipe(map((response: any) => {
        return response;
      }), catchError((error: Response) => {
        return "error"
      }));
  }

  public getCalculationsAfterAllocate(mirno: string, plantcode:string): Observable<any> {
    return this.httpd.get(glob.apiUrl + "Dashboard/getCalculationsAfterAllocate", { params: { mirno: mirno, plantcode:plantcode} })
      .pipe(map((response: any) => {
        return response.json();
      }));
  }

  public insertUpdateDelete(query: string): Observable<any> {
    return this.httpd.get(glob.apiUrl + "Dashboard/insertUpdateDelete", { params: { query: query} })
      .pipe(map((response: any) => {
        return response.json();
      }));
  }

  public onLoadScreen(mirno: string, rdrFicep: boolean, rdrVernet: boolean, rdrDrilling: boolean): Observable<any> {
    return this.httpd.get(glob.apiUrl + "Dashboard/onLoadScreen", { params: { mirno: mirno, rdrFicep: rdrFicep, rdrVernet: rdrVernet, rdrDrilling: rdrDrilling } })
      .pipe(map((response: any) => {
        return response.json();
      }));
  }

  public totweight1(mirno: string): Observable<any> {
    return this.httpd.get(glob.apiUrl + "Dashboard/totweight1", { params: { mirno: mirno} })
      .pipe(map((response: any) => {
        return response.json();
      }));
  }

  public onBalPunchMIR(dateTimePicker1:string): Observable<any> {
    return this.httpd.get(glob.apiUrl + "Dashboard/onBalPunchMIR", { params: {dateTimePicker1:dateTimePicker1} })
      .pipe(map((response: any) => {
        return response.json();
      }));
  }

  bindDataGridOnPunchMIR(plantCode: string, dateTimePicker1: string): Observable<any> {
    return this.http.get(glob.apiUrl + "Dashboard/bindDataGridOnPunchMIR", { params: { plantCode: plantCode, dateTimePicker1: dateTimePicker1} })
      .pipe(map((response: any) => {
        return response;
      }), catchError((error: Response) => {
        return "error"
      }));
  }


  public onBalAllocateMIR(): Observable<any> {
    return this.httpd.get(glob.apiUrl + "Dashboard/onBalAllocateMIR", { params: {} })
      .pipe(map((response: any) => {
        return response.json();
      }));
  }

  bindDataGridOnAllocateMIR(plantCode: string): Observable<any> {
    return this.http.get(glob.apiUrl + "Dashboard/bindDataGridOnBalAllocateMIR", { params: { plantCode: plantCode} })
      .pipe(map((response: any) => {
        return response;
      }), catchError((error: Response) => {
        return "error"
      }));
  }

  bindDataGridOnradioButton1(plantCode: string): Observable<any> {
    return this.http.get(glob.apiUrl + "Dashboard/bindDataGridOnradioButton1", { params: { plantCode: plantCode} })
      .pipe(map((response: any) => {
        return response;
      }), catchError((error: Response) => {
        return "error"
      }));
  }

}
