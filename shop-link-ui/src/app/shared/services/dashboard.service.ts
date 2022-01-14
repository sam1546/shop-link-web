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
  public findDB(): Observable<any> {
    return this.httpd.get(glob.devApiUrl + "Dashboard/findDB")
      .pipe(map((response: any) => {
        return response.json();
      }), catchError((error: Response) => {
        return "error"
      }));
  }


  bindDataGridALLRecords(mirno:string, plantCode:string): Observable<any> {
    return this.http.get(glob.devApiUrl + "Dashboard/bindDataGridALLRecords", { params: { mirno: mirno, plantCode:plantCode }})
    .pipe(map((response: any) => { 
      return response;
    }), catchError((error: Response) => {
      return "error"
    }));
  }

  public GetWorkCenter(group:any): Observable<any> {
      return this.httpd.get(glob.devApiUrl + "Dashboard/loadWorkCenters", { params: { group: group }})
        .pipe(map((response: any) => {
          return response.json();
        }), catchError((error: Response) => {
          return "error"
        })); 
  }

  public GetGroups(): Observable<any> {
    return this.httpd.get(glob.devApiUrl + "Dashboard/loadGroups")
      .pipe(map((response: any) => {
        return response.json();
      }));
  }
  public getOperationByMirno(mirno: string): Observable<any> {
    return this.httpd.get(glob.devApiUrl + "Dashboard/getOperationByMirno", { params: { mirno: mirno }})
      .pipe(map((response: any) => { 
        return response.json();
      }));
  }

}
