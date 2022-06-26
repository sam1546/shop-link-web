import { HttpClient } from '@angular/common/http';
import { Injectable } from '@angular/core';
import { Response, Headers, RequestOptions, Http } from '@angular/http';
import { map, catchError, tap } from 'rxjs/operators';
import { Observable } from 'rxjs/Rx'
import * as glob from '../models/global';

@Injectable({
  providedIn: 'root'
})

export class MultiprocessService {
  constructor(private http: HttpClient, private httpd: Http) {
  } 

  reversePO(fileSplit: string): Observable<any> {
    return this.http.get(glob.apiUrl + "Multiprocess/reversePO", { params: { fileSplit:fileSplit } })
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
 

}
