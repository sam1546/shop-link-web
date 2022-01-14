import { HttpClient, HttpErrorResponse } from '@angular/common/http';
import { Injectable } from '@angular/core';
import { Response, Headers, RequestOptions, Http } from '@angular/http';
import { throwError } from 'rxjs';
import { map, catchError, tap } from 'rxjs/operators';
import { Observable } from 'rxjs/Rx'
import * as glob from '../models/global';
import { ToastrManager } from 'ng6-toastr-notifications';

@Injectable({
  providedIn: 'root'
})

export class UsermasterService {
  constructor(private http: HttpClient, private httpd: Http, public notificationMessages: ToastrManager) {
  }
  public UserLogin(userId: string, userPassword): Observable<any> {
    try {
      return this.httpd.get(glob.devApiUrl + "Login/User", { params: { UserId: userId, UserPassword: userPassword } })
        .pipe(map((response: any) => {
          return response.json();
        })
          // , catchError((error: Response) => {
          //   return "error"
          // })
        );
    }
    catch {
      var res: any = {
        code: 404,
        error: "connection error"
      };
      return res;

    }

  }
}
