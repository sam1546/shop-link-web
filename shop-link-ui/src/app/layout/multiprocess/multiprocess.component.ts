import { Component, OnInit, ViewChild, AfterViewInit, EventEmitter, Input, OnDestroy, Output, } from '@angular/core';
import { routerTransition } from '../../router.animations';
import { ModalDismissReasons, NgbModal } from '@ng-bootstrap/ng-bootstrap';
import xml2js from 'xml2js';
import { HttpClient, HttpHeaders, JsonpClientBackend } from '@angular/common/http';
import * as $ from 'jquery'
// var dt = require( 'datatables.net' )();
import { Router } from '@angular/router';
import { isNumeric } from 'rxjs/internal-compatibility';
import { ToastrManager } from 'ng6-toastr-notifications';
import { MultiprocessService } from '../../shared/services/multiprocesses.service';
import { ProductionModel } from '../../shared/models/productionModule';
import { stringify } from 'querystring';

@Component({
  selector: 'app-multiprocess',
  templateUrl: './multiprocess.component.html',
  styleUrls: ['./multiprocess.component.scss'],
  animations: [routerTransition()]
})
export class MultiprocessComponent implements OnInit {
  loggedInUserName = ''
  machineName = localStorage.getItem("machineName")
  department = localStorage.getItem("department")
  Role = localStorage.getItem("role")
  plantCode = localStorage.getItem("plancode")
  plantAddress = localStorage.getItem("plantAddress")
  userName = localStorage.getItem("userName")

  richTextBox1 : string =""
  constructor(private modalService: NgbModal,
    private http: HttpClient,
    public router: Router,
    private multoprocess: MultiprocessService,
    public notificationMessages: ToastrManager) { 
  }

  ngOnInit(): void {
    if (localStorage.getItem('isLoggedin') == 'true') {
      this.loggedInUserName = localStorage.getItem('userName');
    }
    else {
      this.notificationMessages.errorToastr('Session logged out!! Please login again!!')
      this.router.navigate(['/login']);
    }
  }

  onReversePOClick(){
    if(this.richTextBox1.trim()==null || this.richTextBox1.trim()==''){
      this.notificationMessages.errorToastr('No WorkOrder to reverse, please enter WorkOrders!')
      return
    }
    var strconfirm = confirm("Are you sure you want to Reverse records?");
    if (strconfirm == true) {
      var files = this.richTextBox1;
      var fileSplit: any = files.split('\n'); 
      this.multoprocess.reversePO(fileSplit).subscribe((result: Response) => {
        var data = result
        
      });
    }

  }


}
