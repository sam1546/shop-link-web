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
import { DashboardService } from '../../shared/services/dashboard.service';
import { ProductionModel } from '../../shared/models/productionModule';

@Component({
  selector: 'app-dashboard',
  templateUrl: './dashboard.component.html',
  styleUrls: ['./dashboard.component.scss'],
  animations: [routerTransition()]
})

export class DashboardComponent implements OnInit {

  constructor(private modalService: NgbModal,
    private http: HttpClient,
    public router: Router,
    private dashboardservice: DashboardService,
    public notificationMessages: ToastrManager) {

    if (localStorage.getItem('isLoggedin') == 'true') {
      this.loggedInUserName = localStorage.getItem('userName');
      //loading group data at page load
      this.loadGroups();
      this.bindDataGridALLRecords('000080000014','TM02')
    }
    else {
      this.router.navigate(['/login']);
      this.notificationMessages.errorToastr('Session logged out!! Please login again!!')
    }
  }

  waitSpinner = true;
  showSpinner = false;
  barcode: string = '';
  values: string[] = [];

  txtAckMirno: string = ''
  txtweightinton: string = ''
  txttotalsetup: string = ''
  txttotaloptions: string = ''
  txttheortime: string = ''
  section1
  loggedInUserName = ''
  productionDetails: ProductionModel;
  ngOnInit(): void {

  }
  onEnter() {
    this.findDB();
    
  }
 
  optionSelectedGroup = 'null';
  groups: any = [];
  onOptionsSelectedGroup(event) {
    console.log(event); //option value will be sent as event
    this.optionSelectedGroup= event 
    this.loadWorkCenter(event)
  }
  loadGroups(){
    this.dashboardservice.GetGroups().subscribe(groupsData => {
      this.groups = groupsData; 
    });
  }
  
  workCenter: any = [];
  optionSelectedWorkCenter = 'null';
  onOptionsSelectedWorkCenter(event) {
    this.optionSelectedWorkCenter = event;
    // console.log(event); //option value will be sent as event
  }
  loadWorkCenter(group) {
    console.log('group', group);
    
    this.dashboardservice.GetWorkCenter(group).subscribe(workCenterData => {
      this.workCenter = workCenterData;
    });
  }
 
  SECTION_TYPE = ["ShiftA", "ShiftB", "ShiftC", "ShiftD"]
  optionSelectedShift = "null";
  onOptionsSelectedShift(event) {
    console.log(event); //option value will be sent as event
  }

  clearText() {
    this.txtweightinton = ''
    this.txttotalsetup = ''
    this.txttotaloptions = ''
    this.txttheortime = ''
  }
  pddetails:any[]=[];
  bindDataGridALLRecords(mirno: string, plantCode : string) {
    this.dashboardservice.bindDataGridALLRecords(mirno, plantCode).subscribe(productionData => {
      this.pddetails = productionData; 
      console.log(this.pddetails);
      
    });
  }

  findDB() {
    this.waitSpinner = false;
    this.showSpinner = true;
    this.txtAckMirno = this.txtAckMirno.padStart(12, '0')
    this.clearText()
    if (isNumeric(this.txtAckMirno) != true) { 
      this.notificationMessages.errorToastr("Enter correct Numeric MIR Number");
      this.waitSpinner = true;
      this.showSpinner = false;
      return
    }
    this.section1 = this.txtAckMirno;
    this.section1 = this.section1.substring(0, (this.section1.length - (this.section1.length - 1)));

    console.log(typeof (this.section1));
    if (this.section1 != 0) {
      this.waitSpinner = true;
      this.showSpinner = false;
      this.notificationMessages.errorToastr("Enter valid MIR number");
      return
    }
    // this.dashboardservice.getOperationByMirno(this.txtAckMirno).subscribe((data: Response) => {
    //   var bp1 = data
    //   if (bp1 != this.section1) { 
    //     this.notificationMessages.errorToastr("MIR " + this.txtAckMirno + " is from " + data + " Plant")
    //     return;
    //   } 
    // });
    var strconfirm = confirm("MIR Number " + this.txtAckMirno + " Already Added in Shoplink:Do you want to add missing Production Orders?");
    if (strconfirm == true) {
       alert('yes')
    }
    else{
      alert('no')
    }
    
    // this.dashboardservice.findDB().subscribe((data: Response) => {
    //   console.log('data', data); 
    //   if (0 != null) { 
    //   }
    //   else {
    //     this.waitSpinner = true;
    //     this.showSpinner = false;
    //     this.notificationMessages.errorToastr('Added!!')
    //     return;
    //   }
    // });
  }










  closeModal: string;
  triggerModal(content) {  // size?: 'sm' | 'lg' | 'xl';
    this.modalService.open(content, { ariaLabelledBy: 'modal-basic-title', size: 'xl', backdrop: 'static' }).result.then((res) => {
      this.closeModal = `Closed with: ${res}`;

    }, (res) => {
      this.closeModal = `Dismissed ${this.getDismissReason(res)}`;
    });
  }

  private getDismissReason(reason: any): string {
    if (reason === ModalDismissReasons.ESC) {
      return 'by pressing ESC';
    } else if (reason === ModalDismissReasons.BACKDROP_CLICK) {
      return 'by clicking on a backdrop';
    } else {
      return `with: ${reason}`;
    }
  }

  refresh() {
    // this.router.navigate(['/dashboard']);
    window.location.reload()
  }


}
