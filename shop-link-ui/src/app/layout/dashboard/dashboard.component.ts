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
import { stringify } from 'querystring';

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
      // this.bindDataGridALLRecords('000080000014','TM02')
    }
    else {
      this.notificationMessages.errorToastr('Session logged out!! Please login again!!')
      this.router.navigate(['/login']);
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
  plantCode = localStorage.getItem("plantcode")
  poType: string = ''
  productionDetails: ProductionModel;
  ngOnInit(): void {

  }
  onEnter() {
    this.findDB();

  }

  optionSelectedGroup = 'null';
  groups: any = [];
  onOptionsSelectedGroup(event) {
    this.optionSelectedGroup = event
    this.loadWorkCenter(event)
  }
  loadGroups() {
    this.dashboardservice.GetGroups().subscribe(groupsData => {
      this.groups = groupsData;
    });
  }

  workCenter: any = [];
  optionSelectedWorkCenter = 'null';
  onOptionsSelectedWorkCenter(event) {
    this.optionSelectedWorkCenter = event;
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
    this.optionSelectedShift = event;
  }

  clearText() {
    this.txtweightinton = ''
    this.txttotalsetup = ''
    this.txttotaloptions = ''
    this.txttheortime = ''
  }
  // pddetails: any[] = [];
  // bindDataGridALLRecords(mirno: string, plantCode: string) {
  //   this.dashboardservice.bindDataGridALLRecords(mirno, plantCode).subscribe(productionData => {
  //     this.pddetails = productionData;
  //     console.log(this.pddetails);

  //   });
  // }

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
    this.dashboardservice.getBpByMirno(this.txtAckMirno).subscribe((data: Response) => {
      var bp1 = data
      if (bp1 != this.section1) {
        this.waitSpinner = true;
        this.showSpinner = false;
        this.notificationMessages.errorToastr("MIR " + this.txtAckMirno + " is from " + data + " Plant")
        return;
      }
    });

    var strconfirm = confirm("MIR Number " + this.txtAckMirno + " Already Added in Shoplink:Do you want to add missing Production Orders?");
    if (strconfirm == true) {
      this.dashboardservice.findDB(this.txtAckMirno).subscribe((data: Response) => {
        console.log('data', data);
        if (0 != null) {
        }
        else {
          this.waitSpinner = true;
          this.showSpinner = false;
          this.notificationMessages.errorToastr('Added!!')
          return;
        }
      });
    }
    else {

    }

  }

  pddetails: any[] = [];
  onSearchClick() {
    this.waitSpinner = false;
    this.showSpinner = true;
    this.clearText()
    this.txtAckMirno = this.txtAckMirno.padStart(12, '0')
    this.dashboardservice.bindDataGridALLRecords(this.txtAckMirno, this.plantCode, this.poType = '').subscribe(result => {
      this.pddetails = result;

      if (this.pddetails.length == 0 || this.pddetails == null) {
        this.notificationMessages.errorToastr("MIR " + this.txtAckMirno + " not available in SHOPlink database")
        return;
      }

      this.dashboardservice.getCalculations(this.txtAckMirno, this.plantCode, this.poType = '').subscribe(result => {
        this.txtweightinton = result[0].totalWheight
        this.txttotalsetup = result[0].rSno
        this.txttotaloptions = result[0].totalOpns
        this.txttheortime = result[0].runTime
      });
    });
    this.waitSpinner = true;
    this.showSpinner = false;
  }

  releasePO() {
    this.dashboardservice.releasePO(this.txtAckMirno, this.plantCode).subscribe(result => {
      var totalPO = result.totalPO
      var releasedPO = result.releasedPO;
      var totalreleased = result.totalreleased;
      if (releasedPO > 0) {
        this.dashboardservice.getCalculations(this.txtAckMirno, this.plantCode, this.poType = 'Primary').subscribe(result => {
          this.txtweightinton = result[0].totalWheight
          this.txttotalsetup = result[0].rSno
          this.txttotaloptions = result[0].totalOpns
          this.txttheortime = result[0].runTime
        });

        this.dashboardservice.bindDataGridALLRecords(this.txtAckMirno, this.plantCode, this.poType = 'Primary').subscribe(result => {
          this.pddetails = result;

        });

        var compltedPO = releasedPO + totalreleased;
        if (totalPO != releasedPO + totalreleased)
          alert(totalPO - releasedPO + " Workorders are pending for release. Please repeat the process to release pending workorders");
        else
          alert("Total " + compltedPO + " Production Orders are Released out of " + totalPO + ", MIR No: " + this.txtAckMirno);
        this.txtAckMirno = "";


      }
    })
  }

  onAckmirClick() {
    this.clearText()
    this.waitSpinner = false;
    this.showSpinner = true;
    this.txtAckMirno = this.txtAckMirno.padStart(12, '0')
    if (isNumeric(this.txtAckMirno) != true) {
      this.notificationMessages.errorToastr("Enter correct Numeric MIR Number");
      this.waitSpinner = true;
      this.showSpinner = false;
      return
    }
    this.dashboardservice.getOperationsByMirno(this.txtAckMirno).subscribe(result => {
      if (result[0].length > 0) {
        var flag_Ack = result[0].flag_Ack
        var bp = result[0].bp;
        if (bp != this.plantCode) {
          this.notificationMessages.errorToastr("MIR " + this.txtAckMirno + " is from " + bp + " Plant. Cannot change WorkCenter")
          this.waitSpinner = true;
          this.showSpinner = false;
          return
        }
        if (flag_Ack == 'TRUE') {
          this.dashboardservice.gettotalWO_Totalreleased(this.txtAckMirno).subscribe(result => {
            if (result == true) {
              var strconfirm = confirm("All Workorders are not released in MIR: " + this.txtAckMirno + ", Do you want to release the pending workorders?");
              if (strconfirm == true) {
                this.releasePO();
              }
              else {
                this.waitSpinner = true;
                this.showSpinner = false;
                return
              }
            }
            else {
              var strconfirm = confirm("All Workorders are released in MIR: " + this.txtAckMirno + ",  Do you want to repeate the releasing Process?");
              if (strconfirm == true) {
                this.releasePO();
              }
              else {
                this.waitSpinner = true;
                this.showSpinner = false;
                return
              }
            }

          });
        }
      }
      else {
        alert("MIR is not available Shoplink database");
      }
    });

  }

  onAckPOclick() {
    this.waitSpinner = false;
    this.showSpinner = true;
    this.txtAckMirno = this.txtAckMirno.padStart(12, '0')
    if (isNumeric(this.txtAckMirno) != true) {
      this.notificationMessages.errorToastr("Enter correct Numeric Production Order Number");
      this.waitSpinner = true;
      this.showSpinner = false;
      return
    }

    this.dashboardservice.ackPO(this.txtAckMirno).subscribe(result => {
      this.notificationMessages.errorToastr(result);
      this.waitSpinner = true;
      this.showSpinner = false;
      return
    })

  }


  allocate() {
    this.clearText()
    this.waitSpinner = false;
    this.showSpinner = true;
    this.txtAckMirno = this.txtAckMirno.padStart(12, '0')
    if (isNumeric(this.txtAckMirno) != true) {
      this.notificationMessages.errorToastr("Enter correct Numeric MIR Number");
      this.waitSpinner = true;
      this.showSpinner = false;
      return
    }

    if (this.txtAckMirno.length != 12) {
      this.notificationMessages.errorToastr("Enter correct MIR Number");
      return;
    }

    this.section1 = this.txtAckMirno;
    this.section1 = this.section1.substring(0, (this.section1.length - (this.section1.length - 1)));
    if (this.section1 != 0) {
      this.waitSpinner = true;
      this.showSpinner = false;
      this.notificationMessages.errorToastr("Enter valid MIR number");
      return
    }
    var txtRack = ""
    this.dashboardservice.allocate(this.txtAckMirno, this.plantCode, this.optionSelectedGroup, this.optionSelectedWorkCenter, this.optionSelectedShift, txtRack).subscribe(result => {
      alert(result);
      this.dashboardservice.bindDataGridAfterAllocate(this.txtAckMirno, this.plantCode).subscribe(result => {
        this.pddetails = result;
      });

      this.dashboardservice.getCalculationsAfterAllocate(this.txtAckMirno, this.plantCode).subscribe(calresult => {
        this.txtweightinton = calresult[0].totalWheight
        this.txttotalsetup = calresult[0].rSno
        this.txttotaloptions = calresult[0].totalOpns
        this.txttheortime = calresult[0].runTime
      });
      this.waitSpinner = true;
      this.showSpinner = false;
    });


  }
  onEnterPlanMIR() {
    this.allocate()
  }

  onAssignClick() {
    this.allocate()
  }
  onMassUpdateClick() {
    this.notificationMessages.errorToastr("Work in progress");
    return;
  }
  onDeleteMIRClick() {
    this.waitSpinner = false;
    this.showSpinner = true;
    if (isNumeric(this.txtAckMirno) != true) {
      this.notificationMessages.errorToastr("Enter correct Numeric MIR Number");
      this.waitSpinner = true;
      this.showSpinner = false;
      return
    }
    if (this.txtAckMirno != "") {
      var strconfirm = confirm("Are you sure you want to delete record?");
      if (strconfirm == true) {
        var query = "delete from Operations where Mirno='" + this.txtAckMirno + "' and MachineName is null";

        this.dashboardservice.insertUpdateDelete(query).subscribe(result => {
          if (result == 'true') {
            this.notificationMessages.successToastr("Record deleted successfully");
          }
          else {
            this.notificationMessages.errorToastr("MIR is Allocated to Machine First Deallocate then delete MIR");
          }
        });

        this.dashboardservice.bindDataGridALLRecords(this.txtAckMirno, this.plantCode, this.poType = '').subscribe(result => {
          this.pddetails = result;
        });

        this.waitSpinner = true;
        this.showSpinner = false;
      }
      else {
        this.waitSpinner = true;
        this.showSpinner = false;
        return
      }
    }
    else {
      this.waitSpinner = true;
      this.showSpinner = false;
      this.notificationMessages.errorToastr("MIR should not be blank");
      return
    }

  }

  onMainScreenClick() {
    this.waitSpinner = false;
    this.showSpinner = true;
    this.dashboardservice.bindDataGridALLRecords(this.txtAckMirno, this.plantCode, this.poType = '').subscribe(result => {
      this.pddetails = result;
      this.waitSpinner = true;
      this.showSpinner = false;
    });
  }
  
  rdrFicep = false;
  rdrVernet = false
  rdrDrilling = false
  onRadiobuttonClick(flag) {
    if (flag == 1) {
      this.rdrFicep = true
      this.rdrVernet = false
      this.rdrDrilling = false
    }
    if (flag == 2) {
      this.rdrFicep = false
      this.rdrVernet = true
      this.rdrDrilling = false
    }
    if (flag == 3) {
      this.rdrFicep = false
      this.rdrVernet = false
      this.rdrDrilling = true
    }
  }

  onLoadScreenClick() {
    this.clearText()
    if (this.rdrFicep == false && this.rdrVernet == false && this.rdrDrilling == false) {
      this.notificationMessages.errorToastr("Please select machine type 1. FICEP 2. VERNET 3. DRILLING");
      return;
    }

    if (this.txtAckMirno != "") {

      this.dashboardservice.onLoadScreen(this.txtAckMirno, this.rdrFicep, this.rdrVernet, this.rdrDrilling).subscribe(result => {
        this.txtweightinton = result[0].totalWheight
        this.txttotalsetup = result[0].rSno
        this.txttotaloptions = result[0].totalOpns
        this.txttheortime = result[0].runTime

      });
    }
    else { 
      this.dashboardservice.totweight1(this.txtAckMirno).subscribe(result => {
        if(result!=""){
          this.txtweightinton = result[0].totalWheight
          this.txttotalsetup = result[0].rSno
          this.txttotaloptions = result[0].totalOpns
          this.txttheortime = result[0].runTime
        }
    });
    }
    this.waitSpinner = true;
    this.showSpinner = false;
  }

  onBalPunchMIRClick(){
    var dateTimePicker1 = Date.now()
    this.dashboardservice.onBalPunchMIR(dateTimePicker1.toString()).subscribe(result => {
      if(result!=""){
        this.txtweightinton = result[0].totalWheight
        this.txttotalsetup = result[0].rSno
        this.txttotaloptions = result[0].totalOpns
        this.txttheortime = result[0].runTime
      }

      this.dashboardservice.bindDataGridOnPunchMIR(this.plantCode, dateTimePicker1.toString()).subscribe(result => {
        this.pddetails = result;
      });

  }); 
  }
 
  onDeleteMIRLeftClick(){
    this.waitSpinner = false;
    this.showSpinner = true;
    this.clearText();
    if (isNumeric(this.txtAckMirno) != true) {
      this.notificationMessages.errorToastr("Enter correct Numeric MIR Number");
      this.waitSpinner = true;
      this.showSpinner = false;
      return
    }

    this.dashboardservice.getOperationsByMirno(this.txtAckMirno).subscribe(result => {
      if (result[0].length > 0) {
        var flag_Ack = result[0].flag_Ack
      }
      else{
        this.notificationMessages.errorToastr("Data is not Available in SHOPLink");
        return;
      }
      if (this.txtAckMirno != "") {
        var strconfirm = confirm("Are you sure you want to delete record?");
        if (strconfirm == true) {
          var query = "delete from Operations where Mirno='" + this.txtAckMirno + "' ";
  
          this.dashboardservice.insertUpdateDelete(query).subscribe(result => {
            if (result == 'true') {
              this.notificationMessages.successToastr("Record deleted successfully");
            } 
            else{
              this.notificationMessages.errorToastr("Ooops problem occured while deleting records.!");
            }
          });
          this.dashboardservice.bindDataGridALLRecords(this.txtAckMirno, this.plantCode, this.poType = '').subscribe(result => {
            this.pddetails = result;
          });
  
          this.waitSpinner = true;
          this.showSpinner = false;
        }
        else {
          this.waitSpinner = true;
          this.showSpinner = false;
          return
        }
      }
      else {
        this.waitSpinner = true;
        this.showSpinner = false;
        this.notificationMessages.errorToastr("MIR should not be blank");
        return
      }

    })

  }

  onBalAllocateClick(){
    var dateTimePicker1 = Date.now()
    this.dashboardservice.onBalAllocateMIR().subscribe(result => {
      if(result!=""){
        this.txtweightinton = result[0].totalWheight
        this.txttotalsetup = result[0].rSno
        this.txttotaloptions = result[0].totalOpns
        this.txttheortime = result[0].runTime
      }
      this.dashboardservice.bindDataGridOnAllocateMIR(this.plantCode).subscribe(result => {
        this.pddetails = result;
      });

  }); 
  }

  radioButton1 = false;
  radioButton2 = false 
  radioButton1Click(flag) {
    if (flag == 1) { 
      this.radioButton1 = true
      this.dashboardservice.bindDataGridOnradioButton1(this.plantCode).subscribe(result => {
        this.pddetails = result;
      });
    }
    if (flag == 2) {
      this.radioButton2 = true
      this.pddetails = [];
    } 
  }

  onExportToExcelClick(){
    this.notificationMessages.errorToastr("Work in progress");
  }

  onShop1(){
    this.notificationMessages.errorToastr("Work in progress");
  }

  onButton1click(){
    this.notificationMessages.errorToastr("Work in progress");
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
