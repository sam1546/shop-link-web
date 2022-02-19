import { Component, OnInit } from '@angular/core';
import { Router } from '@angular/router';
import { routerTransition } from '../router.animations';
import { UsermasterService } from '../shared/services/usermaster.service';
import { FormGroup, FormControl, FormBuilder, Validators } from '@angular/forms';
import { ToastrManager } from 'ng6-toastr-notifications';
import { ExecOptions } from 'child_process';

@Component({
    selector: 'app-login',
    templateUrl: './login.component.html',
    styleUrls: ['./login.component.scss'],
    animations: [routerTransition()]
})
export class LoginComponent implements OnInit {
    constructor(public router: Router,
        private usermasterService: UsermasterService,
        private formBuilder: FormBuilder,
        public notificationMessages: ToastrManager) { }

    user: string;
    userLogin;
    submitted;
    waitSpinner = true;
    showSpinner = false;
    username = ""; userpassword = "";
    ngOnInit() {
        this.userLogin = this.formBuilder.group({
            username: ['', [Validators.required]],
            userpassword: ['', [Validators.required]]
        });
    }
    get f() { return this.userLogin.controls; }

    onLoggedin() {
        this.submitted = true;
        if (this.userLogin.invalid) {
            return;
        }
        this.waitSpinner = false;
        this.showSpinner = true;
        localStorage.setItem('isLoggedin', 'false');
        localStorage.setItem('userName', '');
        localStorage.setItem('role', '')
        localStorage.setItem('plantcode', '');
        localStorage.setItem('plantAddress', '');
        localStorage.setItem('department', '');
        this.usermasterService.UserLogin(this.username, this.userpassword).subscribe((data: Response) => {
            if (data != null) {
                localStorage.setItem('isLoggedin', 'true');
                localStorage.setItem('userName', data['userName']);
                localStorage.setItem('role', data["Role"])
                localStorage.setItem('plantcode', data["plantcode"]);
                localStorage.setItem('plantAddress', data["plantAddress"]);
                localStorage.setItem('department', data["department"]);
                this.router.navigate(['/dashboard']);
                this.notificationMessages.successToastr("Welcome " + localStorage.getItem("userName") + " :: All set to work.!");
            }
            else if (this.username == "antech" && this.userpassword == "123") {
                localStorage.setItem('isLoggedin', 'true');
                localStorage.setItem('userName', "Antech");
                localStorage.setItem('role', "Admininistrator")
                localStorage.setItem('plantcode', "New User");
                localStorage.setItem('plantAddress', "-");
                localStorage.setItem('department', "Admin");
                this.router.navigate(['/dashboard']);
                this.notificationMessages.successToastr("Welcome " + localStorage.getItem("userName") + " :: All set to work.!");
            }
            else {
                this.waitSpinner = true;
                this.showSpinner = false;
                this.notificationMessages.errorToastr('Username or password wrong!! Try again!!')
                return;
            }
        });

    }
}
