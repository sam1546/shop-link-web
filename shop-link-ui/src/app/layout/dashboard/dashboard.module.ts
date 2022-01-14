import { CommonModule } from '@angular/common';
import { NgModule } from '@angular/core';
import { NgbAlertModule, NgbCarouselModule } from '@ng-bootstrap/ng-bootstrap';
import { StatModule } from '../../shared'; 
import { DashboardRoutingModule } from './dashboard-routing.module';

import { DashboardComponent } from './dashboard.component';  
import { FormsModule } from '@angular/forms';
import {} from 'fs' 
@NgModule({
    imports: [
        CommonModule, 
        NgbCarouselModule, 
        NgbAlertModule, 
        DashboardRoutingModule, 
        StatModule,
        FormsModule, 
    ],
    declarations: [DashboardComponent],
    bootstrap: [DashboardComponent],
})
export class DashboardModule {}
