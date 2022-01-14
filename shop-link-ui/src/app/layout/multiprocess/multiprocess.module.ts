import { CommonModule } from '@angular/common';
import { NgModule } from '@angular/core';
import { NgbAlertModule, NgbCarouselModule } from '@ng-bootstrap/ng-bootstrap';
import { StatModule } from '../../shared'; 
import { MultiprocessRoutingModule } from './multiprocess-routing-module';

import { MultiprocessComponent } from './multiprocess.component';  
import { FormsModule } from '@angular/forms';
import {} from 'fs' 
@NgModule({
    imports: [
        CommonModule, 
        NgbCarouselModule, 
        NgbAlertModule, 
        MultiprocessRoutingModule, 
        StatModule,
        FormsModule, 
    ],
    declarations: [MultiprocessComponent],
    bootstrap: [MultiprocessComponent],
})
export class MultiprocessModule {}
