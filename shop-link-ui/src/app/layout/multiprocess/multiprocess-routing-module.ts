import { NgModule } from '@angular/core';
import { RouterModule, Routes } from '@angular/router';
import { MultiprocessComponent } from './multiprocess.component';

const routes: Routes = [
    {
        path: '',
        component: MultiprocessComponent
    }
];

@NgModule({
    imports: [RouterModule.forChild(routes)],
    exports: [RouterModule]
})
export class MultiprocessRoutingModule {}
