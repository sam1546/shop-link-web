import { CommonModule } from '@angular/common';
import { HttpClient,HttpClientModule } from '@angular/common/http';
import { NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';
import { BrowserAnimationsModule } from '@angular/platform-browser/animations';
import { AppRoutingModule } from './app-routing.module';
import { AppComponent } from './app.component';
import { AuthGuard } from './shared';
import { LanguageTranslationModule } from './shared/modules/language-translation/language-translation.module';
import { HttpModule } from '@angular/http'; 
import { ToastrModule } from 'ng6-toastr-notifications';  
import { FormsModule, ReactiveFormsModule } from '@angular/forms';    
import { NgbModule } from '@ng-bootstrap/ng-bootstrap'; 

@NgModule({
    imports: [
        CommonModule,
        BrowserModule,  
        FormsModule,
        ReactiveFormsModule,
        BrowserAnimationsModule,
        HttpClientModule,  HttpModule,
        LanguageTranslationModule,
        AppRoutingModule,
        AppRoutingModule,ToastrModule.forRoot(),
        NgbModule, 
    ],
    declarations: [AppComponent],
    providers: [AuthGuard],
    bootstrap: [AppComponent]
})
export class AppModule {}
