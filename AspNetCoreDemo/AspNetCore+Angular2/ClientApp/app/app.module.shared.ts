import { NgModule } from '@angular/core';
import { RouterModule } from '@angular/router';

import { AppComponent } from './components/app/app.component'
import { NavMenuComponent } from './components/navmenu/navmenu.component';
import { HomeComponent } from './components/home/home.component';
import { ExcelTemplateDemoComponent } from './components/excelTemplateDemo/excelTemplateDemo.component';
import { ProgrammingDemoComponent } from './components/programmingDemo/programmingDemo.component';
import { ExcelIODemoComponent } from './components/excelIODemo/excelIODemo.component';
import { SpreadSheetsModule } from './gc.spread.sheets.angular2.10.2.0';

export const sharedConfig: NgModule = {
    bootstrap: [ AppComponent ],
    declarations: [
        AppComponent,
        NavMenuComponent,
        ExcelIODemoComponent,
        ExcelTemplateDemoComponent,
        ProgrammingDemoComponent,
        HomeComponent,
    ],
    imports: [
        SpreadSheetsModule,
        RouterModule.forRoot([
            { path: '', redirectTo: 'home', pathMatch: 'full' },
            { path: 'home', component: HomeComponent },
            { path: 'excelTemplateDemo', component: ExcelTemplateDemoComponent },
            { path: 'programmingDemo', component: ProgrammingDemoComponent },
            { path: 'excelIODemo', component: ExcelIODemoComponent },
            { path: '**', redirectTo: 'home' }
        ])
    ]
};
