import { BrowserModule } from '@angular/platform-browser';
import { NgModule } from '@angular/core';

import { AppRoutingModule } from './app-routing.module';
import { FilesComponent } from './components/files/files.component';
import { ReactiveFormsModule } from '@angular/forms';
import { NavigationComponent } from './components/navigation/navigation.component';
import { RootComponent } from './components/root/root.component';
import { PowerpointAddinService } from './services/powerpoint/powerpoint-addin.service';
import { OfficeService } from './services/office-service';
import { NoPowerpointService } from './services/powerpoint/no-powerpoint.service';
import { PowerpointService } from './services/powerpoint/powerpoint.service';
import { StartComponent } from './components/start/start.component';

export function getPowerpointService(officeService: OfficeService) {
  if (officeService.isInPowerPoint()) {
    return new PowerpointAddinService();
  }
  return new NoPowerpointService();
}

@NgModule({
  declarations: [FilesComponent, NavigationComponent, RootComponent, StartComponent],
  imports: [BrowserModule, ReactiveFormsModule, AppRoutingModule],
  providers: [{provide: PowerpointService, useFactory: getPowerpointService, deps: [OfficeService]}],
  bootstrap: [RootComponent],
})
export class AppModule {
}
