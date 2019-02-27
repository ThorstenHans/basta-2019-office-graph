import { NgModule } from '@angular/core';
import { Routes, RouterModule } from '@angular/router';
import { FilesComponent } from './components/files/files.component';
import { StartComponent } from './components/start/start.component';
import { HasTokenGuard } from './guards/has-token.guard';

const routes: Routes = [
  {path: 'start', component: StartComponent},
  {path: 'files', canActivate: [HasTokenGuard], component: FilesComponent},
  {path: '', pathMatch: 'full', redirectTo: '/start'},
];

@NgModule({
  imports: [RouterModule.forRoot(routes, {useHash: true})],
  exports: [RouterModule],
})
export class AppRoutingModule {
}
