import { Component, OnInit } from '@angular/core';
import { AuthService } from '../../services/auth.service';
import { Router } from '@angular/router';
import { Observable, from } from 'rxjs';
import { GraphService } from '../../services/graph.service';
import { pluck, switchMap, tap } from 'rxjs/operators';

@Component({
  selector: 'app-navigation',
  templateUrl: './navigation.component.html',
})
export class NavigationComponent implements OnInit {
  public username: string;


  constructor(private readonly _authService: AuthService, private readonly _graphService: GraphService, private readonly _router: Router) {
  }

  public get isLoggedIn(): boolean {
    return this._authService.hasToken;
  }

  public login() {
    console.log('login');
    this._authService.signIn()
      .pipe(
        switchMap(res => this._graphService.getProfile())
      )
      .subscribe(profile => {
        if (profile) {
          this.username = profile['displayName'];
        }
      });
  }

  public ngOnInit() {
    if (this._authService.hasToken) {
      this._graphService.getProfile()
        .pipe(
          pluck('displayName')
        ).subscribe((name: string) => this.username = name);
    }
  }
}
