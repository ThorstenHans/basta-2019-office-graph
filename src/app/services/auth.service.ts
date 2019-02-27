import { Injectable } from '@angular/core';
import { Authenticator, DefaultEndpoints } from '@microsoft/office-js-helpers';
import { environment } from 'src/environments/environment';
import { Observable } from 'rxjs';

@Injectable({
  providedIn: 'root',
})
export class AuthService {
  private _authenticator: Authenticator;
  private _token: string;

  public get hasToken(): boolean {
    return this._authenticator.tokens.get(DefaultEndpoints.Microsoft) &&
      !!this._authenticator.tokens.get(DefaultEndpoints.Microsoft).access_token;
  }

  public get token(): string {
    return this._authenticator.tokens.get(DefaultEndpoints.Microsoft).access_token;
  }

  constructor() {
    this._authenticator = new Authenticator();
    this._authenticator.endpoints.registerMicrosoftAuth(environment.aad.clientId);
  }

  public signIn(): Observable<boolean> {
    return Observable.create(observer => {
      this._authenticator
        .authenticate(DefaultEndpoints.Microsoft)
        .then(token => {
          console.log('authenticated');
          console.dir(token);
          this._token = token.access_token;
          observer.next(true);
          observer.complete();
        })
        .catch(() => {
          observer.next(false);
          observer.complete();
        });
    });
  }

}
