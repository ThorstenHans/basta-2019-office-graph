import { Injectable } from '@angular/core';
import { Client } from '@microsoft/microsoft-graph-client';

import { AuthService } from './auth.service';
import { File } from '../models/file';
import { Observable } from 'rxjs';

@Injectable({
  providedIn: 'root',
})
export class GraphService {
  private graphClient: Client;

  constructor(private authService: AuthService) {
    // Initialize the Graph client
    this.graphClient = Client.init({
      authProvider: async done => {
        // Get the token from the auth service
        const token = this.authService.token;
        if (token) {
          done(null, token);
        } else {
          done('Could not get an access token', null);
        }
      },
    });
  }

  public getProfile(): Observable<any> {
    return Observable.create(observer => {
      this.graphClient.api('/me')
        .get()
        .then(result => {
          observer.next(result);
          observer.complete();
        }).catch(error => {
        observer.next(null);
        observer.complete();
      });
    });
  }

  public getFiles(query: string): Observable<Array<File>> {
    // wrap the promise based API to Observables ❤️
    return Observable.create(observer => {
      // use the graphClient to query the files api
      // use select() to implement projection
      // use orderby() to sort
      // specify the HTTP Method using the corresponding method
      // emit result's value
      this.graphClient
        .api(`/me/drive/root/search(q='${encodeURIComponent(query)}.pptx')`)
        .select('name,webUrl,lastModifiedBy')
        .orderby('createdDateTime DESC')
        .get()
        .then(result => {
          observer.next(result.value);
          observer.complete();
        })
        .catch(error => {
          observer.error(error);
          observer.complete();
        });
    });

  }
}
