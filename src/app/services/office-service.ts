import { Observable, of } from 'rxjs';
import { Injectable } from '@angular/core';

@Injectable({
  providedIn: 'root',
})
export class OfficeService {
  public isInOfficeSync(): boolean {
    return !!Office;
  }

  public isInOffice(): Observable<boolean> {
    return of(!!Office && !!Office.context);
  }

  public isInExcel(): Observable<boolean> {
    if (!Office || !Office.context) {
      return of(false);
    }
    console.info(Office.context.host);
    return of(Office.context.host === Office.HostType.Excel);
  }

  public isInPowerPoint(): Observable<boolean> {
    if (!Office || !Office.context) {
      return of(false);
    }
    console.info(Office.context.host);
    return of(Office.context.host === Office.HostType.PowerPoint);
  }

  public get officeAppType(): Observable<Office.HostType> {
    return of(<Office.HostType>Office.context.host);
  }

  public getSelection(type: Office.CoercionType): Observable<string> {
    return Observable.create(observer => {
      Office.context.document.getSelectedDataAsync(type, (result: Office.AsyncResult<any>) => {
        if (result.error) {
          observer.error(result.error);
        } else {
          observer.next(result.value);
        }
        observer.complete();
      });
    });
  }

  public setSelection(content: string, type: Office.CoercionType): Observable<boolean> {
    return Observable.create(observer => {
      Office.context.document.setSelectedDataAsync(
        content,
        {
          coercionType: type,
        } as Office.SetSelectedDataOptions,
        (result: Office.AsyncResult<any>) => {
          observer.next(!result.error);
          observer.complete();
        }
      );
    });
  }
}
