import { Observable, of } from 'rxjs';
import { PowerpointService } from './powerpoint.service';

export class PowerpointAddinService extends PowerpointService {
  constructor() {
    super();
  }

  public goToFirstSlide(): Observable<boolean> {

    return Observable.create(observer => {
      Office.context.document.goToByIdAsync(Office.Index.First, Office.GoToType.Index, (result: Office.AsyncResult<any>) => {
        observer.next(!result.error);
        observer.complete();
      });
    });
  }

  public goToLastSlide(): Observable<boolean> {
    return Observable.create(observer => {
      Office.context.document.goToByIdAsync(Office.Index.Last, Office.GoToType.Index, (result: Office.AsyncResult<any>) => {
        observer.next(!result.error);
        observer.complete();
      });
    });
  }
}
