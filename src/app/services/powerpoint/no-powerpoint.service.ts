import { Observable, of } from 'rxjs';
import { PowerpointService } from './powerpoint.service';

export class NoPowerpointService extends PowerpointService {
  constructor() {
    super();
  }

  public goToFirstSlide(): Observable<boolean> {
    return of(false);
  }

  public goToLastSlide(): Observable<boolean> {
    return of(false);
  }
}
