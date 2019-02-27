import { Observable, of } from 'rxjs';

export abstract class PowerpointService {

  abstract goToFirstSlide(): Observable<boolean>;

  abstract goToLastSlide(): Observable<boolean>;

}
