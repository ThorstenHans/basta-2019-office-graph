import { enableProdMode } from '@angular/core';
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';

import { AppModule } from './app/app.module';
import { environment } from './environments/environment';
import * as OfficeHelpers from '@microsoft/office-js-helpers';

if (environment.production) {
  enableProdMode();
}

Office.onReady(() => {
  if (OfficeHelpers.Authenticator.isAuthDialog()) return;
  platformBrowserDynamic()
    .bootstrapModule(AppModule)
    .catch(err => console.error(err));
});
