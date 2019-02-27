import { Component, OnInit } from '@angular/core';
import { GraphService } from '../../services/graph.service';
import { File } from '../../models/file';
import { from } from 'rxjs';
import { FormBuilder, FormGroup } from '@angular/forms';
import { OfficeService } from '../../services/office-service';

@Component({
  selector: 'app-files',
  templateUrl: './files.component.html',
})
export class FilesComponent implements OnInit {
  public queryFormModel: FormGroup = this._formBuilder.group({
    query: [''],
  });

  constructor(
    private readonly _graphService: GraphService,
    private readonly _officeService: OfficeService,
    private readonly _formBuilder: FormBuilder
  ) {
  }

  public files: Array<File> = [];

  public ngOnInit() {
  }

  public load() {
    this._officeService
      .getSelection(Office.CoercionType.Text)
      .subscribe(value => this.queryFormModel.setValue({query: value.trim().toLowerCase()}), err => console.warn(err));
  }

  public find() {
    this._graphService
      .getFiles(this.queryFormModel.value.query)
      .subscribe(files => {
        this.files = files;
      });
  }
}
