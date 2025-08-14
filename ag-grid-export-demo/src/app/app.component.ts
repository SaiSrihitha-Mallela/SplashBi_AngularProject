import { Component } from '@angular/core';
import { DataGridComponent } from './data-grid/data-grid.component';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [DataGridComponent],
  templateUrl: './app.component.html'
})
export class AppComponent { }
