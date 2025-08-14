import { Component, OnInit } from '@angular/core';
import { AgGridAngular } from 'ag-grid-angular';
import { ModuleRegistry, AllCommunityModule, ColDef, GridReadyEvent, GridApi } from 'ag-grid-community';
import { HttpClient } from '@angular/common/http';
import { CommonModule } from '@angular/common';
import { firstValueFrom } from 'rxjs';

// Import libraries for PDF, Excel, and PPTX export
import jsPDF from 'jspdf';
import 'jspdf-autotable';
import * as XLSX from 'xlsx';
import PptxGenJS from 'pptxgenjs';

ModuleRegistry.registerModules([AllCommunityModule]);

@Component({
  selector: 'app-data-grid',
  standalone: true,
  imports: [CommonModule, AgGridAngular],
  templateUrl: './data-grid.component.html',
  styleUrls: ['./data-grid.component.css']
})
export class DataGridComponent implements OnInit {

  public columnDefs: ColDef[] = [
    { headerName: 'Name', field: 'Name', sortable: true, filter: true },
    { headerName: 'Email', field: 'email', sortable: true, filter: true },
    { headerName: 'Country', field: 'country', sortable: true, filter: true },
    { headerName: 'Phone', field: 'phone', sortable: true, filter: true }
  ];

  public rowData: any[] = [];
  private gridApi!: GridApi;

  constructor(private http: HttpClient) { }

  async ngOnInit() {
    await this.loadData();
  }

  async loadData() {
    let requests = [];
    for (let i = 0; i < 10; i++) {
      requests.push(firstValueFrom(this.http.get<any>(`https://api.npoint.io/b66e5ba94ad1ae231518`)));
    }
    const responses = await Promise.all(requests);
    this.rowData = responses.flat();
  }

  onGridReady(params: GridReadyEvent) {
    this.gridApi = params.api;
  }

  // Custom Excel Export using the 'xlsx' library
  exportAsExcel() {
    if (this.rowData.length === 0) return;
    const ws: XLSX.WorkSheet = XLSX.utils.json_to_sheet(this.rowData);
    const wb: XLSX.WorkBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, 'aggrid-data.xlsx');
  }

  // Custom PDF Export using 'jspdf' and 'jspdf-autotable' to create a table
  exportAsPdf() {
    if (this.rowData.length === 0) return;
    const doc = new jsPDF();
    const headers = this.columnDefs.map(col => col.headerName!);
    const body = this.rowData.map(row => this.columnDefs.map(col => row[col.field!]));

    (doc as any).autoTable({
      head: [headers],
      body: body,
    });

    doc.save('aggrid-data.pdf');
  }

  // Custom PowerPoint Export using 'pptxgenjs' to create a table
  exportAsPpt() {
    if (this.rowData.length === 0) return;
    const pptx = new PptxGenJS();
    const slide = pptx.addSlide();

    const headers = this.columnDefs.map(col => col.headerName!);
    // For pptxgenjs, the headers must be part of the body data array
    const body = this.rowData.map(row => this.columnDefs.map(col => row[col.field!]));
    const tableData = [headers, ...body];

    slide.addTable(tableData, {
      x: 0.5,
      y: 0.5,
      w: '90%',
      h: '90%',
      colW: [2, 3, 2, 2] // Adjust column widths as needed
    });

    pptx.writeFile({ fileName: 'aggrid-data.pptx' });
  }
}
