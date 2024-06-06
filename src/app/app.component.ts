import { Component } from '@angular/core';
import { ExcelService } from './excel.service';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  constructor(private excelService: ExcelService) { }

  data: any[][] = [];
  jsonData: any[] = [];

  async onFileChange(event: any) {
    const file = event.target.files[0];
    if (file) {
      try {
        this.data = await this.excelService.readExcel(file);
        console.log(this.data);

        if (this.data.length > 0) {
          const keys = this.data[0].map(key => key.toLowerCase());
          this.jsonData = this.data.slice(1).map(row => {
            let obj:any = {};
            row.forEach((value:any, index:number) => {
              obj[keys[index]] = value;
            });
            return obj;
          });
        }
        console.log('JSON data:', JSON.stringify(this.jsonData, null, 2));

      } catch (error) {
        console.error('Error reading Excel file:', error);
      }
    }
  }

  // Export Array to Excel
  exportArrayToExcel() {
    if (this.data.length > 0) {
      this.excelService.exportToExcel(this.data, 'exported-data');
    } else {
      console.error('No data available to export');
    }
  }
  // Export JSON to Excel
  exportJsonToExcel() {
    if (this.jsonData.length > 0) {
      this.excelService.exportAsExcelFile(this.jsonData,"Exported File",(success: boolean) => {
          if (success) {
            console.log("Exported Successful");
          }
        });
    } else {
      console.error('No data available to export');
    }
  }

}
