import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './Amtest1WebPart.module.scss';
import * as strings from 'Amtest1WebPartStrings';
import * as xlsx from 'xlsx';

export interface IAmtest1WebPartProps {
  description: string;
}

export default class Amtest1WebPart extends BaseClientSideWebPart<IAmtest1WebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.amtest1 }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Testing Export to Excel</span>
              <p class="${ styles.subTitle }">Click the button below to export the description text</p>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>
              <a href="javascript:exportToExcel();" class="${ styles.button }">
                <span id="exportButton" class="${ styles.label }">Export to Excel</span>
              </a>
            </div>
          </div>
        </div>
      </div>`;
      let clickEvent= document.getElementById("exportButton");
      clickEvent.addEventListener("click", (e: Event) => this.exportToExcel());
  }

  private exportToExcel(): void {
    //--- Generate Workbook ---
    let wb: xlsx.WorkBook = xlsx.utils.book_new();
    let sheetData: any = [ { Description:this.properties.description } ];
    let wsSheet:any = xlsx.utils.json_to_sheet(sheetData);
    xlsx.utils.book_append_sheet(wb, wsSheet, 'Sample Data');
    // --- Download Excel File ---
    const dateValue: string = new Date().toUTCString();
    let filename: string = 'AutomationMatrixTest (' + dateValue + ').xlsx'; 
    xlsx.writeFile(wb, filename, { compression:true });
}

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
