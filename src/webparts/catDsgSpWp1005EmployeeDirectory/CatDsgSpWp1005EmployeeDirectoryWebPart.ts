import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './CatDsgSpWp1005EmployeeDirectoryWebPart.module.scss';
import * as strings from 'CatDsgSpWp1005EmployeeDirectoryWebPartStrings';
import { SPComponentLoader } from "@microsoft/sp-loader";
import JQueryLoader from './JQueryLoader';
 
export interface ICatDsgSpWp1005EmployeeDirectoryWebPartProps {
  description: string;
  searchResultPageRelativeUrl:string;
}

export default class CatDsgSpWp1005EmployeeDirectoryWebPart extends BaseClientSideWebPart<ICatDsgSpWp1005EmployeeDirectoryWebPartProps> {

  public render(): void {
    this.properties.description=strings.CatDsgSpWp1005EmployeeDirectoryDescription;
    SPComponentLoader.loadCss("https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.0/css/fabric.min.css");
    this.domElement.innerHTML = `
      <div class="${ styles.catDsgSpWp1005EmployeeDirectory }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">  
              <div class="${styles.catDsgSpWp1005EmployeeDirectorySearchBox}">
                <form>
                  <input class="${styles.catDsgSpWp1005EmployeeDirectorySearchInput}" type="text" name="search" placeholder="${strings.CatDsgSpWp1005EmployeeDirectorySearchBoxInputPlaceHolderString}"/>
                  <button class="${styles.catDsgSpWp1005EmployeeDirectorySearchButton}">
                    <i class="ms-Icon ms-Icon--Search" role="presentation" aria-hidden="true" data-icon-name="Search"></i>
                  </button>
                </form>
              </div>
            </div>
          </div>
        </div>
      </div>`;
      let webAbsoluteUrl=this.context.pageContext.web.absoluteUrl;
      let resultPagerelativeUrl=this.properties.searchResultPageRelativeUrl;
      if(resultPagerelativeUrl.indexOf('/')!=0) {
        resultPagerelativeUrl='/'+resultPagerelativeUrl;
      }
      JQueryLoader.LoadDependencies("https://ajax.aspnetcdn.com/ajax/jQuery/jquery-1.11.1.min.js", []).then((object) => {
      $("."+styles.catDsgSpWp1005EmployeeDirectorySearchBox+">form").submit(function (event) {
        var searchValue = $(this).find("input").val();
        window.location.href = webAbsoluteUrl + resultPagerelativeUrl+'?k=' + searchValue;
        event.preventDefault();
    });
  });
  }

  private validateRequiredProperty(value: string): string {
    if (value === null || (value != null && value.trim().length === 0)) {
      return strings.CatDsgSpWp1005EmployeeDirectoryRequiredPropertyMessage;
    }
    return "";
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
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
                }),
                PropertyPaneTextField('searchResultPageRelativeUrl', {
                  label: strings.CatDsgSpWp1005EmployeeDirectoryFieldLabelSearchResultPageRelativeUrl,
                  onGetErrorMessage:this.validateRequiredProperty.bind(this)
                }), 
              ]
            }
          ]
        }
      ]
    };
  }
}
