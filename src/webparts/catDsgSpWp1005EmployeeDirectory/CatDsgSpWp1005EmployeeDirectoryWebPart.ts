import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import styles from './CatDsgSpWp1005EmployeeDirectoryWebPart.module.scss';
import * as strings from 'CatDsgSpWp1005EmployeeDirectoryWebPartStrings';
import JQueryLoader from './JQueryLoader';
import Utils from './CatDsgWp1005EmployeeDirectoryWebPartUtils';

export interface ICatDsgSpWp1005EmployeeDirectoryWebPartProps {
  description: string;
  searchResultPageRelativeUrl: string;
  queryTextParameterName: string;
}

export default class CatDsgSpWp1005EmployeeDirectoryWebPart extends BaseClientSideWebPart<ICatDsgSpWp1005EmployeeDirectoryWebPartProps> {

  public render(): void {
    this.properties.description = strings.CatDsgSpWp1005EmployeeDirectoryDescription;
    this.domElement.innerHTML = `
      <div class="${ styles.catDsgSpWp1005EmployeeDirectory}">
        <div class="${ styles.container}">
          <div class="${ styles.row}">
            <div class="${ styles.column}">  
              <div class="${styles.catDsgSpWp1005EmployeeDirectorySearchBox}">
                  <input class="${styles.catDsgSpWp1005EmployeeDirectorySearchInput}" type="text" name="search" placeholder="${strings.CatDsgSpWp1005EmployeeDirectorySearchBoxInputPlaceHolderString}"/>
                  <button class="${styles.catDsgSpWp1005EmployeeDirectorySearchButton}">
                    <svg xmlns="http://www.w3.org/2000/svg" class="${styles.catDsgSpWp1005EmployeeDirectorySearchIcon}" aria-hidden="true" viewBox="0 0 32 32" focusable="false" width="32" height="32"><path d="M 20.992 0 c 1.024 0 1.984 0.128 2.944 0.384 s 1.792 0.64 2.624 1.088 S 28.096 2.56 28.8 3.2 c 0.64 0.704 1.216 1.408 1.728 2.24 s 0.832 1.664 1.088 2.624 s 0.384 1.92 0.384 2.944 s -0.128 1.984 -0.384 2.944 s -0.64 1.792 -1.088 2.624 s -1.088 1.536 -1.728 2.176 c -0.704 0.64 -1.408 1.216 -2.24 1.728 s -1.664 0.832 -2.624 1.088 s -1.92 0.384 -2.944 0.384 c -1.28 0 -2.56 -0.192 -3.712 -0.64 a 10.46 10.46 0 0 1 -3.264 -1.92 L 1.728 31.68 c -0.192 0.192 -0.448 0.32 -0.704 0.32 s -0.512 -0.128 -0.704 -0.32 s -0.32 -0.384 -0.32 -0.704 s 0.128 -0.512 0.32 -0.704 l 12.224 -12.288 a 12.145 12.145 0 0 1 -1.92 -3.264 c -0.448 -1.216 -0.64 -2.432 -0.64 -3.712 c 0 -1.024 0.128 -1.984 0.384 -2.944 s 0.64 -1.792 1.088 -2.624 s 1.024 -1.536 1.728 -2.24 c 0.704 -0.64 1.408 -1.216 2.24 -1.728 S 17.088 0.64 18.048 0.384 C 19.008 0.128 19.968 0 20.992 0 Z m 0 19.968 c 1.216 0 2.432 -0.256 3.52 -0.704 s 2.048 -1.088 2.88 -1.92 s 1.472 -1.792 1.92 -2.88 s 0.704 -2.24 0.704 -3.52 s -0.256 -2.432 -0.704 -3.52 s -1.088 -2.048 -1.92 -2.88 S 25.6 3.2 24.512 2.688 c -1.088 -0.448 -2.24 -0.704 -3.52 -0.704 s -2.432 0.256 -3.52 0.704 s -2.048 1.088 -2.88 1.92 s -1.472 1.792 -1.92 2.88 s -0.704 2.24 -0.704 3.52 s 0.256 2.432 0.704 3.52 s 1.088 2.048 1.92 2.88 s 1.792 1.472 2.88 1.92 c 1.152 0.448 2.304 0.64 3.52 0.64 Z" /></svg>
                  </button>
              </div>
            </div>
          </div>
        </div>
      </div>`;
    let siteAbsoluteUrl = this.context.pageContext.site.absoluteUrl;
    let tenantAbsoluteUrl=Utils.getTenantUrl(this.context.pageContext.site.absoluteUrl,this.context.pageContext.site.serverRelativeUrl);
    let resultPageRelativeUrl = this.properties.searchResultPageRelativeUrl?this.properties.searchResultPageRelativeUrl:'';
    let resultPageAbsoluteURl='';
    if(this.properties.searchResultPageRelativeUrl!='') {
      if (resultPageRelativeUrl.indexOf('/') != 0) {
        resultPageRelativeUrl = '/' + resultPageRelativeUrl;
      }
      resultPageAbsoluteURl=siteAbsoluteUrl+resultPageRelativeUrl;
    }
    else {
      resultPageAbsoluteURl=tenantAbsoluteUrl+'/search/Pages/peopleresults.aspx';
    }
    let queryTextParameterName = this.properties.queryTextParameterName;
    JQueryLoader.LoadDependencies("https://ajax.aspnetcdn.com/ajax/jQuery/jquery-1.11.1.min.js", []).then((object) => {
      $("." + styles.catDsgSpWp1005EmployeeDirectorySearchBox + ">" + "."+styles.catDsgSpWp1005EmployeeDirectorySearchButton).click(function (event) {
        var searchValue = $(this).parent().find("input").val();
        window.location.href = resultPageAbsoluteURl + '?' + queryTextParameterName + '=' + searchValue;
        event.preventDefault();
      });
      $("." + styles.catDsgSpWp1005EmployeeDirectorySearchBox + ">input").keydown(function (event) {
        if (event.keyCode == 13) {
          var searchValue = $(this).val();
          window.location.href = resultPageAbsoluteURl + '?' + queryTextParameterName + '=' + searchValue;
          event.preventDefault();
        }
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
                }),
                PropertyPaneTextField('queryTextParameterName', {
                  label: strings.CatDsgSpWp1005EmployeeDirectoryFieldLabelQueryTextParameterName,
                  onGetErrorMessage: this.validateRequiredProperty.bind(this)
                })
              ]
            }
          ]
        }
      ]
    };
  }
}


