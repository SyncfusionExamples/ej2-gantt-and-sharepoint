import { ServiceScope, Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneButton,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

// import Essential JS 2 Gantt
import { Gantt, Sort, Edit } from '@syncfusion/ej2-gantt';

// add Syncfusion Essential JS 2 style reference from node_modules
require('../../../node_modules/@syncfusion/ej2/fluent.css');

import styles from './GanttChartWebPart.module.scss';
import * as strings from 'GanttChartWebPartStrings';

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { gantt } from '@syncfusion/ej2';
import { ActionCompleteArgs, TaskFieldsModel } from '@syncfusion/ej2/gantt';
import { DataManager, ODataAdaptor, Query, WebApiAdaptor } from '@syncfusion/ej2/data';
import { DateTime } from '@syncfusion/ej2/charts';
import { Ajax, BeforeSendEventArgs } from '@syncfusion/ej2/base';

export interface IGanttChartWebPartProps {
  description: string;
  dataSource: object[];
  allowSorting: boolean;
  allowEditing: boolean;
  selectedList: string;
  taskID: string | null,
  taskName: string | null,
  startDate: string | null,
  duration: string | null,
  progress: string | null,
  parentID: string | null

}

export default class GanttChartWebPart extends BaseClientSideWebPart<IGanttChartWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  private _siteLists: string[];
  private ganttInstance: Gantt;
  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.ganttChart} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
            <div id="Gantt-${this.instanceId}"> </div>
    </section>`;

    
  Gantt.Inject(Sort, Edit);
  this.ganttInstance = new Gantt({
       dataSource: this.properties.dataSource,
       allowSorting: this.properties.allowSorting,
       editSettings: {allowEditing: this.properties.allowEditing, allowTaskbarEditing: this.properties.allowEditing},
       taskFields: {
          id: this.properties.taskID,
          name: this.properties.taskName,
          startDate: this.properties.startDate,
          duration: this.properties.duration,
          progress: this.properties.progress,
          parentID: this.properties.parentID
      },
      actionComplete: (args: ActionCompleteArgs) => {
        if (args.requestType == 'save'){
          for (var i = 0; i < args.modifiedTaskData.length; i++) {
            var url = this.context.pageContext.web.absoluteUrl + '/_api/web/lists/' + this.properties.selectedList + 'List/items(' + args.modifiedTaskData[i]['Id'] + ')';
            this.context.spHttpClient.post(url, SPHttpClient.configurations.v1,  {
              headers: {
                'Accept': 'application/json;odata=nometadata',
                'Content-type': 'application/json;odata=nometadata',
                'odata-version': '',
                'IF-MATCH': '*',
                'X-HTTP-Method': 'MERGE'
              },
              body:  JSON.stringify(args.modifiedTaskData[i])
            })
          }

        }
      }
   });

  this.ganttInstance.appendTo('#Gantt-'+this.instanceId);
  }
  protected async onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();
    this._siteLists = await this._getSiteLists();
    this.properties.allowSorting = true;
    this.properties.dataSource = [];
    this.properties.taskID = 'TaskID'; this.properties.taskName = 'TaskName'; this.properties.startDate = 'StartDate';
    this.properties.duration = 'Duration'; this.properties.progress = 'Progress'; this.properties.parentID = 'ParentID';
    return super.onInit();
  }

  public async _getSiteLists(): Promise<string[]> {
    const endpoint: string = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$select=Title&$filter=Hidden eq false&$orderby=Title&$top=10`;

    const rawResponse: SPHttpClientResponse = await this.context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1);
  
    return (await rawResponse.json()).value.map(
      (list: {Title: string}) => {
        return list.Title;
      }
    );
  }
  public async _getSiteListItems(): Promise<void> {
    const endpoint: string = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/${this.properties.selectedList}List/items`;

    const rawResponse: SPHttpClientResponse = await this.context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1);
      var response = await rawResponse.json();
    this.properties.dataSource = response.value.map(
      (list: {Id: number, TaskID: number, TaskName: string, StartDate: DateTime, Duration: number, Progress: number, ParentID: number}) => {
        return {Id: list.Id, TaskID: list.TaskID, TaskName: list.TaskName, StartDate: list.StartDate, Duration: list.Duration, Progress: list.Progress, ParentID: list.ParentID}
      }
    );
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

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
                }),
                PropertyPaneCheckbox('allowEditing', {
                  checked: this.properties.allowEditing,
                  text: "Editing in Gantt"
                 }),
                PropertyPaneCheckbox('allowSorting', {
                 checked: this.properties.allowSorting,
                 text: "Sorting in Gantt"
                }),
                PropertyPaneDropdown('selectedList', {
                  label: 'Connect to List source',                  
                  options:  this._siteLists.map((list: string) => {
                    return <IPropertyPaneDropdownOption>{
                      key: list, text: list
                    }
                  }),                  
                  
                })
              ]
            },
            {
              groupName: "Gantt Task Field Mapping",
              groupFields: [
                PropertyPaneTextField('taskID', {
                  value: this.properties.taskID,
                  label: "Task ID"
                }),
                PropertyPaneTextField('taskName', {
                  value: this.properties.taskName,
                  label: "Task Name"
                }),
                PropertyPaneTextField('startDate', {
                  value: this.properties.startDate,
                  label: "Start Date"
                }),
                PropertyPaneTextField('duration', {
                  value: this.properties.duration,
                  label: "Duration"
                }),
                PropertyPaneTextField('progress', {
                  value: this.properties.progress,
                  label: "Progress"
                }),
                PropertyPaneTextField('parentID', {
                  value: this.properties.parentID,
                  label: "Parent ID"
                }),
                PropertyPaneButton("", {
                  text: "Apply",
                  onClick: () => {
                     this._getSiteListItems();
                  }
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
