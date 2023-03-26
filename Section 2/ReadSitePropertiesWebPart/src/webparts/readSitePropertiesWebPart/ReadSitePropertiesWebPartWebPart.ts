import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { escape } from "@microsoft/sp-lodash-subset";

import styles from "./ReadSitePropertiesWebPartWebPart.module.scss";
import * as strings from "ReadSitePropertiesWebPartWebPartStrings";

import {
  Environment,
  EnvironmentType,
  DisplayMode,
} from "@microsoft/sp-core-library";

export interface IReadSitePropertiesWebPartWebPartProps {
  description: string;
  environmenttitle: string;
  displaymode: string;
}

export default class ReadSitePropertiesWebPartWebPart extends BaseClientSideWebPart<IReadSitePropertiesWebPartWebPartProps> {
  private _findOutDisplayMode(): void {
    if (this.displayMode === DisplayMode.Edit) {
      this.properties.displaymode = "Page is in Edit Mode";
    } else {
      this.properties.displaymode = "Page is in Read Mode";
    }
  }

  private _findOutEnvironment(): void {
    // Local environment
    if (Environment.type === EnvironmentType.Local) {
      this.properties.environmenttitle = "Local SharePoint Environment";
    } else if (
      Environment.type === EnvironmentType.SharePoint ||
      Environment.type === EnvironmentType.ClassicSharePoint
    ) {
      this.properties.environmenttitle = "Online SharePoint Environment";
    }
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.readSitePropertiesWebPart}">
    <div class="${styles.container}">
      <div class="${styles.row}">
        <div class="${styles.column}">
          <span class="${styles.title}">Welcome to SharePoint!</span>
  <p class="${
    styles.subTitle
  }">Customize SharePoint experiences using Web Parts.</p>
    <p class="${styles.description}">${escape(this.properties.description)}</p>

    <p class="${styles.description}">Absolute URL ${escape(
      this.context.pageContext.web.absoluteUrl
    )}</p>
    <p class="${styles.description}">Title ${escape(
      this.context.pageContext.web.title
    )}</p>
    <p class="${styles.description}">Relative URL ${escape(
      this.context.pageContext.web.serverRelativeUrl
    )}</p>
    <p class="${styles.description}">User Name ${escape(
      this.context.pageContext.user.displayName
    )}</p>
      <a href="https://aka.ms/spfx" class="${styles.button}">
        <span class="${styles.label}">Learn more</span>
          </a>
          </div>
          </div>
          </div>
          </div>`;

    this._findOutEnvironment();
    this._findOutDisplayMode();
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
