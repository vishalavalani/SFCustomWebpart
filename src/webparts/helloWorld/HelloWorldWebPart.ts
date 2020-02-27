import {
  Version,
  DisplayMode,
  Environment,
  EnvironmentType,
  Log
} from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  PropertyPaneSlider
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { escape } from "@microsoft/sp-lodash-subset";
import {
  PropertyPaneContinentSelector,
  IPropertyPaneContinentSelectorProps
} from "../../controls/PropertyPaneContinentSelector";
import styles from "./HelloWorldWebPart.module.scss";
import * as strings from "HelloWorldWebPartStrings";
import MockHttpClient from "./MockHttpClient";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

export interface IHelloWorldWebPartProps {
  description: string;
  test: string;
  test1: boolean;
  test2: string;
  test3: boolean;
  myContinent: string;
  numContinentsVisited: number;
}

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Id: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<
  IHelloWorldWebPartProps
> {
  public render(): void {
    const pageMode: string =
      this.displayMode === DisplayMode.Edit
        ? "You are in edit mode"
        : "You are in read mode";
    const environmentType: string =
      Environment.type === EnvironmentType.Local
        ? "You are in local environment"
        : "You are in SharePoint environment";

    this.context.statusRenderer.displayLoadingIndicator(
      this.domElement,
      "from server. Please wait!"
    );
    setTimeout(() => {
      this.context.statusRenderer.clearLoadingIndicator(this.domElement);

      this._getListData().then(d => {
        console.log("lists", d);
      });
      this.domElement.innerHTML = `
      <div class="${styles.helloWorld}">
    <div class="${styles.container}">
      <div class="${styles.row}">
        <div class="${styles.column}">
          <span class="${styles.title}">Welcome to SharePoint!</span>
  <p class="${
    styles.subTitle
  }">Customize SharePoint experiences using Web Parts.</p>
  <p class="${styles.subTitle}"><strong>Page mode:</strong> ${pageMode}</p>
<p class="${
        styles.subTitle
      }"><strong>Environment:</strong> ${environmentType}</p>
    <p class="${styles.description}">${escape(this.properties.description)}</p>

    <p class="${styles.description}">Continent where I reside: ${escape(
        this.properties.myContinent
      )}</p>
<p class="${styles.description}">Number of continents I've visited: ${
        this.properties.numContinentsVisited
      }</p>

    <p class="${styles.description}">${escape(this.properties.test)}</p>
    <p class="${styles.description}">Loading from ${escape(
        this.context.pageContext.web.title
      )}</p>

    

      <a href="#" class="${styles.button}">
        <span class="${styles.label}">Learn more</span>
          </a>
          </div>
          </div>
          <div id="spListContainer" />
          </div>
          </div>`;

      this._renderListAsync();

      this.domElement
        .getElementsByClassName(`${styles.button}`)[0]
        .addEventListener("click", (event: any) => {
          event.preventDefault();
          alert("Welcome to the SharePoint Framework!");
        });
    }, 1000);

    Log.info("HelloWorld", "message", this.context.serviceScope);
    Log.warn("HelloWorld", "WARNING message", this.context.serviceScope);
    Log.error(
      "HelloWorld",
      new Error("Error message"),
      this.context.serviceScope
    );
    Log.verbose("HelloWorld", "VERBOSE message", this.context.serviceScope);
  }

  private _getMockListData(): Promise<ISPLists> {
    return MockHttpClient.get().then((data: ISPList[]) => {
      var listData: ISPLists = { value: data };
      return listData;
    }) as Promise<ISPLists>;
  }

  private _getListData(): Promise<ISPLists> {
    return this.context.spHttpClient
      .get(
        this.context.pageContext.web.absoluteUrl +
          `/_api/web/lists?$filter=Hidden eq false`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  private _renderList(items: ISPList[]): void {
    let html: string = "";
    items.forEach((item: ISPList) => {
      html += `
  <ul class="${styles.list}">
    <li class="${styles.listItem}">
      <span class="ms-font-l">${item.Title}</span>
    </li>
  </ul>`;
    });

    const listContainer: Element = this.domElement.querySelector(
      "#spListContainer"
    );
    listContainer.innerHTML = html;
  }

  private _renderListAsync(): void {
    // Local environment
    if (Environment.type === EnvironmentType.Local) {
      this._getMockListData().then(response => {
        this._renderList(response.value);
      });
    } else if (
      Environment.type == EnvironmentType.SharePoint ||
      Environment.type == EnvironmentType.ClassicSharePoint
    ) {
      this._getListData().then(response => {
        this._renderList(response.value);
      });
    }
  }

  private validateDescription(value: string): string {
    if (value === null || value.trim().length === 0) {
      return "Provide a description";
    }

    if (value.length > 40) {
      return "Description should not be longer than 40 characters";
    }

    return "";
  }

  private validateContinents(textboxValue: string): string {
    const validContinentOptions: string[] = [
      "africa",
      "antarctica",
      "asia",
      "australia",
      "europe",
      "north america",
      "south america"
    ];
    const inputToValidate: string = textboxValue.toLowerCase();

    return validContinentOptions.indexOf(inputToValidate) === -1
      ? 'Invalid continent entry; valid options are "Africa", "Antarctica", "Asia", "Australia", "Europe", "North America", and "South America"'
      : "";
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  private onContinentSelectionChange(
    propertyPath: string,
    newValue: any
  ): void {
    const oldValue: any = this.properties[propertyPath];
    this.properties[propertyPath] = newValue;
    this.render();
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
                PropertyPaneTextField("description", {
                  label: "Description",
                  onGetErrorMessage: this.validateDescription.bind(this)
                }),
                // PropertyPaneTextField("myContinent", {
                //   label: "My Continent",
                //   onGetErrorMessage: this.validateContinents.bind(this)
                // }),
                new PropertyPaneContinentSelector("myContinent", <
                  IPropertyPaneContinentSelectorProps
                >{
                  label: "Continent where I currently reside",
                  disabled: false,
                  selectedKey: this.properties.myContinent,
                  onPropertyChange: this.onContinentSelectionChange.bind(this)
                }),
                PropertyPaneSlider("numContinentsVisited", {
                  label: "Number of continents I've visited",
                  min: 1,
                  max: 7,
                  showValue: true
                }),
                PropertyPaneTextField("test", {
                  label: "Multi-line Text Field",
                  multiline: true
                }),
                PropertyPaneCheckbox("test1", {
                  text: "Checkbox"
                }),
                PropertyPaneDropdown("test2", {
                  label: "Dropdown",
                  options: [
                    { key: "1", text: "One" },
                    { key: "2", text: "Two" },
                    { key: "3", text: "Three" },
                    { key: "4", text: "Four" }
                  ]
                }),
                PropertyPaneToggle("test3", {
                  label: "Toggle",
                  onText: "On",
                  offText: "Off"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
