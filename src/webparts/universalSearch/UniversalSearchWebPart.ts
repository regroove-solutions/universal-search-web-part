import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneSlider,
  PropertyPaneToggle,
  PropertyPaneTextField
} from "@microsoft/sp-property-pane";
import { PropertyPaneLogo } from "../../controls/PropertyPaneLogo/PropertyPaneLogo";
import { PropertyFieldMultiSelect } from "@pnp/spfx-property-controls";
import { PropertyFieldOrder } from "@pnp/spfx-property-controls/lib/propertyFields/order";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { SearchTarget, Settings } from "../../shared/types";
import {
  defaultSearchTargets,
  defaultSettings,
  webPartKey
} from "../../shared/constants";
import { NavoLogoSVG } from "../../shared/svg";
import { PropertiesService } from "../../shared/propertiesService";
import * as strings from "UniversalSearchWebPartStrings";
import {
  UniversalSearch,
  UniversalSearchTeams,
  IUniversalSearchProps
} from "./components/UniversalSearch";

export interface IUniversalSearchWebPartProps extends IUniversalSearchProps {
  searchTargetKeys: string[];
  customSearchUrl: string;
  customSearchName: string;
}

export default class UniversalSearchWebPart extends BaseClientSideWebPart<
  IUniversalSearchWebPartProps
> {
  private async _getTeamsPersonalAppSettings() {
    let settings: Settings;
    try {
      settings = {
        ...defaultSettings,
        ...(await new PropertiesService<Settings>(this.context).getProperties(
          webPartKey
        ))
      };
    } catch (e) {
      settings = defaultSettings;
    }
    Object.keys(settings).forEach(
      (key) => (this.properties[key] = settings[key])
    );
  }

  public async onInit(): Promise<void> {
    const tenancy = this.context.pageContext.site.absoluteUrl.split(/[.\/]/)[2];
    defaultSearchTargets.sharepoint.baseUrl = `https://${tenancy}.sharepoint.com/search/Pages/results.aspx?k=`;
    if (this.context.sdks.microsoftTeams) {
      this.context.domElement.setAttribute(
        "data-theme",
        this.context.sdks.microsoftTeams.context.theme || "default"
      );
    }
    if (
      this.context.sdks.microsoftTeams &&
      !this.context.sdks.microsoftTeams.context.teamName
    )
      await this._getTeamsPersonalAppSettings();
    this.properties.instanceId = this.instanceId;

    return Promise.resolve();
  }

  public render(): void {
    this.properties.searchTargets = this.properties.searchTargetKeys.map(
      (key) => defaultSearchTargets[key]
    );

    if (this.properties.searchTargetKeys.indexOf("custom") > -1) {
      defaultSearchTargets.custom.baseUrl = this.properties.customSearchUrl;
      defaultSearchTargets.custom.text = this.properties.customSearchName
        ? this.properties.customSearchName
        : "Custom";
    }

    if (!this.properties.customSearchUrl)
      this.properties.searchTargets = this.properties.searchTargets.filter(
        (target) => target.key !== "custom"
      );

    const element: React.ReactElement<IUniversalSearchProps> = React.createElement(
      this.context.sdks.microsoftTeams ? UniversalSearchTeams : UniversalSearch,
      this.properties
    );
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  //@ts-ignore
  protected get dataVersion(): Version {
    return Version.parse(this.manifest.version);
  }

  private _handleOrder(
    path: string,
    oldValue: string[],
    value: SearchTarget[]
  ): void {
    this.properties.searchTargetKeys = value.map((target) => target.key);
  }

  private _validateUrlField(value: string): string {
    if (
      !value ||
      this.properties.searchTargetKeys.indexOf("custom") === -1 ||
      value.match(/http(s)?:\/\/[\w.]+\.\w+/)
    )
      return "";
    return "Enter a search URL, e.g. 'https://google.ca/search?q='";
  }

  private _validateColourField(value: string): string {
    if (!value || value.match(/#([A-F0-9]{3}){1,2}$/i)) return "";
    return "Enter a valid hex colour code, e.g. '#F0F0F0'";
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: `${strings.PropertyPaneDescription} (${this.manifest.version})`
          },
          groups: [
            {
              groupName: "Search Options",
              groupFields: [
                PropertyPaneToggle("newTab", {
                  label: "Open results in new tab",
                  disabled: !!this.context.sdks.microsoftTeams
                }),
                PropertyPaneToggle("usePreference", {
                  label: "Default search",
                  onText: "Save user preference",
                  offText: `Use ${
                    this.properties.searchTargets[0]
                      ? this.properties.searchTargets[0].text
                      : "default"
                  }`
                }),
                PropertyFieldMultiSelect("searchTargetKeys", {
                  label: "Search targets",
                  options: (<any>Object).values(defaultSearchTargets),
                  key: "multiSelect",
                  selectedKeys: this.properties.searchTargetKeys
                }),
                PropertyFieldOrder("searchTargets", {
                  label: "",
                  items: this.properties.searchTargets,
                  textProperty: "text",
                  key: "order",
                  onPropertyChange: this._handleOrder,
                  properties: this.properties
                }),
                PropertyPaneTextField("customSearchName", {
                  label: "Custom search title"
                }),
                PropertyPaneTextField("customSearchUrl", {
                  label: "Custom search URL",
                  placeholder: "https://",
                  onGetErrorMessage: this._validateUrlField.bind(this),
                  validateOnFocusOut: true,
                  deferredValidationTime: 800
                }),
                new PropertyPaneLogo({
                  href: "https://getnavo.com",
                  svg: NavoLogoSVG
                })
              ]
            }
          ]
        },
        {
          header: {
            description: `${strings.PropertyPaneDescription} (${this.manifest.version})`
          },
          groups: [
            {
              groupName: "Style Options",
              groupFields: [
                PropertyPaneSlider("boxWidth", {
                  label: "Search box width (%)",
                  min: 32,
                  max: 100,
                  step: 4
                }),
                PropertyPaneSlider("boxHeight", {
                  label: "Search box height (px)",
                  min: 28,
                  max: 60,
                  step: 4,
                  value: 32
                }),
                PropertyPaneSlider("borderWidth", {
                  label: "Border width",
                  min: 0,
                  max: 8
                }),
                PropertyPaneTextField("themeColour", {
                  label: "Colour (leave blank for default)",
                  placeholder: "#FFFFFF",
                  maxLength: 7,
                  onGetErrorMessage: this._validateColourField.bind(this)
                }),
                PropertyPaneToggle("showLogo", {
                  label: "Show Navo logo"
                }),
                new PropertyPaneLogo({
                  href: "https://getnavo.com",
                  svg: NavoLogoSVG
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
