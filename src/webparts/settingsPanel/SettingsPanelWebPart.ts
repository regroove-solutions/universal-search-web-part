import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { IPropertyPaneConfiguration } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { SettingsPanel } from "./components/SettingsPanel";
import { Settings } from "../../shared/types";
import { webPartKey } from "../../shared/constants";
import { PropertiesService } from "../../shared/propertiesService";

export default class SettingsPanelWebPart extends BaseClientSideWebPart<{}> {
  public onInit(): Promise<void> {
    if (this.context.sdks.microsoftTeams) {
      this.context.domElement.setAttribute(
        "data-theme",
        this.context.sdks.microsoftTeams.context.theme || "default"
      );
    }
    return Promise.resolve();
  }

  private async _updateSettings(webPartProps: Settings): Promise<void> {
    try {
      await new PropertiesService<Settings>(this.context).setProperties(
        webPartKey,
        webPartProps
      );
    } catch (err) {
      this.renderError(err);
    }
  }

  private async _getStoredSettings(): Promise<Settings> {
    try {
      return new PropertiesService<Settings>(this.context).getProperties(
        webPartKey
      );
    } catch (err) {
      this.renderError(err);
    }
  }

  public render(): void {
    ReactDom.render(
      React.createElement(SettingsPanel, {
        getSettings: this._getStoredSettings.bind(this),
        updateSettings: this._updateSettings.bind(this)
      }),
      this.domElement
    );
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  //@ts-ignore
  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "Settings panel for Teams tab"
          },
          groups: []
        }
      ]
    };
  }
}
