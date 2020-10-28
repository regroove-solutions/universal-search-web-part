import * as React from "react";
import * as ReactDom from "react-dom";
import {
  IPropertyPaneField,
  PropertyPaneFieldType
} from "@microsoft/sp-property-pane";
import { IPropertyPaneCustomFieldProps } from "@microsoft/sp-property-pane";
import styles from "./PropertyPaneLogo.module.scss";

export interface IPropertyPaneLogoProps {
  svg: JSX.Element;
  href: string;
}

export class PropertyPaneLogo
  implements IPropertyPaneField<IPropertyPaneLogoProps> {
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public properties: IPropertyPaneLogoProps & IPropertyPaneCustomFieldProps;
  private elem: HTMLElement;
  public targetProperty: string;

  constructor(properties: IPropertyPaneLogoProps) {
    this.properties = {
      key: "property-pane-logo",
      svg: properties.svg,
      onRender: this.onRender.bind(this),
      href: properties.href
    };
  }

  public render() {
    if (!this.elem) {
      return;
    }
    this.onRender(this.elem);
  }

  private onRender(elem: HTMLElement) {
    if (!this.elem) {
      this.elem = elem;
    }

    const element = (
      <>
        <div style={{ marginBottom: "64px" }} />
        <div className={styles.propertyPaneLogo}>
          <a
            className={styles.svgBox}
            href={this.properties.href}
            target="_blank"
            title="Powered by Navo"
          >
            {this.properties.svg}
          </a>
        </div>
      </>
    );

    ReactDom.render(element, elem);
  }
}
