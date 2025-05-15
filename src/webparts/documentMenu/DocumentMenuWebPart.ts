import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import * as strings from "DocumentMenuWebPartStrings";
import DocumentMenu from "./components/DocumentMenu";
import { IDocumentMenuProps } from "./interfaces/IDocumentMenuProps";
import { DocumentMenuService } from "./services/DocumentMenuService";

export interface IDocumentMenuWebPartProps {
  description: string;
  documentUrl: string;
  layoutDropdown: string;
}

export default class DocumentMenuWebPart extends BaseClientSideWebPart<IDocumentMenuWebPartProps> {
  private _service: DocumentMenuService;

  public render(): void {
    const element: React.ReactElement<IDocumentMenuProps> = React.createElement(
      DocumentMenu,
      {
        context: this.context,
        description: this.properties.description,
        userDisplayName: this.context.pageContext.user.displayName,
        documentUrl: this.properties.documentUrl,
        layoutDropdown: this.properties.layoutDropdown,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  public async onInit(): Promise<void> {
    this._service = new DocumentMenuService(this.context);
    console.log(this._service);
    return Promise.resolve();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const layoutDropdownOptions: IPropertyPaneDropdownOption[] = [
      { key: "1", text: "Icon with Documents" },
      { key: "2", text: "Tiles" },
    ];

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
                // PropertyPaneTextField("description", {
                //   label: strings.DescriptionFieldLabel,
                // }),
                PropertyPaneTextField("documentUrl", {
                  // Add this block
                  label: "Document Menu URL",
                  value: this.properties.documentUrl || "",
                }),
                PropertyPaneDropdown("layouyDropdown", {
                  label: "Select an Layout",
                  options: layoutDropdownOptions,
                  selectedKey: this.properties.layoutDropdown || "1",
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
