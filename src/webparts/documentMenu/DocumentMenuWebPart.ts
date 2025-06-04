import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  // PropertyPaneButton,
  // PropertyPaneButtonType,
} from "@microsoft/sp-property-pane";
import { PropertyFieldIconPicker } from "@pnp/spfx-property-controls/lib/PropertyFieldIconPicker";
import {
  PropertyFieldColorPicker,
  PropertyFieldColorPickerStyle,
} from "@pnp/spfx-property-controls/lib/PropertyFieldColorPicker";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import * as strings from "DocumentMenuWebPartStrings";
import DocumentMenu from "./components/DocumentMenu";
import { IDocumentMenuProps } from "./interfaces/IDocumentMenuProps";
import { DocumentMenuService } from "./services/DocumentMenuService";

export interface IDocumentMenuWebPartProps {
  description: string;
  documentLibraryUrl: string;
  layoutDropdownValue: any;
  siteCollectionUrl: string;
  itemsRowCount: string;
  heightSliderValue?: number;
  widthSliderValue?: number;
  itemIcons: { [key: string]: string };
  itemColors: { [key: string]: string };
  categoryDropdownValue: string;
}

export default class DocumentMenuWebPart extends BaseClientSideWebPart<IDocumentMenuWebPartProps> {
  private _service: DocumentMenuService;
  private _currentItems: any[] = [];
  private _libraryOptions: (IPropertyPaneDropdownOption & {
    serverRelativeUrl: string;
  })[] = [];
  private _categoryDropdownOptions: IPropertyPaneDropdownOption[] = [];
  private _activeIconLayout: string = "icon";
  private _navigationStackLength: number = 0;

  public render(): void {
    const element: React.ReactElement<IDocumentMenuProps> = React.createElement(
      DocumentMenu,
      {
        // This function is called when the current items change
        onCurrentItemsChange: (items) => {
          this._currentItems = items;
          if (this.context.propertyPane.isPropertyPaneOpen()) {
            this.context.propertyPane.refresh();
          }
        },
        // This function is called when the active icon layout change
        onLayoutStateChange: (activeIconLayout, navigationStackLength) => {
          this._activeIconLayout = activeIconLayout;
          this._navigationStackLength = navigationStackLength;
          if (this.context.propertyPane.isPropertyPaneOpen()) {
            this.context.propertyPane.refresh();
          }
        },
        context: this.context,
        description: this.properties.description,
        userDisplayName: this.context.pageContext.user.displayName,
        documentLibraryUrl: this.properties.documentLibraryUrl,
        layoutDropdownValue: this.properties.layoutDropdownValue,
        itemIcons: this.properties.itemIcons,
        siteCollectionUrl: this.properties.siteCollectionUrl,
        itemsRowCount: this.properties.itemsRowCount,
        heightSliderValue: this.properties.heightSliderValue,
        widthSliderValue: this.properties.widthSliderValue,
        itemColors: this.properties.itemColors,
        categoryDropdownValue: this.properties.categoryDropdownValue,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  public async onInit(): Promise<void> {
    this._service = new DocumentMenuService(
      this.context,
      this.properties.siteCollectionUrl
    );
    console.log(this._service);

    if (this.properties.siteCollectionUrl) {
      this._service
        .getLibraryOptions(this.properties.siteCollectionUrl)
        .then((libs: any[]) => {
          this._libraryOptions = libs.map((lib: any) => ({
            key: lib.Id,
            text: lib.Title,
            serverRelativeUrl: lib.ServerRelativeUrl,
          }));
          // Refresh property pane to update dropdown options
          if (this.context.propertyPane.isPropertyPaneOpen()) {
            this.context.propertyPane.refresh();
          }
        });
    }

    if (this.properties.documentLibraryUrl) {
      this._categoryDropdownOptions = await this._service.getFieldsForUrl(
        this.properties.documentLibraryUrl
      );
    }

    return Promise.resolve();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: any,
    newValue: any
  ): void {
    if (propertyPath === "siteCollectionUrl") {
      this.properties.siteCollectionUrl = newValue;
      // this.properties.documentLibraryUrl = "";

      if (newValue === "") {
        this._libraryOptions = [];
      } else {
        this._service.getLibraryOptions(newValue).then((libs: any[]) => {
          this._libraryOptions = libs.map((lib: any) => ({
            key: lib.Id,
            text: lib.Title,
            serverRelativeUrl: lib.ServerRelativeUrl,
          }));
          // Refresh property pane to update dropdown options
          if (this.context.propertyPane.isPropertyPaneOpen()) {
            this.context.propertyPane.refresh();
          }
        });
      }
    }

    if (propertyPath === "layoutDropdown") {
      this.properties.layoutDropdownValue = newValue;
    }

    if (propertyPath === "documentLibrary") {
      const selectedLibrary = this._libraryOptions.find(
        (lib) => lib.key === newValue
      );
      if (selectedLibrary) {
        this.properties.documentLibraryUrl = selectedLibrary.serverRelativeUrl;
      }
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const layoutDropdownOptions: IPropertyPaneDropdownOption[] = [
      { key: "1", text: "Icon with Documents" },
      { key: "2", text: "Tiles" },
    ];

    //Define item icon fields based on the current layout
    const itemIconFields =
      this._currentItems &&
      this._currentItems.length > 0 &&
      this._activeIconLayout === "icon" &&
      this._navigationStackLength === 1 &&
      this.properties.layoutDropdownValue === "1"
        ? this._currentItems
            .filter((item) => item.folder)
            .map((item) =>
              PropertyFieldIconPicker(`itemIcons_${item.Name}`, {
                currentIcon: this.properties.itemIcons?.[item.Name] || "Folder",
                key: `icon_${item.Name}`,
                onSave: (icon: string) => {
                  this.properties.itemIcons = {
                    ...this.properties.itemIcons,
                    [item.Name]: icon,
                  };
                  this.render();
                },
                buttonLabel: `Icon for ${item.Name}`,
                renderOption: "panel",
                properties: this.properties,
                onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                label: `Pick icon for ${item.Name}`,
              })
            )
        : [];

    // Define item colors for the icon layout
    const itemIconColors =
      this._currentItems &&
      this._currentItems.length > 0 &&
      this._activeIconLayout === "icon" &&
      this._navigationStackLength === 1 &&
      this.properties.layoutDropdownValue === "1"
        ? this._currentItems
            .filter((item) => item.folder)
            .map((item) =>
              PropertyFieldColorPicker(`itemColors_${item.Name}`, {
                label: `Pick color for ${item.Name}`,
                key: `color_${item.Name}`,
                selectedColor:
                  this.properties.itemColors?.[item.Name] || "#ffffff",
                onPropertyChange: (propertyPath, oldValue, newValue) => {
                  this.properties.itemColors = {
                    ...this.properties.itemColors,
                    [item.Name]: newValue,
                  };
                  this.render();
                },
                properties: this.properties,
                style: PropertyFieldColorPickerStyle.Full,
              })
            )
        : [];

    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              // groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("siteCollectionUrl", {
                  // Add this block
                  label: "Site Collection URL",
                  value: this.properties.siteCollectionUrl,
                }),
                PropertyPaneDropdown("documentLibrary", {
                  label: "Select a Document Library",
                  options: this._libraryOptions,
                  selectedKey: this.properties.documentLibraryUrl,
                  disabled: this.properties.siteCollectionUrl === "",
                }),
                PropertyPaneDropdown("layoutDropdown", {
                  label: "Select a Layout",
                  options: layoutDropdownOptions,
                  selectedKey: this.properties.layoutDropdownValue,
                  disabled: this.properties.siteCollectionUrl === "",
                }),
                // PropertyPaneButton("resetLibrary", {
                //   text: "Update",
                //   buttonType: PropertyPaneButtonType.Primary,
                //   onClick: () => {
                //     if (this.properties.layoutDropdownValue === undefined) {
                //       this.properties.layoutDropdownValue = "1"; // Default to icon layout
                //     }
                //     if (this.context.propertyPane.isPropertyPaneOpen()) {
                //       this.context.propertyPane.refresh();
                //     }
                //   },
                // }),
                PropertyPaneDropdown("categoryDropdown", {
                  label: "Select a Category",
                  options: this._categoryDropdownOptions,
                  selectedKey: this.properties.categoryDropdownValue,
                  disabled: this.properties.siteCollectionUrl === "",
                }),
                PropertyPaneTextField("itemsRowCount", {
                  label: "Items-Row Count",
                  description: "Number of items in each row",
                  value: this.properties.itemsRowCount || "5",
                }),
                PropertyPaneSlider("heightSliderValue", {
                  label: "Select a height value",
                  min: 100,
                  max: 200,
                  step: 2,
                  value: this.properties.heightSliderValue || 150,
                  showValue: true,
                }),
                PropertyPaneSlider("widthSliderValue", {
                  label: "Select a width value",
                  min: 100,
                  max: 200,
                  step: 2,
                  value: this.properties.widthSliderValue || 150,
                  showValue: true,
                }),
                ...itemIconFields,
                ...itemIconColors,
              ],
            },
          ],
        },
      ],
    };
  }
}
