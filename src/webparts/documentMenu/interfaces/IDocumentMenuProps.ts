import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IDocumentMenuProps {
  onCurrentItemsChange: (items: IDocumentItem[]) => void;
  onLayoutStateChange?: (
    activeIconLayout: string,
    navigationStackLength: number
  ) => void;
  context: WebPartContext;
  description: string;
  userDisplayName: string;
  documentLibraryUrl: string;
  layoutDropdownValue: any;
  siteCollectionUrl: string;
  itemsRowCount: string;
  heightSliderValue?: number;
  widthSliderValue?: number;
  itemIcons: { [key: string]: string };
  itemColors: { [key: string]: string };
}

// This interface is used to define the structure of the document item
export interface IDocumentItem {
  Name: string;
  ServerRelativeUrl: string;
  LastAccessed: string;
  folder?: boolean;
  items?: IDocumentItem[];
}
