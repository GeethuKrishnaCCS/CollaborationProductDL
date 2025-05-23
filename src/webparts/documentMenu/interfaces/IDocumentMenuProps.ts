import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IDocumentMenuProps {
  context: WebPartContext;
  description: string;
  userDisplayName: string;
  documentUrl: string;
  layoutDropdown: string;
}

// This interface is used to define the structure of the document item
export interface IDocumentItem {
  Name: string;
  ServerRelativeUrl: string;
  LastAccessed: string;
  folder?: boolean;
  items?: IDocumentItem[];
}
