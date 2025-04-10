import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IDocumentMenuProps {
  context: WebPartContext;
  description: string;
  userDisplayName: string;
  documentUrl: string;
}

// This interface is used to define the structure of the document item
export interface IDocumentItem {
  Name: string;
  ServerRelativeUrl: string;
  items?: IDocumentItem[];
}

