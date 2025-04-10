export interface FolderItem {
    Name: string;
    ServerRelativeUrl: string;
    items?: FolderItem[]; // Include items for the subfolder
}

export interface IBaseService {
    getCurrentUser(): Promise<any>;
    getPagedListItems(queryurl: string): Promise<any>;
    getItemsSelect(queryurl: string, select: string): Promise<any>;
    getItemsSelectExpand(queryurl: string, select: string, expand: string): Promise<any>;
    getItemsById(queryurl: string, id: any): Promise<any>;
    getItemsByIdSelect(queryurl: string, id: any, select: string): Promise<any>;
    getItemsFilter(queryurl: string, filter: string): Promise<any>;
    getItemsSelectFilter(queryurl: string, select: string, filter: string): Promise<any>;
    getItemsSelectExpandFilter(queryurl: string, select: string, expand: string, filter: string): Promise<any>;
    getPagedItemsSelectExpand(queryurl: string, select: string, expand: string): Promise<any>;
    getPagedItemsSelectExpandFilter(queryurl: string, select: string, expand: string, filter: string): Promise<any>;
    getListItems(url: string): Promise<any>;
    getItemsByIdSelectExpand(queryurl: string, id: any, select: string, expand: string): Promise<any>;
    createNewItem(url: string, data: any): Promise<any>;
    updateItem(url: string, data: any, id: number): Promise<any>;
    DeleteItem(url: string, id: number): Promise<any>;
    getSelectExpand(queryurl: string, select: string, expand: string): Promise<any>;
    getUser(userId: number): Promise<any>;
    getAllFoldersAndFiles(folderUrl: string): Promise<FolderItem[]>;
    getFileContent(fileUrl: string): Promise<ArrayBuffer>;
    getDocLibrary(url: string): Promise<any>;
    uploadDocument(libraryName: string, Filename: any, filedata: any): Promise<any>;
    getLibraryItem(queryurl: string, id: number): Promise<any>;
    getChoiceListItems(url: string, field: string): Promise<any>;
}