import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI, SPFx } from "@pnp/sp/presets/all";
import { getSP } from "../PnP/PnPConfig";
import { FolderItem } from "../interfaces";
import { IBaseService } from "../interfaces";
import "@pnp/sp/attachments";
import "@pnp/sp/site-groups/web";
import "@pnp/sp/sites";
import "@pnp/sp/webs";
import "@pnp/sp/lists"
import "@pnp/sp/fields";

export class BaseService implements IBaseService {
    private sp: SPFI;
    constructor(context: WebPartContext, siteUrl?: string) {
        this.sp = siteUrl ? new SPFI(siteUrl).using(SPFx(context)) : getSP(context);
    }

    public getCurrentUser() {
        return this.sp.web.currentUser();
    }

    public getCurrentUserGroups() {
        return this.sp.web.currentUser.groups();
    }

    public async getPagedListItems(queryurl: string): Promise<any> {
        let collection: any = []
        for await (const items of this.sp.web.getList(queryurl).items.select('*').top(250) as any) {
            collection = collection.concat(items)
        }
        return collection
    }

    public getItemsSelect(queryurl: string, select: string): Promise<any> {
        return this.sp.web.getList(queryurl).items
            .select(select)()
    }

    public getItemsSelectExpand(queryurl: string, select: string, expand: string): Promise<any> {
        return this.sp.web.getList(queryurl).items
            .select(select)
            .expand(expand)()
    }

    public getItemsById(queryurl: string, id: any): Promise<any> {
        return this.sp.web.getList(queryurl).items
            .getById(id)()
    }

    public getItemsByIdSelect(queryurl: string, id: any, select: string): Promise<any> {
        return this.sp.web.getList(queryurl).items
            .getById(id)
            .select(select)()
    }

    public getItemsFilter(queryurl: string, filter: string): Promise<any> {
        return this.sp.web.getList(queryurl).items
            .filter(filter)()
    }

    public getItemsSelectFilter(queryurl: string, select: string, filter: string): Promise<any> {
        return this.sp.web.getList(queryurl).items
            .select(select)
            .filter(filter)()
    }

    public getItemsSelectExpandFilter(queryurl: string, select: string, expand: string, filter: string): Promise<any> {
        return this.sp.web.getList(queryurl).items
            .select(select)
            .expand(expand)
            .filter(filter)()
    }

    public async getPagedItemsSelectExpand(queryurl: string, select: string, expand: string): Promise<any> {
        let collection: any = []
        for await (const items of this.sp.web.getList(queryurl).items.select(select).expand(expand).top(250) as any) {
            collection = collection.concat(items)
        }
        return collection
    }
    public async getPagedItemsSelectExpandFilter(queryurl: string, select: string, expand: string, filter: string): Promise<any> {
        let collection: any = []
        for await (const items of this.sp.web.getList(queryurl).items.select(select).expand(expand).filter(filter).top(250) as any) {
            collection = collection.concat(items)
        }
        return collection
    }

    public getListItems(url: string): Promise<any> {
        return this.sp.web.getList(url).items();
    }

    public getItemsByIdSelectExpand(queryurl: string, id: any, select: string, expand: string): Promise<any> {
        return this.sp.web.getList(queryurl).items
            .getById(id)
            .select(select)
            .expand(expand)()
    }
    public createNewItem(url: string, data: any): Promise<any> {
        return this.sp.web.getList(url).items.add(data);
    }
    public updateItem(url: string, data: any, id: number): Promise<any> {
        return this.sp.web.getList(url).items.getById(id).update(data);
    }
    public DeleteItem(url: string, id: number): Promise<any> {
        return this.sp.web.getList(url).items.getById(id).delete();
    }

    public getSelectExpand(queryurl: string, select: string, expand: string): Promise<any> {
        return this.sp.web.getList(queryurl).items
            .select(select)
            .expand(expand)()
    }

    public async getUser(userId: number): Promise<any> {
        return this.sp.web.getUserById(userId)();
    }

    public async getAllFoldersAndFiles(folderUrl: string): Promise<FolderItem[]> {
        const folder = this.sp.web.getFolderByServerRelativePath(folderUrl);

        // Fetch subfolders and files in parallel
        const [subfolders, files] = await Promise.all([
            folder.folders(),
            folder.files()
        ]);

        // Create the file items array
        let allItems: FolderItem[] = files.map((file: any) => ({
            Name: file.Name,
            ServerRelativeUrl: file.ServerRelativeUrl,
        }));

        // Use Promise.all to fetch subfolder contents in parallel
        const subfolderItemsPromises = subfolders
            .filter((subfolder: any) => subfolder.Name !== "Forms")
            .map(async (subfolder: any) => {
                const subfolderItems = await this.getAllFoldersAndFiles(subfolder.ServerRelativeUrl);
                return {
                    Name: subfolder.Name,
                    ServerRelativeUrl: subfolder.ServerRelativeUrl,
                    items: subfolderItems
                };
            });

        const subfolderItems = await Promise.all(subfolderItemsPromises);
        allItems = allItems.concat(subfolderItems); // Combine files and subfolders
        return allItems;
    }
    public async getFileContent(fileUrl: string): Promise<ArrayBuffer> {
        return this.sp.web.getFileByServerRelativePath(fileUrl).getBuffer();
    }

    public getDocLibrary(url: string): Promise<any> {
        return this.sp.site.getDocumentLibraries(url);
    }

    public async uploadDocument(libraryName: string, Filename: any, filedata: any): Promise<any> {
        const response = await this.sp.web.getFolderByServerRelativePath(libraryName).files.addUsingPath(Filename, filedata, { Overwrite: true });
        return response
    }

    public getLibraryItem(queryurl: string, id: number): Promise<any> {
        return this.sp.web.getList(queryurl).items.getById(id).select("FileRef,FileLeafRef")()
    }

    public getChoiceListItems(url: string, field: string): Promise<any> {
        return this.sp.web.getList(url).fields.getByInternalNameOrTitle(field)();
    }

}