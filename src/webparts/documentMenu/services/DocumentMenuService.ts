import { BaseService } from "../../../common/services/BaseService";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI } from "@pnp/sp";
import { getSP } from "../../../common/PnP/PnPConfig";
import {FolderItem} from "../../../common/interfaces";

export class DocumentMenuService extends BaseService {
    private spfi : SPFI;
    constructor(context: WebPartContext) {
        super(context);
        this.spfi = getSP(context);
    }

    public getCurrentUser() {
        return this.spfi.web.currentUser();
    }

    public async getLibraryData(folderUrl: string, skip:number): Promise<any> {
        const folder = this.spfi.web.getFolderByServerRelativePath(folderUrl);
        console.log(folder)

        const [subfolders, files] = await Promise.all([
            folder.folders.top(3).skip(skip)(),
            folder.files.top(3).skip(skip)()
        ]);

        let allItems: FolderItem[] = files.map((file: any) => ({
            Name: file.Name,
            ServerRelativeUrl: file.ServerRelativeUrl,
        }));

        const subfolderItemsPromises = subfolders
            .filter((subfolder: any) => subfolder.Name !== "Forms")
            .map(async (subfolder: any) => {
                const subfolderItems = await this.getLibraryData(subfolder.ServerRelativeUrl, skip);
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

    public async getFileCountInFolder(folderUrl: string): Promise<number> {
        try {
            const subfolders = await this.spfi.web.getFolderByServerRelativePath(folderUrl).folders()
            const files = await this.spfi.web.getFolderByServerRelativePath(folderUrl).files();
            let fileCount = files.length;

            for (const subfolder of subfolders) {
                if (subfolder.Name !== "Forms") {
                    const subfolderFileCount = await this.getFileCountInFolder(subfolder.ServerRelativeUrl);
                    fileCount += subfolderFileCount;
                }
            }

            return fileCount;
        } catch (error) {
            console.error("Error fetching file count:", error);
            throw error;
        }
    }

    // public async getCurrentLevelData(url: string, skip: number): Promise<any> {
    //     const folder = this.spfi.web.getFolderByServerRelativePath(url);

    //     const [subfolders, files] = await Promise.all([
    //         folder.folders.top(3).skip(skip)(),
    //         folder.files.top(3).skip(skip)()
    //     ]);

    // }

    // public async getFolderData(url: string, skip: number): Promise<any> {
    //     const folder = this.spfi.web.getFolderByServerRelativePath(url);
    //     const [subfolders, files] = await Promise.all([
    //         folder.folders.top(100).skip(skip)(),
    //         folder.files.top(100).skip(skip)()
    //     ]);

    //     let allFiles: FolderItem[] = files.map((file: any) => ({
    //                 Name: file.Name,
    //                 ServerRelativeUrl: file.ServerRelativeUrl,
    //             }));

    //     let allSubFolders: FolderItem[] = subfolders.map((subfolder: any) => ({
    //         Name: subfolder.Name,
    //         ServerRelativeUrl: subfolder.ServerRelativeUrl,
    //     }));

    //     return { allFiles, allSubFolders };
    // }

    public addNewFolder(url: string): Promise<any> {
        return this.spfi.web.folders.addUsingPath(url)
    }

    public addNewFile(url: string, fileName: any): Promise<any> {
        return this.spfi.web.getFolderByServerRelativePath(url).files.addUsingPath(fileName, "", { Overwrite: true });
    }
}