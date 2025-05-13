import { BaseService } from "../../../common/services/BaseService";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";
import { getSP } from "../../../common/PnP/PnPConfig";
import { FolderItem } from "../../../common/interfaces";
import "@pnp/sp/search";

export class DocumentMenuService extends BaseService {
  private spfi: SPFI;
  constructor(context: WebPartContext) {
    super(context);
    this.spfi = getSP(context);
  }

  public getCurrentUser() {
    return this.spfi.web.currentUser();
  }

  public async getLibraryData(folderUrl: string, skip: number): Promise<any> {
    const folder = this.spfi.web.getFolderByServerRelativePath(folderUrl);
    // console.log(folder);

    const [subfolders, files] = await Promise.all([
      folder.folders.top(3).skip(skip)(),
      folder.files.top(3).skip(skip)(),
    ]);

    let allItems: FolderItem[] = files.map((file: any) => ({
      Name: file.Name,
      ServerRelativeUrl: file.ServerRelativeUrl,
      folder: false,
    }));

    const subfolderItemsPromises = subfolders
      .filter((subfolder: any) => subfolder.Name !== "Forms")
      .map(async (subfolder: any) => {
        // const subfolderItems = await this.getLibraryDataWithoutSkip(
        //   subfolder.ServerRelativeUrl
        // );
        return {
          Name: subfolder.Name,
          ServerRelativeUrl: subfolder.ServerRelativeUrl,
          folder: true,
        };
      });

    const subfolderItems = await Promise.all(subfolderItemsPromises);
    allItems = allItems.concat(subfolderItems); // Combine files and subfolders
    return allItems;
  }

  public async getLibraryDataWithoutSkip(folderUrl: string): Promise<any> {
    const folder = this.spfi.web.getFolderByServerRelativePath(folderUrl);
    // console.log(folder);

    const [subfolders, files] = await Promise.all([
      folder.folders.top(3)(),
      folder.files.top(3)(),
    ]);

    let allItems: FolderItem[] = files.map((file: any) => ({
      Name: file.Name,
      ServerRelativeUrl: file.ServerRelativeUrl,
      folder: false,
    }));

    const subfolderItemsPromises = subfolders
      .filter((subfolder: any) => subfolder.Name !== "Forms")
      .map(async (subfolder: any) => {
        // const subfolderItems = await this.getLibraryDataWithoutSkip(
        //   subfolder.ServerRelativeUrl
        // );
        return {
          Name: subfolder.Name,
          ServerRelativeUrl: subfolder.ServerRelativeUrl,
          folder: true,
        };
      });

    const subfolderItems = await Promise.all(subfolderItemsPromises);
    allItems = allItems.concat(subfolderItems); // Combine files and subfolders
    return allItems;
  }

  //   public async getNextFoldersAndFiles(folderUrl: string): Promise<any> {}

  public async getFileCountInFolder(folderUrl: string): Promise<number> {
    try {
      const subfolders = await this.spfi.web
        .getFolderByServerRelativePath(folderUrl)
        .folders();
      const files = await this.spfi.web
        .getFolderByServerRelativePath(folderUrl)
        .files();
      let fileCount = files.length;

      for (const subfolder of subfolders) {
        if (subfolder.Name !== "Forms") {
          const subfolderFileCount = await this.getFileCountInFolder(
            subfolder.ServerRelativeUrl
          );
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
  //   const folder = this.spfi.web.getFolderByServerRelativePath(url);
  //   const [subfolders, files] = await Promise.all([
  //     folder.folders.top(3).skip(skip)(),
  //     folder.files.top(3).skip(skip)(),
  //   ]);

  //   let allFiles: FolderItem[] = files.map((file: any) => ({
  //     Name: file.Name,
  //     ServerRelativeUrl: file.ServerRelativeUrl,
  //   }));

  //   let allSubFolders: FolderItem[] = subfolders.map((subfolder: any) => ({
  //     Name: subfolder.Name,
  //     ServerRelativeUrl: subfolder.ServerRelativeUrl,
  //   }));

  //   return { allFiles, allSubFolders };
  // }

  //   public addNewFolder(url: string): Promise<any> {
  //     return this.spfi.web.folders.addUsingPath(url);
  //   }

  //   public addNewFile(url: string, fileName: any): Promise<any> {
  //     return this.spfi.web
  //       .getFolderByServerRelativePath(url)
  //       .files.addUsingPath(fileName, "", { Overwrite: true });
  //   }

  public async searchFilesAndFolders(
    filename: string,
    currentFolderPath: string
  ): Promise<any> {
    const searchQuery = `(IsDocument:1 OR IsContainer:true) AND 
    Path:"https://ccsdev01.sharepoint.com/${currentFolderPath}/*" AND 
    (FileName:"${filename}" OR Title:"${filename}" OR Path:"${filename}")`;

    // Execute search with additional parameters
    const results = await this.spfi.search({
      Querytext: searchQuery,
      RowLimit: 500,
      SelectProperties: [
        "Title",
        "Path",
        "Filename",
        "FileExtension",
        "ServerRelativeUrl",
        "ContentClass",
        "IsContainer",
        "IsDocument",
        "LastModifiedTime",
      ],
      TrimDuplicates: false,
    });
    // console.log("Primary results:", results.PrimarySearchResults);

    const formattedResults = await Promise.all(
      results.PrimarySearchResults.map(async (item) => {
        // console.log("container", item.IsContainer === false);
        // const subItems = await this.getLibraryDataWithoutSkip(
        //   (item.Path ?? "").replace(/^https?:\/\/[^/]+/, "")
        // );
        if (item.FileExtension === null) {
          // console.log("container", item.IsContainer);
          return {
            Name: item.Title,
            ServerRelativeUrl: (item.Path ?? "").replace(
              /^https?:\/\/[^/]+/,
              ""
            ),
            folder: true,
            // items: subItems,
          };
        } else {
          return {
            Name: item.Title, // Handle files vs folders
            ServerRelativeUrl: (item.Path ?? "").replace(
              /^https?:\/\/[^/]+/,
              ""
            ), // Fallback to Path
            folder: false,
          };
        }
      })
    );
    // console.log("Formatted results:", formattedResults);
    return formattedResults;
  }
  catch(error: unknown) {
    console.error("Search error:", error);
    throw error;
  }
}
