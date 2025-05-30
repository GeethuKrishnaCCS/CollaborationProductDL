import { BaseService } from "../../../common/services/BaseService";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";
import { getSP } from "../../../common/PnP/PnPConfig";
import { FolderItem } from "../../../common/interfaces";
import "@pnp/sp/search";
import "@pnp/sp/sites";
import "@pnp/sp/batching";
import { IDocumentItem } from "../interfaces/IDocumentMenuProps";

// import { spfi } from "@pnp/sp";

export class DocumentMenuService extends BaseService {
  private spfi: SPFI;

  constructor(context: WebPartContext, siteCollectionUrl?: string) {
    super(context);
    this.spfi = getSP(context, siteCollectionUrl);
  }

  public getCurrentUser() {
    return this.spfi.web.currentUser();
  }

  public async getLibraryData(folderUrl: string, skip: number): Promise<any> {
    const folder = this.spfi.web.getFolderByServerRelativePath(folderUrl);

    const [subfolders, files] = await Promise.all([
      folder.folders.top(5).skip(skip)(),
      folder.files.top(5).skip(skip)(),
    ]);

    // console.log("subfolders", subfolders);
    // console.log("files", files);

    let allItems: FolderItem[] = files.map((file: any) => ({
      Name: file.Name,
      ServerRelativeUrl: file.ServerRelativeUrl,
      LastAccessed: file.TimeLastModified,
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
          LastAccessed: subfolder.TimeLastModified,
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
      folder.folders.top(5)(),
      folder.files.top(5)(),
    ]);

    console.log("subfolders", subfolders);
    console.log("files", files);

    let allItems: FolderItem[] = files.map((file: any) => ({
      Name: file.Name,
      ServerRelativeUrl: file.ServerRelativeUrl,
      LastAccessed: file.TimeLastModified,
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
          LastAccessed: subfolder.TimeLastModified,
          folder: true,
        };
      });

    const subfolderItems = await Promise.all(subfolderItemsPromises);
    allItems = allItems.concat(subfolderItems); // Combine files and subfolders
    return allItems;
  }

  public async _getRecursiveFileCountInFolder(
    folderUrl: string
    // We don't pass the batch itself, but rather create one per level or use the main one if preferred
    // For simplicity and to batch operations at each level:
  ): Promise<number> {
    const [batch, executeBatch] = this.spfi.batched();
    let totalFiles = 0;

    // Get files in the current folder
    const filesPromise = batch.web
      .getFolderByServerRelativePath(folderUrl)
      .files();

    // Get subfolders in the current folder
    const subFoldersPromise = batch.web
      .getFolderByServerRelativePath(folderUrl)
      .folders();

    // Execute the batch for the current folder's files and subfolders
    await executeBatch();

    // Await the results after batch execution
    const files = await filesPromise;
    totalFiles += files.length;

    const subFolders = await subFoldersPromise;

    // Recursively call for each subfolder
    // We can do this sequentially or in parallel. Parallel is faster but more resource-intensive.
    const subFolderCountPromises: Promise<number>[] = [];
    for (const subFolder of subFolders) {
      // subFolder object from PnPjs usually has ServerRelativeUrl
      if (subFolder.ServerRelativeUrl) {
        subFolderCountPromises.push(
          this._getRecursiveFileCountInFolder(subFolder.ServerRelativeUrl)
        );
      } else {
        console.warn("Subfolder object missing ServerRelativeUrl:", subFolder);
      }
    }

    // Wait for all recursive calls for subfolders to complete
    const subFolderCounts = await Promise.all(subFolderCountPromises);
    subFolderCounts.forEach((count) => {
      totalFiles += count;
    });

    return totalFiles;
  }

  public async getFileCountInFolder(
    initialFolderItems: IDocumentItem
  ): Promise<any> {
    // Process each initial folder. You can do this sequentially or in parallel.
    // Using Promise.all for parallel processing of initial folders:
    const processingPromises = async () => {
      try {
        console.log(
          `Processing folder: ${initialFolderItems.ServerRelativeUrl}`
        );
        const count = await this._getRecursiveFileCountInFolder(
          initialFolderItems.ServerRelativeUrl
        );
        return count;
      } catch (error) {
        console.error(
          `Error processing folder ${initialFolderItems.ServerRelativeUrl}:`,
          error
        );
        // Decide how to handle errors for individual folders
        // For example, return a result indicating failure or skip it
        return -1;
      }
    };

    const allResults = await Promise.all([processingPromises()]);
    // results.push(...allResults); // Filter out errors if you marked them

    return allResults;
  }

  // for (let i = 0; i < res.length; i++) {
  //   // console.log("res", res[i]);
  // }
  // try {
  //   const subfolders = await this.spfi.web
  //     .getFolderByServerRelativePath(folderUrl)
  //     .folders();
  //   const files = await this.spfi.web
  //     .getFolderByServerRelativePath(folderUrl)
  //     .files();
  //   let fileCount = files.length;

  //   for (const subfolder of subfolders) {
  //     if (subfolder.Name !== "Forms") {
  //       const subfolderFileCount = await this.getFileCountInFolder(
  //         subfolder.ServerRelativeUrl
  //       );
  //       fileCount += subfolderFileCount;
  //     }
  //   }

  //   return fileCount;
  // } catch (error) {
  //   console.error("Error fetching file count:", error);
  //   throw error;
  // }

  public async searchFilesAndFolders(
    filename: string,
    currentFolderPath: string
  ): Promise<any> {
    const searchQuery = `(IsDocument:1 OR IsContainer:true) AND 
    Path:"https://ccsdev01.sharepoint.com/${currentFolderPath}/*" AND 
    (FileName:"*${filename}*")`;

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
            LastAccessed: item.LastModifiedTime,
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
            LastAccessed: item.LastModifiedTime,
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

  public async getLibraryOptions(siteCollectionUrl: string): Promise<any> {
    try {
      const docLibs = await this.spfi.site.getDocumentLibraries(
        siteCollectionUrl
      );

      console.log("docLibs", docLibs);
      return docLibs;
    } catch (error) {
      console.error("Error fetching library by server relative URL:", error);
      throw error;
    }
  }

  public async getFieldsForUrl(): Promise<any[]> {
    try {
      const item = await this.spfi.web
        .getFileByServerRelativePath(
          "/sites/ProductDevelopment/Shared Documents/Applications/Book1.xlsx"
        )
        .listItemAllFields();
      console.log("item", item);
      return Object.keys(item).map((key) => ({
        field: key,
        value: item[key],
      }));
    } catch (error) {
      console.error("Error retrieving fields for URL:", error);
      throw error;
    }
  }
}
