import { BaseService } from "../../../common/services/BaseService";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI } from "@pnp/sp";
import { getSP } from "../../../common/PnP/PnPConfig";

export class DocumentMenuService extends BaseService {
    private spfi : SPFI;
    constructor(context: WebPartContext) {
        super(context);
        this.spfi = getSP(context);
    }

    public getCurrentUser() {
        return this.spfi.web.currentUser();
    }

    public addNewFolder(url: string): Promise<any> {
        return this.spfi.web.folders.addUsingPath(url)
    }

    public addNewFile(url: string, fileName: any): Promise<any> {
        return this.spfi.web.getFolderByServerRelativePath(url).files.addUsingPath(fileName, "", { Overwrite: true });
    }
}