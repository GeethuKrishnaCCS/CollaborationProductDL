import * as React from "react";
import {
  IDocumentItem,
  IDocumentMenuProps,
} from "../interfaces/IDocumentMenuProps";
// import { DocumentMenuService } from "../services/DocumentMenuService";
import styles from "./ListView.module.scss";

interface IListViewProps extends IDocumentMenuProps {
  currentItems: IDocumentItem[];
  currentFolderPath: string;
  navigationStack: IDocumentItem[][];
  handleFolderClick: (item: IDocumentItem) => void;
  getSharePointFileUrl: (url: string) => string;
  handleBackClick: () => void;
  renderBreadcrumb: () => JSX.Element;
}

export default function ListView(props: IListViewProps) {
  const {
    navigationStack,
    currentItems,
    handleFolderClick,
    getSharePointFileUrl,
    handleBackClick,
    renderBreadcrumb,
  } = props;

  //   const documentMenuService = new DocumentMenuService(props.context);

  return (
    <div className={styles.MainListViewContainer}>
      <div style={{ marginBottom: "10px", fontSize: "14px", color: "#333" }}>
        {renderBreadcrumb()}
      </div>
      {navigationStack.length > 0 && (
        <button
          onClick={handleBackClick}
          style={{
            marginBottom: "10px",
            background: "none",
            border: "1px solid #ccc",
            padding: "5px 10px",
            cursor: "pointer",
          }}
        >
          🔙 Back
        </button>
      )}
      <div className={styles.ListViewContentContainer}>
        <table className={styles.ListTable}>
          <thead className={styles.ListTableHeader}>
            <tr>
              <th className={styles.FileTypeIcon}></th>
              <th className={styles.HeaderName}>Name</th>
              <th className={styles.HeaderLastAccessed}>Last Accessed</th>
            </tr>
          </thead>
          <tbody>
            {currentItems.map((item, index) => (
              <tr key={index} className={styles.ListTableRow}>
                <td className={styles.FileTypeIconContainer}>
                  <div
                    className={
                      item.ServerRelativeUrl.endsWith(".pdf")
                        ? styles.PdfIcon
                        : item.ServerRelativeUrl.endsWith(".docx")
                        ? styles.DocxIcon
                        : item.ServerRelativeUrl.endsWith(".xlsx")
                        ? styles.XlsxIcon
                        : styles.DefaultFolderIcon
                    }
                  ></div>
                </td>
                <td className={styles.FileNameContainer}>
                  {item.folder ? (
                    <span
                      className={styles.FolderName}
                      style={{ cursor: "pointer", fontWeight: 600 }}
                      onClick={() => item.items && handleFolderClick(item)}
                    >
                      {item.Name}
                    </span>
                  ) : (
                    <a
                      href={getSharePointFileUrl(item.ServerRelativeUrl)}
                      target="_blank"
                      rel="noopener noreferrer"
                      className={styles.FileName}
                    >
                      {item.Name}
                    </a>
                  )}
                </td>
                <td className={styles.LastAccessedContainer}>
                  {item.LastAccessed
                    ? new Date(item.LastAccessed).toLocaleString()
                    : "-"}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
