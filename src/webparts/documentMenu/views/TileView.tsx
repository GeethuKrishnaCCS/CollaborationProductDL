import * as React from "react";
import styles from "./TileView.module.scss";
import { useState, useEffect } from "react";
import { DocumentMenuService } from "../services/DocumentMenuService";
import {
  IDocumentItem,
  IDocumentMenuProps,
} from "../interfaces/IDocumentMenuProps";

interface ITileViewProps extends IDocumentMenuProps {
  currentItems: IDocumentItem[];
  currentFolderPath: string;
  navigationStack: IDocumentItem[][];
  handleFolderClick: (item: IDocumentItem) => void;
  getSharePointFileUrl: (url: string) => string;
  handleBackClick: () => void;
  renderBreadcrumb: () => JSX.Element;
}

export default function TileView(props: ITileViewProps) {
  const {
    navigationStack,
    currentItems,
    handleFolderClick,
    getSharePointFileUrl,
    handleBackClick,
    renderBreadcrumb,
  } = props;
  const documentMenuService = new DocumentMenuService(props.context);

  const FileCount = ({ folderUrl }: { folderUrl: string }) => {
    const [fileCount, setFileCount] = useState<number | null>(null);

    useEffect(() => {
      const fetchFileCount = async () => {
        try {
          const count = await documentMenuService.getFileCountInFolder(
            folderUrl
          );
          setFileCount(count);
        } catch (error) {
          console.error("Error fetching file count:", error);
          setFileCount(null);
        }
      };

      fetchFileCount();
    }, [currentItems]);

    return <span>{fileCount !== null ? fileCount : "..."}</span>; // Show "..." while loading
  };

  return (
    <div className={styles.MainTileViewContainer}>
      {/* {showModal && renderModal()} */}
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
      <div className={styles.TileViewContentContainer}>
        <button
          // onClick={handlePreviousFolderFileSet}
          className={styles.ArrowPrevious}
        ></button>
        <div className={styles.FolderAndFileContainer}>
          {currentItems.map((item, index) => (
            <div
              key={index}
              className={styles.ItemContainer}
              onClick={() => item.items && handleFolderClick(item)}
            >
              {item.folder ? (
                <>
                  <div className={styles.FolderContainer}>
                    <div className={styles.FolderIcon}>
                      <span className={styles.FileCount}>
                        {/* Dynamically fetch and display file count */}
                        <FileCount folderUrl={item.ServerRelativeUrl} />
                      </span>{" "}
                    </div>
                    <div className={styles.FolderName}>{item.Name}</div>
                  </div>
                </>
              ) : (
                <div className={styles.FileContainer}>
                  <a
                    className={
                      item.ServerRelativeUrl.endsWith(".pdf")
                        ? styles.PdfIcon
                        : item.ServerRelativeUrl.endsWith(".docx")
                        ? styles.DocxIcon
                        : item.ServerRelativeUrl.endsWith(".xlsx")
                        ? styles.XlsxIcon
                        : styles.DefaultIcon
                    }
                    href={getSharePointFileUrl(item.ServerRelativeUrl)}
                    target="_blank"
                    rel="noopener noreferrer"
                  ></a>
                  <div className={styles.FileName}>{item.Name}</div>
                </div>
              )}
            </div>
          ))}
        </div>
        <button
          // onClick={handleNextFolderFileSet}
          className={styles.ArrowNext}
        ></button>
      </div>
    </div>
  );
}
