import * as React from "react";
import styles from "./TileView.module.scss";
import { useState, useEffect } from "react";
import {
  TextField,
  ITextFieldStyles,
  IStyleFunctionOrObject,
  ITextFieldStyleProps,
} from "@fluentui/react";
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
  searchValue: string;
  handlesearchValue: (value: string) => void;
  handleBackClick: () => void;
  renderBreadcrumb: () => JSX.Element;
}

const searchFieldStyles: IStyleFunctionOrObject<
  ITextFieldStyleProps,
  ITextFieldStyles
> = {
  fieldGroup: {
    height: "40px", // Set height
    width: "400px", // Set width
    backgroundColor: "#FFFFFF", // Set background color
    border: "1px solid #C8EFFE", // Set border color
    borderRadius: "10px",
    selectors: {
      "::after": {
        border: "none",
        borderRadius: "10px",
      },
      ":focus-within": {
        border: "1px solid rgb(177, 217, 233)",
        borderRadius: "10px",
      },
      ":focus": {
        border: "none",
      },
      ":active": {
        border: "none",
      },
      ":hover": {
        border: "1px solid rgb(177, 217, 233)",
      },
    },
  },
};

export default function TileView(props: ITileViewProps) {
  const {
    navigationStack,
    currentItems,
    handleFolderClick,
    getSharePointFileUrl,
    handlesearchValue,
    searchValue,
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
    }, [folderUrl]);

    return <span>{fileCount !== null ? fileCount : "..."}</span>; // Show "..." while loading
  };

  return (
    <div className={styles.MainContainer}>
      <div className={styles.Header}>
        {/* Render Breadcrumb */}
        {/* <button
                  onClick={() => setShowModal(true)}
                  style={{
                    marginRight: "10px",
                    background: "#4CAF50",
                    color: "white",
                    border: "none",
                    padding: "10px 15px",
                    cursor: "pointer",
                    borderRadius: "5px",
                  }}
                >
                  ➕ Create/Upload File
                </button>
                <button
                  onClick={handleCreateFolder}
                  style={{
                    marginRight: "10px",
                    background: "#2196F3",
                    color: "white",
                    border: "none",
                    padding: "10px 15px",
                    cursor: "pointer",
                    borderRadius: "5px",
                  }}
                >
                  ➕ Create Folder
                </button> */}
        <div>
          <button
            // onClick={handlePreviousFolderFileSet}
            style={{
              marginRight: "10px",
              background: "#FF9800",
              color: "white",
              border: "none",
              padding: "10px 15px",
              cursor: "pointer",
              borderRadius: "5px",
            }}
          >
            ⏪ Back
          </button>
          <button
            // onClick={handleNextFolderFileSet}
            style={{
              background: "#FF9800",
              color: "white",
              border: "none",
              padding: "10px 15px",
              cursor: "pointer",
              borderRadius: "5px",
            }}
          >
            ⏩ Next
          </button>
        </div>
        {/* Search Field */}
        <div className={styles.SearchField}>
          <TextField
            styles={searchFieldStyles}
            placeholder="Search..."
            value={searchValue}
            onChange={(e, newValue) => handlesearchValue(newValue || "")}
          />
        </div>
      </div>
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
      <div className={styles.ContentContainer}>
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
