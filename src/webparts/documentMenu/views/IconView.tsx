import * as React from "react";
import {
  IDocumentItem,
  IDocumentMenuProps,
} from "../interfaces/IDocumentMenuProps";
import styles from "./IconView.module.scss";
import { Icon } from "@fluentui/react";

interface IIconViewProps extends IDocumentMenuProps {
  currentItems: IDocumentItem[];
  currentFolderPath: string;
  handleFolderClick: (item: IDocumentItem) => void;
  getSharePointFileUrl: (url: string) => string;
  onSwitchToListView: () => void;
  handleNextFolderFileSet?: () => void;
  handlePreviousFolderFileSet?: () => void;
}

export default function IconView(props: IIconViewProps) {
  const {
    currentItems,
    // handleFolderClick,
    getSharePointFileUrl,
    handleNextFolderFileSet,
    handlePreviousFolderFileSet,
  } = props;

  if (props.itemIcons === undefined) {
    props.itemIcons = {
      Folder: "Folder",
    };
  }

  const handleFolderClickAndSwitch = (item: IDocumentItem) => {
    if (item.items) {
      props.handleFolderClick(item); // existing navigation logic
      props.onSwitchToListView(); // switch to list view
    }
  };

  console.log(props.itemColors);
  return (
    <div className={styles.MainIconViewContainer}>
      {/* <div style={{ marginBottom: "10px", fontSize: "14px", color: "#333" }}>
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
      )} */}
      <div className={styles.IconViewContentContainer}>
        <button
          onClick={handlePreviousFolderFileSet}
          className={styles.ArrowPrevious}
        ></button>
        <div className={styles.IconViewFolderContainer}>
          {currentItems.map((item, index) => (
            <div
              key={index}
              className={styles.ItemContainer}
              onClick={() => item.items && handleFolderClickAndSwitch(item)}
            >
              {item.folder ? (
                <>
                  <div className={styles.FolderIconContainer}>
                    <div className={styles.FolderIcon}>
                      <Icon
                        iconName={props.itemIcons[item.Name] || "Folder"}
                        style={{
                          fontSize: 75,
                          color: props.itemColors[item.Name] || "#ffffff",
                        }}
                      />
                    </div>
                    <div className={styles.FolderName}>{item.Name}</div>
                  </div>
                </>
              ) : (
                <div className={styles.FileContainer}>
                  <a
                    // className={
                    //   item.ServerRelativeUrl.endsWith(".pdf")
                    //     ? styles.PdfIcon
                    //     : item.ServerRelativeUrl.endsWith(".docx")
                    //     ? styles.DocxIcon
                    //     : item.ServerRelativeUrl.endsWith(".xlsx")
                    //     ? styles.XlsxIcon
                    //     : styles.DefaultIcon
                    // }
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
          onClick={handleNextFolderFileSet}
          className={styles.ArrowNext}
        ></button>
      </div>
    </div>
  );
}
