import * as React from "react";
import type {
  IDocumentMenuProps,
  IDocumentItem,
} from "../interfaces/IDocumentMenuProps";
import styles from "./DocumentMenu.module.scss";
import { useState, useEffect, useRef } from "react";
// import { BaseService } from "../../../common/services/BaseService";
import { DocumentMenuService } from "../services/DocumentMenuService";
import TileView from "../views/TileView";
import {
  TextField,
  ITextFieldStyles,
  IStyleFunctionOrObject,
  ITextFieldStyleProps,
} from "@fluentui/react";
import ListView from "../views/ListView";
import IconView from "../views/IconView";
import debounce from "lodash/debounce";
import { Spinner } from "@fluentui/react/lib/Spinner";

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

export default function DocumentMenu(props: IDocumentMenuProps) {
  // const [libraryData, setLibraryData] = useState<IDocumentItem[]>([]);
  const [currentItems, setCurrentItems] = useState<IDocumentItem[]>([]); // Items to display at the current level
  const [navigationStack, setNavigationStack] = useState<IDocumentItem[][]>([]); // Stack to track navigation levels
  const [currentFolderPath, setCurrentFolderPath] = useState(
    props.documentLibraryUrl ? props.documentLibraryUrl : ""
  );
  // const [currentSearchItems, setCurrentSearchItems] = useState<IDocumentItem[]>([]);
  // const [showModal, setShowModal] = useState(false);
  const [breadCrumbItems, setBreadCrumbItems] = useState(["Documents"]);
  const [searchValue, setSearchValue] = useState("");
  const [activeListTileLayout, setActiveListTileLayout] = useState<
    "tile" | "list"
  >("tile");
  const [activeIconLayout, setActiveIconLayout] = useState<"icon" | "list">(
    "icon"
  );
  const [loading, setLoading] = useState(false);

  const libraryName = props.documentLibraryUrl ? props.documentLibraryUrl : "";
  const documentMenuService = new DocumentMenuService(
    props.context,
    props.siteCollectionUrl
  );
  const [paginationStack, setPaginationStack] = useState<IDocumentItem[][]>([]);
  // const [pageCount, setPageCount] = useState(0);
  // const baseService = new BaseService(props.context);
  const pageCount = useRef(0);

  //Initial library data and current items.
  //Triggers when documentLibraryUrl or siteCollectionUrl changes
  useEffect(() => {
    console.log(
      "libraryName:",
      libraryName,
      "props.siteCollectionUrl:",
      props.siteCollectionUrl
    );
    if (libraryName === "" || props.siteCollectionUrl === "") {
      setCurrentItems([]);
      console.log("Document library URL is not provided.");
      return;
    }

    //Get all Folder information
    documentMenuService
      .getLibraryData(libraryName, 0)
      .then((data) => {
        // console.log("Fetched library data:", data);
        // setLibraryData(data);
        for (let item of data) {
          if (item.folder) {
            documentMenuService
              .getLibraryDataWithoutSkip(item.ServerRelativeUrl)
              .then((data) => {
                item["items"] = data; // Initialize items array for folders
                // console.log(currentItems);
              })
              .catch((error) => {
                console.error("Error fetching folder data:", error);
              });
          }
        }
        setCurrentItems(data);
        setNavigationStack([data]);
      })
      .catch((error) => console.error("Error fetching library data:", error));

    // documentMenuService.searchFilesAndFolders();
  }, [props.documentLibraryUrl, props.siteCollectionUrl]);

  // Fetch distinct category values when categoryDropdownValue changes
  useEffect(() => {
    documentMenuService
      .getCategoryDistinctValues(props.documentLibraryUrl, "Department")
      .then(async (distinctCategoryValues) => {
        // Create an array of promises to fetch files for each category
        const categoryResults = await Promise.all(
          distinctCategoryValues.map(async (value: any) => {
            const files = await documentMenuService.getCategoryValueFiles(
              props.documentLibraryUrl,
              value
            );
            console.log("documentUrl", props.documentLibraryUrl);
            return {
              Name: value,
              items: files, // Attach the files here
              folder: true,
            };
          })
        );
        console.log("Mapped category results:", categoryResults);
        // You can now set this to state if you want to display it
        // setCurrentItems(categoryResults);g
      });
  }, [props.categoryDropdownValue]);

  // Fetch folder data for each item in currentItems
  // This effect runs when currentItems changes, fetching data for folders
  // User doesnt have to wait for another api call to fetch folder data
  useEffect(() => {
    const fetchFolderData = async () => {
      if (currentItems) {
        for (let item of currentItems) {
          if (item.folder) {
            try {
              const data = await documentMenuService.getLibraryDataWithoutSkip(
                item.ServerRelativeUrl
              );
              item["items"] = data; // Initialize items array for folders
              // console.log("Fetched folder data:", currentItems);
            } catch (error) {
              console.error("Error fetching folder data:", error);
            }
          }
        }
      }
    };

    fetchFolderData();
    props.onCurrentItemsChange && props.onCurrentItemsChange(currentItems);
  }, [currentItems]);

  useEffect(() => {
    props.onLayoutStateChange &&
      props.onLayoutStateChange(activeIconLayout, navigationStack.length);
  }, [activeIconLayout, navigationStack.length]);

  //Handle next folder/file set
  const handleNextFolderFileSet = async () => {
    pageCount.current += 5; // Increment page count
    setPaginationStack((prevStack) => [...prevStack, currentItems]); // Push the current items to the stack
    documentMenuService
      .getLibraryData(currentFolderPath, pageCount.current)
      .then((data) => {
        console.log("Fetched library data:", data);
        // setLibraryData(data);
        setCurrentItems(data);
      })
      .catch((error) => console.error("Error fetching library data:", error));
  };

  // Handle previous folder/file set
  const handlePreviousFolderFileSet = async () => {
    if (pageCount.current === 0) {
      alert("No more previous folders or files.");
      return;
    }
    pageCount.current -= 5; // Decrement page count

    if (paginationStack.length > 0) {
      const previousItems = paginationStack[paginationStack.length - 1]; // Get the last items from the stack
      setCurrentItems(previousItems);
      setPaginationStack((prevStack) => prevStack.slice(0, -1)); // Remove the last items from the stack
    }
    // documentMenuService
    //   .getLibraryData(currentFolderPath, pageCount - 5)
    //   .then((data) => {
    //     console.log("Fetched library data:", data);
    //     // setLibraryData(data);
    //     setCurrentItems(data);
    //   })
    //   .catch((error) => console.error("Error fetching library data:", error));
  };

  // Handle folder click
  const handleFolderClick = (folder: IDocumentItem) => {
    if (folder.items) {
      pageCount.current = 0; // Reset page count when navigating to a folder
      setNavigationStack((prevStack) => [...prevStack, folder.items ?? []]); // Push the current level to the stack
      setCurrentItems(folder.items);
      let newFolderName = folder.ServerRelativeUrl.split("/");
      setBreadCrumbItems((prevItems) => [
        ...prevItems,
        newFolderName[newFolderName.length - 1],
      ]); // Update breadcrumb items
      setCurrentFolderPath(folder.ServerRelativeUrl); // Set the current folder path
      setPaginationStack([]);
    }
  };

  // Handle back navigation
  // const handleBackClick = () => {
  //   if (navigationStack.length > 0) {
  //     const previousLevel = navigationStack[navigationStack.length - 1]; // Get the previous level
  //     setNavigationStack((prevStack) => prevStack.slice(0, -1)); // Pop the last level from the stack
  //     setCurrentItems(previousLevel);
  //     setCurrentFolderPath(currentFolderPath.split("/").slice(0, -1).join("/")); // Update the current folder path
  //     setBreadCrumbItems((prevItems) => prevItems.slice(0, -1)); // Remove the last breadcrumb item
  //     setActiveIconLayout("icon");
  //   }
  // };

  // Handle file creation
  // const handleCreateFile = async () => {
  //   let fileName = prompt("Enter the name of the new file:");
  //   fileName += ".docx";
  //   if (fileName) {
  //     try {
  //       documentMenuService.addNewFile(currentFolderPath, fileName);

  //       const newFile: IDocumentItem = {
  //         Name: fileName,
  //         ServerRelativeUrl: `${currentFolderPath}/${fileName}`,
  //       };
  //       for (const folderItems of navigationStack) {
  //         const folder = folderItems.find(
  //           (item) => item.Name === currentFolderPath.split("/").pop()
  //         );
  //         if (folder) {
  //           folder.items = folder.items || [];
  //           folder.items.push(newFile);
  //         }
  //       }
  //       setCurrentItems((prevItems) => [...prevItems, newFile]);
  //       alert("File created successfully!");
  //     } catch (error) {
  //       console.error("Error creating file:", error);
  //       alert("Failed to create file. Please try again.");
  //     }
  //   } else {
  //     alert("Invalid file name. Only .docx files are supported for creation.");
  //   }
  // };

  // Handle file upload
  // const handleUploadFile = async () => {
  //   const input = document.createElement("input");
  //   input.type = "file";

  //   const allowedExtensions = [
  //     ".doc",
  //     ".docx",
  //     ".xls",
  //     ".xlsx",
  //     ".ppt",
  //     ".pptx",
  //     ".pdf",
  //     ".txt",
  //     ".csv",
  //     ".one",
  //     ".vsd",
  //     ".vsdx",
  //   ];

  //   // Set accept attribute to show only these files in dialog
  //   input.accept = allowedExtensions
  //     .map((ext) => `${ext},.${ext.toUpperCase()}`)
  //     .join(",");

  //   input.onchange = async (event: any) => {
  //     const file = event.target.files[0];
  //     if (file) {
  //       try {
  //         // Get file extension
  //         const fileExtension = file.name.split(".").pop().toLowerCase();

  //         // Validate file type
  //         if (!allowedExtensions.includes(`.${fileExtension}`)) {
  //           alert(
  //             `Invalid file type. Please upload only Microsoft Office files or PDFs. Allowed formats: ${allowedExtensions.join(
  //               ", "
  //             )}`
  //           );
  //           return;
  //         }

  //         await baseService.uploadDocument(currentFolderPath, file.name, file);

  //         const newFile: IDocumentItem = {
  //           Name: file.name,
  //           ServerRelativeUrl: `${currentFolderPath}/${file.name}`,
  //         };
  //         for (const folderItems of navigationStack) {
  //           const folder = folderItems.find(
  //             (item) => item.Name === currentFolderPath.split("/").pop()
  //           );
  //           if (folder) {
  //             folder.items = folder.items || [];
  //             folder.items.push(newFile);
  //           }
  //         }
  //         setCurrentItems((prevItems) => [...prevItems, newFile]);
  //         alert("File uploaded successfully!");
  //       } catch (error) {
  //         console.error("Error uploading file:", error);
  //         alert("Failed to upload file. Please try again.");
  //       }
  //     }
  //   };
  //   input.click();
  // };

  // Handle folder creation
  // const handleCreateFolder = async () => {
  //   const folderName = prompt("Enter the name of the new folder:");
  //   // const documentMenuService = new DocumentMenuService(props.context);
  //   if (folderName) {
  //     try {
  //       await documentMenuService.addNewFolder(
  //         `${currentFolderPath}/${folderName}`
  //       );

  //       const newFolder: IDocumentItem = {
  //         Name: folderName,
  //         ServerRelativeUrl: `${currentFolderPath}/${folderName}`,
  //         items: [],
  //       };
  //       for (const folderItems of navigationStack) {
  //         const folder = folderItems.find(
  //           (item) => item.Name === currentFolderPath.split("/").pop()
  //         );
  //         if (folder) {
  //           folder.items = folder.items || [];
  //           folder.items.push(newFolder);
  //         }
  //       }
  //       setCurrentItems((prevItems) => [...prevItems, newFolder]);
  //       alert("Folder created successfully!");
  //     } catch (error) {
  //       console.error("Error creating folder:", error);
  //       alert("Failed to create folder. Please try again.");
  //     }
  //   }
  // };

  // Function to generate SharePoint file URL
  const getSharePointFileUrl = (serverRelativeUrl: string): string => {
    return `https://ccsdev01.sharepoint.com/:x:/r/sites/ProductDevelopment/_layouts/15/Doc.aspx?sourcedoc=${encodeURIComponent(
      serverRelativeUrl
    )}&action=default&mobileredirect=true`;
  };

  // Render breadcrumb navigation
  const renderBreadcrumb = () => {
    // If there's no navigation, don't show anything
    if (navigationStack.length === 1) {
      return <div />;
    }
    return (
      <div className={styles.BreadCrumbNavContainer}>
        {breadCrumbItems.map((item, index) => (
          <div key={index}>
            {index > 0 && " > "}
            <span
              className={
                index === breadCrumbItems.length - 1
                  ? styles.BreadCrumbItemActive
                  : styles.BreadCrumbItem
              }
              {...(index !== breadCrumbItems.length - 1
                ? { onClick: () => handleBreadcrumbClick(index) }
                : {})}
            >
              {item}
            </span>
          </div>
        ))}
      </div>
    );
  };

  // Handle breadcrumb click
  const handleBreadcrumbClick = (index: number) => {
    // Navigate to a specific folder in the breadcrumb
    if (index == 0) {
      if (props.layoutDropdownValue === "1") {
        setActiveIconLayout("icon");
      }
      pageCount.current = 0; // Reset page count when navigating to the root
      setBreadCrumbItems(["Documents"]);
      setCurrentItems(navigationStack[0]);
      setCurrentFolderPath(libraryName);
      setNavigationStack((prevStack) => prevStack.slice(0, 1));
      setPaginationStack([]); // Reset pagination stack
      return;
    } else {
      setCurrentItems(navigationStack[index]); // Get the items for the clicked level
      setBreadCrumbItems((prevItems) => prevItems.slice(0, index + 1)); // Keep only the levels up to the clicked breadcrumb
      setNavigationStack(
        (prevStack) => prevStack.slice(0, index + 1) // Keep only the levels up to the clicked breadcrumb
      );
      let segments = currentFolderPath.split("/");
      let new_index = segments.indexOf(breadCrumbItems[index + 1]);
      setCurrentFolderPath(segments.slice(0, new_index).join("/"));
      setPaginationStack([]); // Reset pagination stack
    }
  };

  // Debounced search function to slow down search requests and not skip
  const debouncedSearch = React.useRef(
    debounce((value: string, currentFolderPath: string) => {
      setLoading(true);

      if (!value || value.trim() === "") {
        if (props.layoutDropdownValue === "1") {
          documentMenuService
            .getLibraryData(currentFolderPath, 0)
            .then((data) => {
              setCurrentItems(data);
              setActiveIconLayout("icon");
            })
            .catch((error) =>
              console.error("Error fetching library data:", error)
            )
            .finally(() => setLoading(false));
          return;
        } else {
          documentMenuService
            .getLibraryData(currentFolderPath, 0)
            .then((data) => {
              setCurrentItems(data);
            })
            .catch((error) =>
              console.error("Error fetching library data:", error)
            )
            .finally(() => setLoading(false));
          return;
        }
      }

      if (props.layoutDropdownValue === "1") {
        documentMenuService
          .searchFilesAndFolders(value, currentFolderPath)
          .then((results) => {
            setActiveIconLayout("list");
            setCurrentItems(results);
          })
          .catch((error) => {
            console.error("Error searching files and folders:", error);
          })
          .finally(() => setLoading(false));
      } else {
        documentMenuService
          .searchFilesAndFolders(value, currentFolderPath)
          .then((results) => {
            setCurrentItems(results);
          })
          .catch((error) => {
            console.error("Error searching files and folders:", error);
          })
          .finally(() => setLoading(false));
      }
    }, 200)
  ).current;

  const handlesearchValue = (value: string) => {
    setSearchValue(value);
    debouncedSearch(value, currentFolderPath);
  };

  useEffect(() => {
    return () => {
      debouncedSearch.cancel();
    };
  }, [debouncedSearch]);

  const handleSwitchToListView = () => {
    setActiveIconLayout("list");
  };

  // console.log("currentFolderPath:", currentFolderPath);
  // // documentMenuService.getFieldsForUrl(props.documentLibraryUrl);
  // documentMenuService.getCategoryDistinctValues(
  //   props.documentLibraryUrl,
  //   "Department"
  // );
  // documentMenuService.getCategoryValueFiles(props.documentLibraryUrl, "HR");
  // console.log(props.layoutDropdownValue);
  // console.log("paginationStack:", paginationStack);
  // console.log("navigationStack:", navigationStack);
  // console.log("breadcrumbItems:", breadCrumbItems);
  console.log("currentItems:", currentItems);

  return (
    <div>
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
        {/* Search Field */}
        <div className={styles.SearchField}>
          <TextField
            styles={searchFieldStyles}
            placeholder="Search..."
            value={searchValue}
            onChange={(e, newValue) => {
              handlesearchValue(newValue || "");
              console.log(newValue);
            }}
          />
          {loading && (
            <div className={styles.LoadingSpinner}>
              <Spinner />
            </div>
          )}
        </div>

        {props.layoutDropdownValue === "2" ? (
          <div className={styles.ListTileLayoutButtons}>
            <button
              className={
                activeListTileLayout === "list"
                  ? styles.ListButtonActive
                  : styles.ListButtonInactive
              }
              onClick={() => setActiveListTileLayout("list")}
            ></button>
            <button
              className={
                activeListTileLayout === "tile"
                  ? styles.TileButtonActive
                  : styles.TileButtonInactive
              }
              onClick={() => setActiveListTileLayout("tile")}
            ></button>
          </div>
        ) : null}
      </div>
      <div className={styles.BreadCrumbNavSectionContainer}>
        {renderBreadcrumb()}
      </div>
      {props.layoutDropdownValue === "1" ? (
        activeIconLayout === "icon" && navigationStack.length === 1 ? (
          <IconView
            {...props}
            currentItems={currentItems}
            currentFolderPath={currentFolderPath}
            handleFolderClick={handleFolderClick}
            getSharePointFileUrl={getSharePointFileUrl}
            onSwitchToListView={handleSwitchToListView}
            handleNextFolderFileSet={handleNextFolderFileSet}
            handlePreviousFolderFileSet={handlePreviousFolderFileSet}
          />
        ) : (
          <ListView
            {...props}
            currentItems={currentItems}
            currentFolderPath={currentFolderPath}
            handleFolderClick={handleFolderClick}
            getSharePointFileUrl={getSharePointFileUrl}
          />
        )
      ) : (
        <>
          {activeListTileLayout === "tile" ? (
            <TileView
              {...props}
              currentItems={currentItems}
              currentFolderPath={currentFolderPath}
              handleFolderClick={handleFolderClick}
              getSharePointFileUrl={getSharePointFileUrl}
              handleNextFolderFileSet={handleNextFolderFileSet}
              handlePreviousFolderFileSet={handlePreviousFolderFileSet}
            />
          ) : (
            <ListView
              {...props}
              currentItems={currentItems}
              currentFolderPath={currentFolderPath}
              handleFolderClick={handleFolderClick}
              getSharePointFileUrl={getSharePointFileUrl}
            />
          )}
        </>
      )}
    </div>
  );
}
