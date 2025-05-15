import * as React from "react";
import type {
  IDocumentMenuProps,
  IDocumentItem,
} from "../interfaces/IDocumentMenuProps";
import { useState, useEffect } from "react";
// import { BaseService } from "../../../common/services/BaseService";
import { DocumentMenuService } from "../services/DocumentMenuService";
import TileView from "../views/TileView";

export default function DocumentMenu(props: IDocumentMenuProps) {
  // const [libraryData, setLibraryData] = useState<IDocumentItem[]>([]);
  const [currentItems, setCurrentItems] = useState<IDocumentItem[]>([]); // Items to display at the current level
  const [navigationStack, setNavigationStack] = useState<IDocumentItem[][]>([]); // Stack to track navigation levels
  const [currentFolderPath, setCurrentFolderPath] = useState(
    "/sites/ProductDevelopment/Shared Documents"
  );
  // const [currentSearchItems, setCurrentSearchItems] = useState<IDocumentItem[]>([]);
  // const [showModal, setShowModal] = useState(false);
  const [breadCrumbItems, setBreadCrumbItems] = useState(["Documents"]);
  const [searchValue, setSearchValue] = useState("");

  const libraryName = props.documentUrl
    ? props.documentUrl
    : "/sites/ProductDevelopment/Shared Documents";
  const documentMenuService = new DocumentMenuService(props.context);
  // const baseService = new BaseService(props.context);
  // const pageCount = useRef(0);

  useEffect(() => {
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
      })
      .catch((error) => console.error("Error fetching library data:", error));

    // documentMenuService.searchFilesAndFolders();
  }, [props.context]);

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
              console.log("Fetched folder data:", currentItems);
            } catch (error) {
              console.error("Error fetching folder data:", error);
            }
          }
        }
      }
    };

    fetchFolderData();
  }, [currentItems]);

  //Handle next folder/file set
  // const handleNextFolderFileSet = async () => {
  //   pageCount.current += 3;
  //   documentMenuService
  //     .getLibraryData(currentFolderPath, pageCount.current)
  //     .then((data) => {
  //       console.log("Fetched library data:", data);
  //       // setLibraryData(data);
  //       setCurrentItems(data);
  //     })
  //     .catch((error) => console.error("Error fetching library data:", error));
  // };

  // Handle previous folder/file set
  // const handlePreviousFolderFileSet = async () => {
  //   if (pageCount.current === 0) {
  //     alert("No more previous folders or files.");
  //     return;
  //   }
  //   pageCount.current -= 3;
  //   documentMenuService
  //     .getLibraryData(currentFolderPath, pageCount.current)
  //     .then((data) => {
  //       console.log("Fetched library data:", data);
  //       // setLibraryData(data);
  //       setCurrentItems(data);
  //     })
  //     .catch((error) => console.error("Error fetching library data:", error));
  // };

  // Handle folder click
  const handleFolderClick = (folder: IDocumentItem) => {
    if (folder.items) {
      setNavigationStack((prevStack) => [...prevStack, currentItems]); // Push the current level to the stack
      setCurrentItems(folder.items);
      let newFolderName = folder.ServerRelativeUrl.split("/");
      setBreadCrumbItems((prevItems) => [
        ...prevItems,
        newFolderName[newFolderName.length - 1],
      ]); // Update breadcrumb items
      setCurrentFolderPath(folder.ServerRelativeUrl); // Set the current folder path
    }
  };

  // Handle back navigation
  const handleBackClick = () => {
    if (navigationStack.length > 0) {
      const previousLevel = navigationStack[navigationStack.length - 1]; // Get the previous level
      setNavigationStack((prevStack) => prevStack.slice(0, -1)); // Pop the last level from the stack
      setCurrentItems(previousLevel);
      setCurrentFolderPath(currentFolderPath.split("/").slice(0, -1).join("/")); // Update the current folder path
      setBreadCrumbItems((prevItems) => prevItems.slice(0, -1)); // Remove the last breadcrumb item
    }
  };

  // const retrieveNextLevelItems = async () => {
  //   try {

  //   } catch (error) { console.error("Error fetching subfolder items:", error);}}

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
    if (navigationStack.length === 0) {
      return <div />;
    }
    return (
      <React.Fragment>
        {breadCrumbItems.map((item, index) => (
          <React.Fragment key={index}>
            {index > 0 && " > "}
            <span onClick={() => handleBreadcrumbClick(index)}>{item}</span>
          </React.Fragment>
        ))}
      </React.Fragment>
    );
  };

  // Handle breadcrumb click
  const handleBreadcrumbClick = (index: number) => {
    // Navigate to a specific folder in the breadcrumb
    if (index == 0) {
      setBreadCrumbItems(["Documents"]);
      setNavigationStack([]);
      setCurrentItems(navigationStack[0]);
      setCurrentFolderPath(libraryName);
      return;
    } else {
      setCurrentItems(navigationStack[index]); // Get the items for the clicked level
      setBreadCrumbItems((prevItems) => prevItems.slice(0, index + 1)); // Keep only the levels up to the clicked breadcrumb
      setNavigationStack(
        (prevStack) => prevStack.slice(0, index) // Keep only the levels up to the clicked breadcrumb
      );
      let segments = currentFolderPath.split("/");
      let new_index = segments.indexOf(breadCrumbItems[index + 1]);
      setCurrentFolderPath(segments.slice(0, new_index).join("/"));
    }
  };

  const handlesearchValue = (value: string) => {
    setSearchValue(value);

    if (!value) {
      // If the search value is empty, reset to the original items
      documentMenuService
        .getLibraryData(currentFolderPath, 0)
        .then((data) => {
          setCurrentItems(data);
        })
        .catch((error) => console.error("Error fetching library data:", error));
      return;
    }

    // Perform search
    documentMenuService
      .searchFilesAndFolders(value, currentFolderPath)
      .then((results) => {
        setCurrentItems(results);
      })
      .catch((error) => {
        console.error("Error searching files and folders:", error);
      });

    // const filteredItems = currentItems.filter((item) =>
    //   item.Name.toLowerCase().includes(value.toLowerCase())
    // );
    // setCurrentItems(filteredItems);
  };

  return (
    <div>
      <TileView
        {...props}
        currentItems={currentItems}
        currentFolderPath={currentFolderPath}
        navigationStack={navigationStack}
        handleFolderClick={handleFolderClick}
        getSharePointFileUrl={getSharePointFileUrl}
        searchValue={searchValue}
        handlesearchValue={handlesearchValue}
        handleBackClick={handleBackClick}
        renderBreadcrumb={renderBreadcrumb}
      />
    </div>
  );
}
