import * as React from "react";
import type {
  IDocumentMenuProps,
  IDocumentItem,
} from "../interfaces/IDocumentMenuProps";
import { useState, useEffect } from "react";
import { BaseService } from "../../../common/services/BaseService";
import { DocumentMenuService } from "../services/DocumentMenuService";

export default function DocumentMenu(props: IDocumentMenuProps) {
  // const [libraryData, setLibraryData] = useState<IDocumentItem[]>([]);
  const [currentItems, setCurrentItems] = useState<IDocumentItem[]>([]); // Items to display at the current level
  const [navigationStack, setNavigationStack] = useState<IDocumentItem[][]>([]); // Stack to track navigation levels
  const [currentFolderPath, setCurrentFolderPath] = useState(
    "/sites/ProductDevelopment/Shared Documents"
  );
  const [showModal, setShowModal] = useState(false);

  const libraryName = props.documentUrl
    ? props.documentUrl
    : "/sites/ProductDevelopment/Shared Documents";
  const documentMenuService = new DocumentMenuService(props.context);
  const baseService = new BaseService(props.context);

  useEffect(() => {
    //Get all Folder information
    baseService
      .getAllFoldersAndFiles(libraryName)
      .then((data) => {
        console.log("Fetched library data:", data);
        // setLibraryData(data);
        setCurrentItems(data);
      })
      .catch((error) => console.error("Error fetching library data:", error));
  }, [props.context]);

  // Handle folder click
  const handleFolderClick = (folder: IDocumentItem) => {
    if (folder.items) {
      setNavigationStack((prevStack) => [...prevStack, currentItems]); // Push the current level to the stack
      setCurrentItems(folder.items);
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
    }
  };

  // Handle file creation
  const handleCreateFile = async () => {
    let fileName = prompt("Enter the name of the new file:");
    fileName += ".docx";
    if (fileName) {
      try {
        documentMenuService.addNewFile(currentFolderPath, fileName);

        const newFile: IDocumentItem = {
          Name: fileName,
          ServerRelativeUrl: `${currentFolderPath}/${fileName}`,
        };

        setCurrentItems((prevItems) => [...prevItems, newFile]);
        alert("File created successfully!");
      } catch (error) {
        console.error("Error creating file:", error);
        alert("Failed to create file. Please try again.");
      }
    } else {
      alert("Invalid file name. Only .docx files are supported for creation.");
    }
  };

  // Handle file upload
  const handleUploadFile = async () => {
    const input = document.createElement("input");
    input.type = "file";

    const allowedExtensions = [
      ".doc",
      ".docx",
      ".xls",
      ".xlsx",
      ".ppt",
      ".pptx",
      ".pdf",
      ".txt",
      ".csv",
      ".one",
      ".vsd",
      ".vsdx",
    ];

    // Set accept attribute to show only these files in dialog (optional)
    input.accept = allowedExtensions
      .map((ext) => `${ext},.${ext.toUpperCase()}`)
      .join(",");

    input.onchange = async (event: any) => {
      const file = event.target.files[0];
      if (file) {
        try {
          // Get file extension
          const fileExtension = file.name.split(".").pop().toLowerCase();

          // Validate file type
          if (!allowedExtensions.includes(`.${fileExtension}`)) {
            alert(
              `Invalid file type. Please upload only Microsoft Office files or PDFs. Allowed formats: ${allowedExtensions.join(
                ", "
              )}`
            );
            return;
          }

          await baseService.uploadDocument(currentFolderPath, file.name, file);

          const newFile: IDocumentItem = {
            Name: file.name,
            ServerRelativeUrl: `${currentFolderPath}/${file.name}`,
          };

          setCurrentItems((prevItems) => [...prevItems, newFile]);
          alert("File uploaded successfully!");
        } catch (error) {
          console.error("Error uploading file:", error);
          alert("Failed to upload file. Please try again.");
        }
      }
    };
    input.click();
  };

  // Handle folder creation
  const handleCreateFolder = async () => {
    const folderName = prompt("Enter the name of the new folder:");
    // const documentMenuService = new DocumentMenuService(props.context);
    if (folderName) {
      try {
        await documentMenuService.addNewFolder(
          `${currentFolderPath}/${folderName}`
        );

        const newFolder: IDocumentItem = {
          Name: folderName,
          ServerRelativeUrl: `${currentFolderPath}/${folderName}`,
          items: [],
        };

        setCurrentItems((prevItems) => [...prevItems, newFolder]);
        alert("Folder created successfully!");
      } catch (error) {
        console.error("Error creating folder:", error);
        alert("Failed to create folder. Please try again.");
      }
    }
  };

  // Function to generate SharePoint file URL
  const getSharePointFileUrl = (serverRelativeUrl: string): string => {
    return `https://ccsdev01.sharepoint.com/:x:/r/sites/ProductDevelopment/_layouts/15/Doc.aspx?sourcedoc=${encodeURIComponent(
      serverRelativeUrl
    )}&action=default&mobileredirect=true`;
  };

  // Function to count files in a folder and its nested folders
  const countFilesInFolder = (folder: IDocumentItem): number => {
    if (!folder.items || folder.items.length === 0) {
      return 0;
    }

    return folder.items.reduce((count, item) => {
      if (item.items) {
        return count + countFilesInFolder(item);
      } else {
        return count + 1;
      }
    }, 0);
  };

  const renderItems = (items: IDocumentItem[]) => {
    console.log(currentItems);
    console.log(currentFolderPath);
    console.log(navigationStack);
    return (
      <div style={{ display: "flex", flexWrap: "wrap", gap: "10px" }}>
        {items.map((item, index) => (
          <div
            key={index}
            style={{
              border: "1px solid #ccc",
              borderRadius: "5px",
              padding: "10px",
              textAlign: "center",
              width: "150px",
              cursor: "pointer",
              backgroundColor: item.items ? "#f0f8ff" : "#fff",
            }}
            onClick={() => item.items && handleFolderClick(item)}
          >
            {item.items ? (
              <>
                <div style={{ fontSize: "24px" }}>📁</div>
                <div>
                  {item.Name} ({countFilesInFolder(item)} files)
                </div>
              </>
            ) : (
              <a
                href={getSharePointFileUrl(item.ServerRelativeUrl)}
                target="_blank"
                rel="noopener noreferrer"
                style={{ textDecoration: "none", color: "inherit" }}
              >
                <div style={{ fontSize: "24px" }}>📄</div>
                <div>{item.Name}</div>
              </a>
            )}
          </div>
        ))}
      </div>
    );
  };

  // Modal for file/folder creation
  const renderModal = () => {
    return (
      <div
        style={{
          position: "fixed",
          top: "0",
          left: "0",
          width: "100%",
          height: "100%",
          backgroundColor: "rgba(0, 0, 0, 0.5)",
          display: "flex",
          justifyContent: "center",
          alignItems: "center",
          zIndex: 1000,
        }}
      >
        <div
          style={{
            background: "white",
            padding: "20px",
            borderRadius: "8px",
            textAlign: "center",
            width: "300px",
          }}
        >
          <h3>Select an Action</h3>
          <button
            onClick={() => {
              setShowModal(false);
              handleCreateFile();
            }}
            style={{
              margin: "10px",
              background: "#4CAF50",
              color: "white",
              border: "none",
              padding: "10px 15px",
              cursor: "pointer",
              borderRadius: "5px",
            }}
          >
            ➕ Create File
          </button>
          <button
            onClick={() => {
              setShowModal(false);
              handleUploadFile();
            }}
            style={{
              margin: "10px",
              background: "#2196F3",
              color: "white",
              border: "none",
              padding: "10px 15px",
              cursor: "pointer",
              borderRadius: "5px",
            }}
          >
            📤 Upload File
          </button>
          <button
            onClick={() => setShowModal(false)}
            style={{
              marginTop: "10px",
              background: "#f44336",
              color: "white",
              border: "none",
              padding: "10px 15px",
              cursor: "pointer",
              borderRadius: "5px",
            }}
          >
            ❌ Cancel
          </button>
        </div>
      </div>
    );
  };

  return (
    <div>
      <h2>Document Library Contents</h2>
      <div style={{ marginBottom: "10px" }}>
        <button
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
            background: "#2196F3",
            color: "white",
            border: "none",
            padding: "10px 15px",
            cursor: "pointer",
            borderRadius: "5px",
          }}
        >
          ➕ Create Folder
        </button>
      </div>
      {showModal && renderModal()}
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
      {renderItems(currentItems)}
    </div>
  );
}
