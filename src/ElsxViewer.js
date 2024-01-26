// Done By Sastik Kumar Das
// mob: +91 9679482955

import React, { useState } from "react";
import * as XLSX from "xlsx";
import {
  AiOutlineZoomIn,
  AiOutlineZoomOut,
  AiOutlinePrinter,
  AiOutlineDownload,
  AiOutlineFullscreen,
  AiOutlinePlusCircle,
  AiOutlineFullscreenExit,
} from "react-icons/ai";

const ElsxViewer = () => {
  const [excelFile, setExcelFile] = useState(null);
  const [excelData, setExcelData] = useState(null);
  const [zoomScale, setZoomScale] = useState(1);
  const [isFullScreen, setIsFullScreen] = useState(false);

  const handleFile = (e) => {
    let selectedFile = e.target.files[0];
    let fileTypes = [
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ];
    if (selectedFile) {
      if (fileTypes.includes(selectedFile.type)) {
        let reader = new FileReader();
        reader.onload = (e) => {
          setExcelFile(e.target.result);
          const excelData = processExcelFile(e.target.result);
          setExcelData(excelData);
        };
        reader.readAsArrayBuffer(selectedFile);
      }
    } else {
      console.log("Please select your file...");
    }
  };
  const processExcelFile = (fileBuffer) => {
    const workbook = XLSX.read(fileBuffer, { type: "buffer" });
    const worksheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[worksheetName];
    const data = XLSX.utils.sheet_to_json(worksheet);
    return data.slice(0, 10);
  };

  const handleZoomIn = () => {
    const newZoomScale = zoomScale + 0.2;
    setZoomScale(newZoomScale);
    updateZoom(newZoomScale);
  };

  const handleZoomOut = () => {
    const newZoomScale = zoomScale - 0.2;
    setZoomScale(newZoomScale);
    updateZoom(newZoomScale);
  };

  const updateZoom = (scale) => {
    const excelContent = document.querySelector(".excel-content");
    if (excelContent) {
      excelContent.style.transform = `scale(${scale})`;
    }
  };

  const handlePrint = () => {
    if (excelData) {
      const printWindow = window.open("", "_blanck");
      printWindow.document.write(`
        <html>
          <head>
            <title>Print</title>
            <style>
              body {
                font-family: Arial, sans-serif;
              }
              table {
                border-collapse: collapse;
                width: 100%;
              }
              th, td {
                border: 1px solid #dddddd;
                text-align: left;
                padding: 8px;
              }
              th {
                background-color: #f2f2f2;
              }
            </style>
          </head>
          <body>
            <div class="excel-sheet">
              <table class="table">
                <thead>
                  <tr>
                    ${Object.keys(excelData[0])
                      .map((key) => `<th>${key}</th>`)
                      .join("")}
                  </tr>
                </thead>
                <tbody>
                  ${excelData
                    .map(
                      (individualExcelData, index) => `
                    <tr>
                      ${Object.keys(individualExcelData)
                        .map((key) => `<td>${individualExcelData[key]}</td>`)
                        .join("")}
                    </tr>`
                    )
                    .join("")}
                </tbody>
              </table>
            </div>
          </body>
        </html>
      `);
      printWindow.document.close();
      printWindow.print();
      printWindow.onafterprint = () => printWindow.close();
    }
  };

  const handleDownload = () => {
    if (excelData) {
      const fileName = prompt("Enter file name") || "downloaded_file";
      const blob = new Blob([excelFile], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const link = document.createElement("a");
      link.href = window.URL.createObjectURL(blob);
      link.download = `${fileName}.xlsx`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    }
  };

  const handleFullScreen = () => {
    const docElement = document.documentElement;

    if (!isFullScreen) {
      if (docElement.requestFullscreen) {
        docElement.requestFullscreen();
      } else if (docElement.mozRequestFullScreen) {
        docElement.mozRequestFullScreen();
      } else if (docElement.webkitRequestFullscreen) {
        docElement.webkitRequestFullscreen();
      } else if (docElement.msRequestFullscreen) {
        docElement.msRequestFullscreen();
      }
    } else {
      if (document.exitFullscreen) {
        document.exitFullscreen();
      } else if (document.mozCancelFullScreen) {
        document.mozCancelFullScreen();
      } else if (document.webkitExitFullscreen) {
        document.webkitExitFullscreen();
      } else if (document.msExitFullscreen) {
        document.msExitFullscreen();
      }
    }
    setIsFullScreen(!isFullScreen);
  };

  return (
    <>
      <div className={`docx-viewer ${isFullScreen ? "full-screen" : ""}`}>
        <div className="toolbar">
          <button onClick={() => handleZoomIn(1.2)}>
            <AiOutlineZoomIn size={17} />
          </button>
          <button onClick={() => handleZoomOut(0.8)}>
            <AiOutlineZoomOut size={17} />
          </button>
          <button onClick={handleFullScreen}>
            {isFullScreen ? (
              <AiOutlineFullscreenExit size={17} />
            ) : (
              <AiOutlineFullscreen size={17} />
            )}
          </button>
          <button onClick={handlePrint}>
            <AiOutlinePrinter size={17} />
          </button>
          <button onClick={handleDownload}>
            <AiOutlineDownload size={17} />
          </button>

          <label
            htmlFor="file-upload"
            className="file-upload-label form-control"
          >
            <input
              type="file"
              id="file-upload"
              accept=".xlsx"
              onChange={handleFile}
            />
            <AiOutlinePlusCircle size={25} />
          </label>
        </div>
        <div className="excel-content">
          {excelData ? (
            <div className="excel-sheet">
              <table className="table">
                <thead>
                  <tr>
                    {Object.keys(excelData[0]).map((key) => (
                      <th key={key}>{key}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {excelData.map((individualExcelData, index) => (
                    <tr key={index}>
                      {Object.keys(individualExcelData).map((key) => (
                        <td key={key}>{individualExcelData[key]}</td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : (
            <div className="uploaded-file">No file is uploaded</div>
          )}
        </div>
      </div>
    </>
  );
};

export default ElsxViewer;
