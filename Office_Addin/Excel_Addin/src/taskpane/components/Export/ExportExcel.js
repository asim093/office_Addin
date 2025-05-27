import React, { useEffect, useState } from "react";
import "./ExcelImport.scss";
import DismissIcon from "../assets/images/Dismiss.png";
import { jwtDecode } from "jwt-decode";

import IconCheck from "../assets/images/IconCheck.png";
import UserIcon from "../assets/images/user.png";
import NeedHelpIcon from "../assets/images/needHelp.png";
import Arrow_white_Import from "../assets/images/Arrow_white_Import.png";
import UpdateIcon from "../assets/images/update.png";
import TableSearch from "../assets/images/TableSearch.png";
import MultiSelectDropdown from "../MultiSelectDropdown/MultiSelectDropdown";
import CustomTabs from "../Tabs/CustomTabs";
import ExpandableTable from "../ExpandableTable/ExpandableTable";
import InsertIcon from "../assets/images/InsertIcon.png";
import ArrowExport from "../assets/images/ArrowExport.png";
import Checkactive from "../assets/images/checkactive.png";
import Checkhover from "../assets/images/checkhover.png";
import { useNavigate } from "react-router-dom";

const ExportExcel = () => {
  const [expandedRow, setExpandedRow] = useState(null);
  const [checked, setChecked] = useState(false);
  const [loading, setLoading] = useState(false);
  const [Rangedata, SetRangedata] = useState([]);
  const [showpopup, SetShowpopup] = useState(false);
  const [isUpdating, setIsUpdating] = useState(false);
  const navigate = useNavigate();

  const handleChange = () => {
    setLoading(true);
    setTimeout(() => {
      setLoading(false);
      navigate("/SignOutUser");
    }, 2000);
  };

  // Updated function to extract any named range as HTML
  const extractNamedRangeAsHtml = async (rangeName, context) => {
    try {
      const namedRange = context.workbook.names.getItem(rangeName).getRange();
      namedRange.load(["text", "rowCount", "columnCount", "values"]);
      await context.sync();

      const rowCount = namedRange.rowCount;
      const columnCount = namedRange.columnCount;
      const textValues = namedRange.text;
      const values = namedRange.values;

      let html = `<table style="border-collapse: collapse;">`;

      for (let i = 0; i < rowCount; i++) {
        html += "<tr>";
        for (let j = 0; j < columnCount; j++) {
          const cell = namedRange.getCell(i, j);
          cell.format.load(["fill/color", "font/color", "font/bold", "font/size", "font/name"]);
          await context.sync();

          const value = textValues[i][j];
          const bgColor = cell.format.fill.color || "#ffffff";
          const fontColor = cell.format.font.color || "#000000";
          const fontSize = cell.format.font.size || 12;
          const fontName = cell.format.font.name || "Arial";
          const bold = cell.format.font.bold ? "bold" : "normal";

          html += `<td style="
            border: 1px solid #000;
            padding: 6px;
            background-color: ${bgColor};
            color: ${fontColor};
            font-size: ${fontSize}px;
            font-family: ${fontName};
            font-weight: ${bold};
          ">${value}</td>`;
        }
        html += "</tr>";
      }

      html += "</table>";

      // Return both HTML and table values
      return {
        html: html,
        value: values.flat().join(", "),
        rowCount: rowCount,
        columnCount: columnCount
      };

    } catch (error) {
      console.error(`Error extracting HTML for range ${rangeName}:`, error);
      return null;
    }
  };

  useEffect(() => {
    getDetectRange()
  }, []);

  const getRanges = async () => {
    try {
      Office.context.auth.getAccessTokenAsync(
        {
          allowConsentPrompt: true,
          allowSignInPrompt: true,
          forMSGraphAccess: true,
        },
        async (result) => {
          console.log("Token callback result:", result);

          if (result.status === "succeeded" && result.value) {
            const decodedToken = jwtDecode(result.value);
            const email = decodedToken.preferred_username;
            console.log("Email:", email);

            await Excel.run(async (context) => {
              const workbook = context.workbook;
              workbook.load("name");
              await context.sync();

              const fileName = workbook.name;

              const names = workbook.names;
              names.load("items/name,items/value");
              await context.sync();

              for (const namedItem of names.items) {
                let payload = {};

                if (namedItem.name.endsWith("_img")) {
                  // Handle image ranges
                  const range = workbook.names.getItem(namedItem.name).getRange();
                  await context.sync();
                  const image = range.getImage();
                  await context.sync();

                  payload = {
                    user: email,
                    rangeName: namedItem.name,
                    value: image.value,
                    type: "image",
                    fileName,
                  };
                } else {
                  // Handle text/table ranges
                  const range = workbook.worksheets
                    .getActiveWorksheet()
                    .getRange(namedItem.value);
                  range.load(["values", "rowCount", "columnCount"]);
                  await context.sync();

                  const isTable = range.rowCount > 1 || range.columnCount > 1;

                  if (isTable) {
                    // For tables, extract HTML with formatting and values
                    const htmlData = await extractNamedRangeAsHtml(namedItem.name, context);

                    if (htmlData) {
                      payload = {
                        user: email,
                        rangeName: namedItem.name,
                        value: htmlData.value, // Table values as comma-separated string
                        html: htmlData.html,   // HTML with formatting
                        type: "table",
                        fileName,
                        rowCount: htmlData.rowCount,
                        columnCount: htmlData.columnCount,
                      };
                    } else {
                      // Fallback if HTML extraction fails
                      const value = range.values.flat().join(", ");
                      payload = {
                        user: email,
                        rangeName: namedItem.name,
                        value: value,
                        type: "table",
                        fileName,
                        rowCount: range.rowCount,
                        columnCount: range.columnCount,
                      };
                    }
                  } else {
                    // For single cells/text
                    const value = range.values.flat().join(", ");
                    payload = {
                      user: email,
                      rangeName: namedItem.name,
                      value: value,
                      type: "text",
                      fileName,
                      rowCount: range.rowCount,
                      columnCount: range.columnCount,
                    };
                  }
                }

                await uploadData(payload);
                console.log(`Processed range: ${namedItem.name}, Type: ${payload.type}`);
              }

              // Handle shapes/images
              const sheet = context.workbook.worksheets.getActiveWorksheet();
              const shapes = sheet.shapes;
              shapes.load("items/name,type");
              await context.sync();

              const imageShapes = shapes.items.filter(shape => shape.type === "Image");

              for (const shape of imageShapes) {
                try {
                  const image = shape.getAsImage(Excel.PictureFormat.png);
                  await context.sync();

                  const imagePayload = {
                    user: email,
                    rangeName: shape.name,
                    value: image.value,
                    type: "image",
                    fileName,
                  };

                  await uploadData(imagePayload);
                } catch (error) {
                  console.error(`❌ Error extracting image '${shape.name}':`, error);
                }
              }

              console.log("All ranges processed successfully!");
            });
          }
        }
      );
    } catch (error) {
      console.error("Error during Get Ranges:", error);
    }
  };

  const uploadData = async (payload) => {
    try {
      const existingData = JSON.parse(localStorage.getItem("ExcelData")) || [];

      const updatedData = existingData.map((item) => {
        if (item.rangeName === payload.rangeName && item.fileName === payload.fileName) {
          return { ...item, ...payload };
        }
        return item;
      });

      const isNew = !existingData.some(
        (item) => item.rangeName === payload.rangeName && item.fileName === payload.fileName
      );

      if (isNew) {
        updatedData.push(payload);
      }

      localStorage.setItem("ExcelData", JSON.stringify(updatedData));
      console.log(`Payload stored/updated in localStorage: ${payload.rangeName}`);
    } catch (error) {
      console.error("Failed to store/update payload in localStorage:", error);
    }
  };

  const getDetectRange = async () => {
    try {
      await Excel.run(async (context) => {
        const workbook = context.workbook;
        workbook.load("name");
        await context.sync();

        const fileName = workbook.name;

        const names = workbook.names;
        names.load("items/name,items/value");
        await context.sync();

        const payload = [];

        for (const namedItem of names.items) {
          const range = workbook.worksheets
            .getActiveWorksheet()
            .getRange(namedItem.value);
          range.load(["values", "rowCount", "columnCount"]);
          await context.sync();

          const value = range.values.flat().join(", ");
          payload.push({
            id: namedItem.name,
            Sheet: namedItem.value,
            RangeName: namedItem.name,
            value: value,
            filename: fileName,
            type: namedItem.name.endsWith("_img")
              ? "Image"
              : range.rowCount > 1 || range.columnCount > 1
                ? "Table"
                : "Text",
          });
        }

        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.load("name");
        const shapes = sheet.shapes;
        shapes.load("items");
        await context.sync();

        const imageShapes = shapes.items.filter((shape) => shape.type === "Image");

        for (const shape of imageShapes) {
          try {
            shape.load("name");
            const image = shape.getAsImage(Excel.PictureFormat.png);
            await context.sync();

            payload.push({
              id: shape.name,
              Sheet: sheet.name,
              RangeName: shape.name,
              value: image.value,
              filename: fileName,
              type: "Image",
            });
          } catch (error) {
            console.error(`❌ Error extracting image '${shape.name}':`, error);
          }
        }

        SetRangedata(payload);
      });
    } catch (error) {
      console.error("Error in getDetectRange:", error);
    }
  };

  const toggleRow = (index) => {
    setExpandedRow(expandedRow === index ? null : index);
  };

  return (
    <div className="excel-import-container">
      {loading ? (
        <div className="main-loading-container">
          <div className="loading-container">
            <div className="spinner"></div>
            <p className="loading-text">Loading...</p>
          </div>
        </div>
      ) : (
        <>
          {/* User Info */}
          {showpopup && (
            <div className="success-message">
              <img src={IconCheck} alt="Success" />
              <span className="LoggedSuccessfully">Export Data successfully</span>
              <img
                src={DismissIcon}
                alt="Dismiss"
                className="dismiss-icon"
                onClick={() => SetShowpopup(false)}
              />
            </div>
          )}

          {/* Import Excel Content */}
          <div className="import-section">
            <h2>Choose content to export</h2>
            <p>Choose export sources</p>

            <button className="update-btn" onClick={getDetectRange}>
              <img src={TableSearch} alt="Update" />
              Detect Ranges
            </button>
            <p className="last-update">Last update: 14/02/2025, 15:29:31</p>
          </div>

          {Rangedata.length > 0 && (
            <ExpandableTable
              className="disabled_image"
              Rangedata={Rangedata}
              source=""
              headingfirst="Export list"
            />
          )}

          <button className="insert-button" onClick={getRanges}>
            <img src={ArrowExport} alt="Insert Icon" />
            Export
          </button>
        </>
      )}
    </div>
  );
}

export default ExportExcel;