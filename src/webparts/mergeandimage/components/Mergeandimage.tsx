import * as React from "react";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { PrimaryButton } from "@fluentui/react/lib/Button";

export interface IMergeandimageProps {
  description: string;
  graphClient: MSGraphClientV3;
}

const Mergeandimage: React.FC<IMergeandimageProps> = ({ description, graphClient }) => {
  const folderPath = "/Shared Documents/IT/2025 IT.xlsx";
  const worksheetName = "Jan 25 IT";

  const handleButtonClick = async () => {
    console.log("Button Clicked!");

    try {
      // Step 1: Get Drive ID
      const driveResponse = await graphClient.api("/me/drive/root:/").get();
      const driveId = driveResponse.id;
      console.log(`Drive ID: ${driveId}`);

      // Step 2: Get File ID
      const fileResponse = await graphClient.api(`/me/drive/root:${folderPath}`).get();
      const fileId = fileResponse.id;
      console.log(`File ID: ${fileId}`);

      // Step 3: Get Worksheets
      const worksheetsResponse = await graphClient
        .api(`/me/drive/items/${fileId}/workbook/worksheets`)
        .get();
      const worksheets = worksheetsResponse.value;
      console.log("Worksheets:", worksheets);

      // Step 4: Find "Jan 25 IT" Worksheet
      const worksheet = worksheets.find((ws: { name: string }) => ws.name === worksheetName);
      if (!worksheet) {
        console.error(`Worksheet "${worksheetName}" not found.`);
        return;
      }
      const worksheetId = worksheet.id;
      console.log(`Found worksheet "${worksheetName}" with ID: ${worksheetId}`);

      // Step 5: Get Second Worksheet
      const secondWorksheet = worksheets.find((ws: { name: string }) => ws.name !== worksheetName);
      if (!secondWorksheet) {
        console.error("No second worksheet found.");
        return;
      }
      const secondWorksheetId = secondWorksheet.id;
      console.log(`Second Worksheet ID: ${secondWorksheetId}`);

      // Step 6: Get Merged Cells from First Worksheet
      const mergedCellsResponse = await graphClient
        .api(`/me/drive/items/${fileId}/workbook/worksheets/${worksheetId}/range/mergedCells`)
        .get();
      const mergedCells = mergedCellsResponse.value;
      console.log("Merged Cells:", mergedCells);

      if (!mergedCells || mergedCells.length === 0) {
        console.warn("No merged cells found in the worksheet.");
        return;
      }

      // Step 7: Paste Merged Cells into the Second Worksheet
      for (const cell of mergedCells) {
        const rangeAddress = cell.address;
        console.log(`Copying merged cell: ${rangeAddress}`);

        await graphClient
          .api(`/me/drive/items/${fileId}/workbook/worksheets/${secondWorksheetId}/range(address='${rangeAddress}')`)
          .patch({ values: cell.values });

        console.log(`Copied ${rangeAddress} to second worksheet.`);
      }

      console.log("Merge cell transfer completed.");
    } catch (error) {
      console.error("Error:", error);
    }
  };

  return (
    <div>
      <h2>{description}</h2>
      <PrimaryButton text="Transfer Merged Cells" onClick={handleButtonClick} />
    </div>
  );
};

export default Mergeandimage;
