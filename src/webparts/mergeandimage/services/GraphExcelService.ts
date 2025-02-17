import { MSGraphClient } from "@microsoft/sp-http";

export default class GraphExcelService {
  private graphClient: MSGraphClient;

  constructor(graphClient: MSGraphClient) {
    this.graphClient = graphClient;
  }

  public async copyMergedCells(): Promise<void> {
    try {
      const folderPath = "/Shared Documents/IT";
      const fileName = "2025 IT.xlsx";
      const worksheetName = "Jan 25 IT";
      const targetWorksheetName = "Copied Data";

      // 1. Get the file item ID
      const fileItem = await this.graphClient
        .api(`/sites/root/drive/root:${folderPath}/${fileName}`)
        .get();

      const fileId = fileItem.id;
      console.log("File ID:", fileId);

      // 2. Get the worksheet ID
      const worksheet = await this.graphClient
        .api(`/me/drive/items/${fileId}/workbook/worksheets/${worksheetName}`)
        .get();

      const worksheetId = worksheet.id;
      console.log("Worksheet ID:", worksheetId);

      // 3. Get all merged cells
      const range = await this.graphClient
        .api(`/me/drive/items/${fileId}/workbook/worksheets/${worksheetId}/range`)
        .get();

      const mergedCells = range.mergedCells;
      console.log("Merged Cells:", mergedCells);

      if (!mergedCells || mergedCells.length === 0) {
        console.log("No merged cells found.");
        return;
      }

      // 4. Check or Create the Second Worksheet
      let targetWorksheet = null;
      try {
        targetWorksheet = await this.graphClient
          .api(`/me/drive/items/${fileId}/workbook/worksheets/${targetWorksheetName}`)
          .get();
      } catch (error) {
        console.log("Target worksheet not found, creating a new one.");
        targetWorksheet = await this.graphClient
          .api(`/me/drive/items/${fileId}/workbook/worksheets`)
          .post({ name: targetWorksheetName });
      }

      const targetWorksheetId = targetWorksheet.id;

      // 5. Copy Merged Cells to the Second Worksheet
      for (const cell of mergedCells) {
        const cellAddress = cell.address;
        const cellValues = await this.graphClient
          .api(`/me/drive/items/${fileId}/workbook/worksheets/${worksheetId}/range(address='${cellAddress}')`)
          .get();

        // Paste data into the second worksheet
        await this.graphClient
          .api(`/me/drive/items/${fileId}/workbook/worksheets/${targetWorksheetId}/range(address='${cellAddress}')`)
          .patch({ values: cellValues.values });

        console.log(`Copied merged cell ${cellAddress}`);
      }

      console.log("All merged cells copied successfully.");
    } catch (error) {
      console.error("Error copying merged cells:", error);
    }
  }
}
