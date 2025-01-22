import 'package:excel/excel.dart';

class ExcelGenerator {
  // Excel For Daily Work Center Report ..................................................
  void createExcel(
      List data, factoryName, floorName, date, int totalDGrCumQty) {
    var excel = Excel.createExcel();
    var sheet = excel["Sheet1"];

    // --- Style Setup ---
    var headerStyle = CellStyle(
      textWrapping: TextWrapping.WrapText,
      bold: true,
      italic: true,
      fontSize: 12,
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
      rotation: 90,
      leftBorder: Border(borderStyle: BorderStyle.Thin),
      rightBorder: Border(borderStyle: BorderStyle.Thin),
      bottomBorder: Border(borderStyle: BorderStyle.Thin),
      topBorder: Border(borderStyle: BorderStyle.Thin),
    );

    var centerStyle = CellStyle(
      horizontalAlign: HorizontalAlign.Center,
      leftBorder: Border(borderStyle: BorderStyle.Thin),
      rightBorder: Border(borderStyle: BorderStyle.Thin),
      bottomBorder: Border(borderStyle: BorderStyle.Thin),
      topBorder: Border(borderStyle: BorderStyle.Thin),
    );

    var verCenterStyle = CellStyle(
      verticalAlign: VerticalAlign.Center,
    );
    var pageTitleStyle = CellStyle(
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
      bold: true,
      fontSize: 12,
    );
    var signStyle = CellStyle(
      horizontalAlign: HorizontalAlign.Left,
      verticalAlign: VerticalAlign.Bottom,
      bold: true,
      fontSize: 12,
    );

    var signStyle2 = CellStyle(
      horizontalAlign: HorizontalAlign.Left,
      verticalAlign: VerticalAlign.Bottom,
      bold: true,
    );

    var commonStyle = CellStyle(
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
      textWrapping: TextWrapping.WrapText,
      fontSize: 10,
      leftBorder: Border(borderStyle: BorderStyle.Thin),
      rightBorder: Border(borderStyle: BorderStyle.Thin),
      bottomBorder: Border(borderStyle: BorderStyle.Thin),
      topBorder: Border(borderStyle: BorderStyle.Thin),
    );

    var rotateStyle = CellStyle(
      rotation: 90,
      leftBorder: Border(borderStyle: BorderStyle.Thin),
      rightBorder: Border(borderStyle: BorderStyle.Thin),
      bottomBorder: Border(borderStyle: BorderStyle.Thin),
      topBorder: Border(borderStyle: BorderStyle.Thin),
    );

    sheet.setRowHeight(1, 45);
    sheet.setRowHeight(0, 30);

    sheet.setColumnWidth(0, 6);
    sheet.setColumnWidth(1, 7);
    sheet.setColumnWidth(2, 6);
    sheet.setColumnWidth(3, 11);
    sheet.setColumnWidth(4, 6);
    sheet.setColumnWidth(5, 7);
    sheet.setColumnWidth(6, 11);
    sheet.setColumnWidth(7, 7);
    sheet.setColumnWidth(8, 10);
    sheet.setColumnWidth(9, 11);
    sheet.setColumnWidth(10, 6);
    sheet.setColumnWidth(11, 6);
    sheet.setColumnWidth(12, 9);
    sheet.setColumnWidth(13, 6);
    sheet.setColumnWidth(14, 6);
    sheet.setColumnWidth(15, 6);
    sheet.setColumnWidth(16, 6);
    sheet.setColumnWidth(17, 6);
    sheet.setColumnWidth(18, 6);
    sheet.setColumnWidth(19, 6);
    sheet.setColumnWidth(20, 6);
    sheet.setColumnWidth(21, 7);
    sheet.setColumnWidth(22, 4);
    sheet.setColumnWidth(23, 5);
    sheet.setColumnWidth(24, 7);
    sheet.setColumnWidth(25, 7);
    sheet.setColumnWidth(26, 7);
    sheet.setColumnWidth(27, 7);
    sheet.setColumnWidth(28, 7);

    // Row 1: Header Titles
    // ----------------------- Factory -------------------------- //
    sheet.merge(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0),
        CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: 0));
    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0)).value =
        TextCellValue("Factory: ${factoryName}");

    sheet
        .cell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0))
        .cellStyle = verCenterStyle;

    // ----------------------- Floor -------------------------- //
    sheet.merge(CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: 0),
        CellIndex.indexByColumnRow(columnIndex: 5, rowIndex: 0));
    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: 0)).value =
        TextCellValue("Floor: ${floorName}");
    sheet
        .cell(CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: 0))
        .cellStyle = verCenterStyle;

    // ----------------------- Title -------------------------- //
    sheet.merge(CellIndex.indexByColumnRow(columnIndex: 6, rowIndex: 0),
        CellIndex.indexByColumnRow(columnIndex: 24, rowIndex: 0));
    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 6, rowIndex: 0)).value =
        TextCellValue("Daily Work Center Report");
    sheet
        .cell(CellIndex.indexByColumnRow(columnIndex: 6, rowIndex: 0))
        .cellStyle = pageTitleStyle;

    // ----------------------- Date -------------------------- //
    sheet.merge(CellIndex.indexByColumnRow(columnIndex: 25, rowIndex: 0),
        CellIndex.indexByColumnRow(columnIndex: 27, rowIndex: 0));
    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 25, rowIndex: 0)).value =
        TextCellValue("Date: ${date}");
    sheet
        .cell(CellIndex.indexByColumnRow(columnIndex: 25, rowIndex: 0))
        .cellStyle = verCenterStyle;

    List header = [
      "Buyer",
      "Style",
      "Buy",
      "Sales Order No",
      "Item",
      "FG(Mat No)",
      "Our No(F.R)",
      "Color Name",
      "Country Name",
      "Prod Order No",
      "PO No",
      "Order/Issue",
      "Size Information",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "D GR Qty",
      "Total GR Cum Qty",
      "Balance GR Qty",
      "Last Day Pack Qty",
      "Remarks"
    ];

    List columnCellName = [
      "A",
      "B",
      "C",
      "D",
      "E",
      "F",
      "G",
      "H",
      "I",
      "J",
      "K",
      "L",
      "X",
      "Y",
      "Z",
      "AA",
      "AB",
    ];

    List mergedColName = [
      "M",
      "N",
      "O",
      "P",
      "Q",
      "R",
      "S",
      "T",
      "U",
      "V",
      "W"
    ];

    for (int i = 0; i < header.length; i++) {
      if (i == 12 ||
          i == 13 ||
          i == 14 ||
          i == 15 ||
          i == 16 ||
          i == 17 ||
          i == 18 ||
          i == 19 ||
          i == 20 ||
          i == 21 ||
          i == 22) {
        // Merge the cells first
        sheet.merge(
            CellIndex.indexByString("M2"), CellIndex.indexByString("W2"));

// Set the value of the merged cell explicitly
        sheet.cell(CellIndex.indexByString("M2")).value =
            TextCellValue("Size Information");
        // sheet.cell(CellIndex.indexByString("M2")).cellStyle = headerStyle;
        for (int m = 0; m < mergedColName.length; m++) {
          sheet
              .cell(CellIndex.indexByString("${mergedColName[m]}2"))
              .cellStyle = headerStyle;
        }
      } else {
        sheet
            .cell(CellIndex.indexByColumnRow(columnIndex: i, rowIndex: 1))
            .value = TextCellValue(header[i]);
        sheet
            .cell(CellIndex.indexByColumnRow(columnIndex: i, rowIndex: 1))
            .cellStyle = headerStyle;
      }

      for (int i = 0; i < data.length; i++) {
        sheet.setRowHeight(6 * i + 2, 25);

        // Merge Column for Decortaion-----------------

        for (int k = 0; k < columnCellName.length; k++) {
          sheet.merge(
              CellIndex.indexByString("${columnCellName[k]}${6 * i + 3}"),
              CellIndex.indexByString("${columnCellName[k]}${6 * i + 7}"));
          for (int l = 3; l <= 7; l++) {
            sheet
                .cell(
                    CellIndex.indexByString("${columnCellName[k]}${6 * i + l}"))
                .cellStyle = CellStyle(
              leftBorder: Border(borderStyle: BorderStyle.Thin),
              rightBorder: Border(borderStyle: BorderStyle.Thin),
              bottomBorder: Border(borderStyle: BorderStyle.Thin),
              topBorder: Border(borderStyle: BorderStyle.Thin),
            );
          }
        }

        // --------------------------------- Buyer --------------------------------- //

        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["buyer"]);

        // ------------------------------- Style ---------------------------------------- //
        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["style"]);

        // ----------------------------- Buy --------------------------------------------- //
        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["buy"]);

        // ----------------------------- Sales Order No ---------------------------------- //
        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["salesOrderNo"]);

        // ----------------------------- Item ------------------------------------------- //

        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 4, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["item"]);

        // ----------------------------- FG(Mat No) -------------------------------- //
        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 5, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["fgMatNo"]);

        // ---------------------------- Our No (FR) -------------------------------- //
        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 6, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["ourNoFR"]);

        // -------------------------- Color Name --------------------------------- //

        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 7, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["colorName"]);

        // ------------------------- Country Name --------------------------------- //

        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 8, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["countryName"]);

        // -------------------------- Production Order No ------------------------ //

        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 9, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["productionOrderNo"]);

        // --------------------------- Order Qty --------------------------------- //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 10, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["purchaseOrderNo"].toString());

        // -------------------------- Issue Qty ----------------------------------- //

        sheet
                .cell(CellIndex.indexByColumnRow(
                    columnIndex: 11, rowIndex: 6 * i + 2))
                .value =
            TextCellValue(data[i]["orderQty"].toString() +
                "/" +
                data[i]["issueQty"].toString());

        // ----------------------- D. Gr Qty ------------------------------------ //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 23, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["dGrQty"].toString());

        // ------------------------- Total Gr Cum Qty ------------------------------------ //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 24, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["totalGrCumQty"].toString());

        // ------------------------ Balance Gr Qty --------------------------------- //

        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 25, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["balanceQty"].toString());

        // ------------------------ Last Day Pack Cum Qty ----------------------------- //

        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 26, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["lastDayPackCumQty"].toString());

        // ------------------- Remarks -----------------------------//

        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 27, rowIndex: 6 * i + 2))
            .value = TextCellValue("");

        // ===================== Size Title ==================== //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 12, rowIndex: 6 * i + 2))
            .value = TextCellValue("Size");

        // -----------------  Different Size Title ---------------- //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 13, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size1"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 14, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size2"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 15, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size3"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 16, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size4"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 17, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size5"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 18, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size6"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 19, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size7"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 20, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size8"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 21, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size9"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 22, rowIndex: 6 * i + 2))
            .value = TextCellValue("");

        // ---------------------------------------------
        // ====================Applying Styles to Column 0-27 and Main Rows =====================
        // ---------------------------------------------

        for (int j = 0; j <= 27; j++) {
          sheet
              .cell(CellIndex.indexByColumnRow(
                  columnIndex: j, rowIndex: 6 * i + 2))
              .cellStyle = commonStyle;
        }
        // ===================== Order Qty ==================== //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 12, rowIndex: 6 * i + 3))
            .value = TextCellValue("Order Qty");
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 13, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size1"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 14, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size2"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 15, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size3"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 16, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size4"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 17, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size5"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 18, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size6"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 19, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size7"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 20, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size8"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 21, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size9"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 22, rowIndex: 6 * i + 3))
            .value = TextCellValue("");
        //////////////////////////////////////////////////////////////////////////////////////

        for (int j = 12; j <= 22; j++) {
          sheet
              .cell(CellIndex.indexByColumnRow(
                  columnIndex: j, rowIndex: 6 * i + 3))
              .cellStyle = centerStyle;
        }
        // ===================== Issue Qty ==================== //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 12, rowIndex: 6 * i + 4))
            .value = TextCellValue("Issue Qty");
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 13, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size1"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 14, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size2"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 15, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size3"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 16, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size4"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 17, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size5"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 18, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size6"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 19, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size7"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 20, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size8"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 21, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size9"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 22, rowIndex: 6 * i + 4))
            .value = TextCellValue("");

        //////////////////////////////////////////////////////////////////////////////////////

        for (int j = 12; j <= 22; j++) {
          sheet
              .cell(CellIndex.indexByColumnRow(
                  columnIndex: j, rowIndex: 6 * i + 4))
              .cellStyle = centerStyle;
        }

        // ===================== Received Qty ==================== //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 12, rowIndex: 6 * i + 5))
            .value = TextCellValue("Received");
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 13, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size1"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 14, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size2"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 15, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size3"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 16, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size4"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 17, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size5"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 18, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size6"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 19, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size7"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 20, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size8"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 21, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size9"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 22, rowIndex: 6 * i + 5))
            .value = TextCellValue("");
        //////////////////////////////////////////////////////////////////////////////////////

        for (int j = 12; j <= 22; j++) {
          sheet
              .cell(CellIndex.indexByColumnRow(
                  columnIndex: j, rowIndex: 6 * i + 5))
              .cellStyle = centerStyle;
        }
        // ===================== Total Cum Qty ==================== //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 12, rowIndex: 6 * i + 6))
            .value = TextCellValue("Total Cum.");
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 13, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size1"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 14, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size2"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 15, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size3"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 16, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size4"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 17, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size5"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 18, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size6"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 19, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size7"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 20, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size8"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 21, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size9"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 22, rowIndex: 6 * i + 6))
            .value = TextCellValue("");
        //////////////////////////////////////////////////////////////////////////////////////

        for (int j = 12; j <= 22; j++) {
          sheet
              .cell(CellIndex.indexByColumnRow(
                  columnIndex: j, rowIndex: 6 * i + 6))
              .cellStyle = centerStyle;
        }

        // Adding a signature field at every 44th, 86th, 128th, etc., row index
      }
    }
    // ================== Merge Cells for Signing Part ======================= //
    // ----------------- Prepared By ------------------- //
    sheet.merge(
        CellIndex.indexByColumnRow(
            columnIndex: 0, rowIndex: data.length * 6 + 3),
        CellIndex.indexByColumnRow(
            columnIndex: 6, rowIndex: data.length * 6 + 3));
    sheet.merge(
        CellIndex.indexByColumnRow(
            columnIndex: 0, rowIndex: data.length * 6 + 4),
        CellIndex.indexByColumnRow(
            columnIndex: 6, rowIndex: data.length * 6 + 4));
    sheet.merge(
        CellIndex.indexByColumnRow(
            columnIndex: 0, rowIndex: data.length * 6 + 5),
        CellIndex.indexByColumnRow(
            columnIndex: 6, rowIndex: data.length * 6 + 5));
    sheet.merge(
        CellIndex.indexByColumnRow(
            columnIndex: 0, rowIndex: data.length * 6 + 6),
        CellIndex.indexByColumnRow(
            columnIndex: 6, rowIndex: data.length * 6 + 6));
    // Set Height---------------------------------------------------------------------------------
    sheet.setRowHeight(data.length * 6 + 3, 50);
    sheet.setRowHeight(data.length * 6 + 4, 35);
    sheet.setRowHeight(data.length * 6 + 5, 35);
    sheet.setRowHeight(data.length * 6 + 6, 35);

    // Set Total D GR Cum Qty by Date-------------------------------------------------------------
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 21, rowIndex: data.length * 6 + 2))
        .value = TextCellValue("Total");

    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 23, rowIndex: data.length * 6 + 2))
        .value = TextCellValue(totalDGrCumQty.toString());

    //
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 0, rowIndex: data.length * 6 + 3))
        .value = TextCellValue("Prepared By");
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 0, rowIndex: data.length * 6 + 3))
        .cellStyle = signStyle;
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 0, rowIndex: data.length * 6 + 4))
        .value = TextCellValue("Name: ");
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 0, rowIndex: data.length * 6 + 4))
        .cellStyle = signStyle2;
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 0, rowIndex: data.length * 6 + 5))
        .value = TextCellValue("ID: ");
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 0, rowIndex: data.length * 6 + 5))
        .cellStyle = signStyle2;
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 0, rowIndex: data.length * 6 + 6))
        .value = TextCellValue("Designation: ");
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 0, rowIndex: data.length * 6 + 6))
        .cellStyle = signStyle2;
    // ----------------- Quality ------------------- //

    sheet.merge(
        CellIndex.indexByColumnRow(
            columnIndex: 7, rowIndex: data.length * 6 + 3),
        CellIndex.indexByColumnRow(
            columnIndex: 13, rowIndex: data.length * 6 + 3));
    sheet.merge(
        CellIndex.indexByColumnRow(
            columnIndex: 7, rowIndex: data.length * 6 + 4),
        CellIndex.indexByColumnRow(
            columnIndex: 13, rowIndex: data.length * 6 + 4));
    sheet.merge(
        CellIndex.indexByColumnRow(
            columnIndex: 7, rowIndex: data.length * 6 + 5),
        CellIndex.indexByColumnRow(
            columnIndex: 13, rowIndex: data.length * 6 + 5));
    sheet.merge(
        CellIndex.indexByColumnRow(
            columnIndex: 7, rowIndex: data.length * 6 + 6),
        CellIndex.indexByColumnRow(
            columnIndex: 13, rowIndex: data.length * 6 + 6));

    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 7, rowIndex: data.length * 6 + 3))
        .value = TextCellValue("Floor Chief(Quality)");
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 7, rowIndex: data.length * 6 + 3))
        .cellStyle = signStyle;
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 7, rowIndex: data.length * 6 + 4))
        .value = TextCellValue("Name: ");
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 7, rowIndex: data.length * 6 + 4))
        .cellStyle = signStyle2;
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 7, rowIndex: data.length * 6 + 5))
        .value = TextCellValue("ID: ");
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 7, rowIndex: data.length * 6 + 5))
        .cellStyle = signStyle2;
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 7, rowIndex: data.length * 6 + 6))
        .value = TextCellValue("Designation: ");
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 7, rowIndex: data.length * 6 + 6))
        .cellStyle = signStyle2;
    // ----------------- Packing ------------------- //
    sheet.merge(
        CellIndex.indexByColumnRow(
            columnIndex: 14, rowIndex: data.length * 6 + 3),
        CellIndex.indexByColumnRow(
            columnIndex: 20, rowIndex: data.length * 6 + 3));
    sheet.merge(
        CellIndex.indexByColumnRow(
            columnIndex: 14, rowIndex: data.length * 6 + 4),
        CellIndex.indexByColumnRow(
            columnIndex: 20, rowIndex: data.length * 6 + 4));
    sheet.merge(
        CellIndex.indexByColumnRow(
            columnIndex: 14, rowIndex: data.length * 6 + 5),
        CellIndex.indexByColumnRow(
            columnIndex: 20, rowIndex: data.length * 6 + 5));
    sheet.merge(
        CellIndex.indexByColumnRow(
            columnIndex: 14, rowIndex: data.length * 6 + 6),
        CellIndex.indexByColumnRow(
            columnIndex: 20, rowIndex: data.length * 6 + 6));
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 14, rowIndex: data.length * 6 + 3))
        .value = TextCellValue("Floor Chief(Packing)");
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 14, rowIndex: data.length * 6 + 3))
        .cellStyle = signStyle;
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 14, rowIndex: data.length * 6 + 4))
        .value = TextCellValue("Name: ");
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 14, rowIndex: data.length * 6 + 4))
        .cellStyle = signStyle2;
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 14, rowIndex: data.length * 6 + 5))
        .value = TextCellValue("ID: ");
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 14, rowIndex: data.length * 6 + 5))
        .cellStyle = signStyle2;
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 14, rowIndex: data.length * 6 + 6))
        .value = TextCellValue("Designation: ");
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 14, rowIndex: data.length * 6 + 6))
        .cellStyle = signStyle2;

    // ----------------- Production ------------------- //
    sheet.merge(
        CellIndex.indexByColumnRow(
            columnIndex: 21, rowIndex: data.length * 6 + 3),
        CellIndex.indexByColumnRow(
            columnIndex: 27, rowIndex: data.length * 6 + 3));
    sheet.merge(
        CellIndex.indexByColumnRow(
            columnIndex: 21, rowIndex: data.length * 6 + 4),
        CellIndex.indexByColumnRow(
            columnIndex: 27, rowIndex: data.length * 6 + 4));
    sheet.merge(
        CellIndex.indexByColumnRow(
            columnIndex: 21, rowIndex: data.length * 6 + 5),
        CellIndex.indexByColumnRow(
            columnIndex: 27, rowIndex: data.length * 6 + 5));
    sheet.merge(
        CellIndex.indexByColumnRow(
            columnIndex: 21, rowIndex: data.length * 6 + 6),
        CellIndex.indexByColumnRow(
            columnIndex: 27, rowIndex: data.length * 6 + 6));
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 21, rowIndex: data.length * 6 + 3))
        .value = TextCellValue("Floor Production In-charge");
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 21, rowIndex: data.length * 6 + 3))
        .cellStyle = signStyle;
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 21, rowIndex: data.length * 6 + 4))
        .value = TextCellValue("Name: ");
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 21, rowIndex: data.length * 6 + 4))
        .cellStyle = signStyle2;
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 21, rowIndex: data.length * 6 + 5))
        .value = TextCellValue("ID: ");
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 21, rowIndex: data.length * 6 + 5))
        .cellStyle = signStyle2;
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 21, rowIndex: data.length * 6 + 6))
        .value = TextCellValue("Designation: ");
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 21, rowIndex: data.length * 6 + 6))
        .cellStyle = signStyle2;

    // Call function save() to download the file
    var fileBytes = excel.save(
        fileName: 'Daily Work Center_${factoryName}_${floorName}_${date}.xlsx');
  }

  // Excel For Deleted Production Order Report ..................................................
  void createExcelForDeletedPO(List data) {
    var excel = Excel.createExcel();
    var sheet = excel["Sheet1"];

    // --- Style Setup ---
    var headerStyle = CellStyle(
      textWrapping: TextWrapping.WrapText,
      bold: true,
      italic: true,
      fontSize: 12,
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
      rotation: 90,
      leftBorder: Border(borderStyle: BorderStyle.Thin),
      rightBorder: Border(borderStyle: BorderStyle.Thin),
      bottomBorder: Border(borderStyle: BorderStyle.Thin),
      topBorder: Border(borderStyle: BorderStyle.Thin),
    );

    var centerStyle = CellStyle(
      horizontalAlign: HorizontalAlign.Center,
      leftBorder: Border(borderStyle: BorderStyle.Thin),
      rightBorder: Border(borderStyle: BorderStyle.Thin),
      bottomBorder: Border(borderStyle: BorderStyle.Thin),
      topBorder: Border(borderStyle: BorderStyle.Thin),
    );

    var verCenterStyle = CellStyle(
      verticalAlign: VerticalAlign.Center,
    );
    var pageTitleStyle = CellStyle(
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
      bold: true,
      fontSize: 12,
    );
    var signStyle = CellStyle(
      horizontalAlign: HorizontalAlign.Left,
      verticalAlign: VerticalAlign.Bottom,
      bold: true,
      fontSize: 12,
    );

    var commonStyle = CellStyle(
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
      textWrapping: TextWrapping.WrapText,
      fontSize: 10,
      leftBorder: Border(borderStyle: BorderStyle.Thin),
      rightBorder: Border(borderStyle: BorderStyle.Thin),
      bottomBorder: Border(borderStyle: BorderStyle.Thin),
      topBorder: Border(borderStyle: BorderStyle.Thin),
    );

    var rotateStyle = CellStyle(
      rotation: 90,
      leftBorder: Border(borderStyle: BorderStyle.Thin),
      rightBorder: Border(borderStyle: BorderStyle.Thin),
      bottomBorder: Border(borderStyle: BorderStyle.Thin),
      topBorder: Border(borderStyle: BorderStyle.Thin),
    );

    sheet.setRowHeight(1, 45);
    sheet.setRowHeight(0, 30);

    sheet.setColumnWidth(0, 6);
    sheet.setColumnWidth(1, 7);
    sheet.setColumnWidth(2, 6);
    sheet.setColumnWidth(3, 11);
    sheet.setColumnWidth(4, 6);
    sheet.setColumnWidth(5, 7);
    sheet.setColumnWidth(6, 11);
    sheet.setColumnWidth(7, 7);
    sheet.setColumnWidth(8, 10);
    sheet.setColumnWidth(9, 11);
    sheet.setColumnWidth(10, 6);
    sheet.setColumnWidth(11, 6);
    sheet.setColumnWidth(12, 9);
    sheet.setColumnWidth(13, 6);
    sheet.setColumnWidth(14, 6);
    sheet.setColumnWidth(15, 6);
    sheet.setColumnWidth(16, 6);
    sheet.setColumnWidth(17, 6);
    sheet.setColumnWidth(18, 6);
    sheet.setColumnWidth(19, 6);
    sheet.setColumnWidth(20, 6);
    sheet.setColumnWidth(21, 7);
    sheet.setColumnWidth(22, 4);
    sheet.setColumnWidth(23, 5);
    sheet.setColumnWidth(24, 7);
    sheet.setColumnWidth(25, 7);
    sheet.setColumnWidth(26, 7);
    sheet.setColumnWidth(27, 7);
    sheet.setColumnWidth(28, 7);

    // Row 1: Header Titles

    // ----------------------- Title -------------------------- //
    sheet.merge(CellIndex.indexByColumnRow(columnIndex: 6, rowIndex: 0),
        CellIndex.indexByColumnRow(columnIndex: 24, rowIndex: 0));
    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 6, rowIndex: 0)).value =
        TextCellValue("Work Center Report");
    sheet
        .cell(CellIndex.indexByColumnRow(columnIndex: 6, rowIndex: 0))
        .cellStyle = pageTitleStyle;

    List header = [
      "Factory",
      "Floor",
      "Style",
      "Sales Order No",
      "Item",
      "FG(Mat No)",
      "Our No(F.R)",
      "Color Name",
      "Country Name",
      "Prod Order No",
      "Purchase No",
      "Order/Issue",
      "Size Information",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "D GR Qty",
      "Total GR Cum Qty",
      "Balance GR Qty",
      "Last Day Pack Qty",
      "Remarks"
    ];

    List columnCellName = [
      "A",
      "B",
      "C",
      "D",
      "E",
      "F",
      "G",
      "H",
      "I",
      "J",
      "K",
      "L",
      "X",
      "Y",
      "Z",
      "AA",
      "AB",
    ];

    List mergedColName = [
      "M",
      "N",
      "O",
      "P",
      "Q",
      "R",
      "S",
      "T",
      "U",
      "V",
      "W"
    ];

    for (int i = 0; i < header.length; i++) {
      if (i == 12 ||
          i == 13 ||
          i == 14 ||
          i == 15 ||
          i == 16 ||
          i == 17 ||
          i == 18 ||
          i == 19 ||
          i == 20 ||
          i == 21 ||
          i == 22) {
        // Merge the cells first
        sheet.merge(
            CellIndex.indexByString("M2"), CellIndex.indexByString("W2"));

// Set the value of the merged cell explicitly
        sheet.cell(CellIndex.indexByString("M2")).value =
            TextCellValue("Size Information");
        // sheet.cell(CellIndex.indexByString("M2")).cellStyle = headerStyle;
        for (int m = 0; m < mergedColName.length; m++) {
          sheet
              .cell(CellIndex.indexByString("${mergedColName[m]}2"))
              .cellStyle = headerStyle;
        }
      } else {
        sheet
            .cell(CellIndex.indexByColumnRow(columnIndex: i, rowIndex: 1))
            .value = TextCellValue(header[i]);
        sheet
            .cell(CellIndex.indexByColumnRow(columnIndex: i, rowIndex: 1))
            .cellStyle = headerStyle;
      }

      for (int i = 0; i < data.length; i++) {
        sheet.setRowHeight(6 * i + 2, 25);

        // Merge Column for Decortaion-----------------

        for (int k = 0; k < columnCellName.length; k++) {
          sheet.merge(
              CellIndex.indexByString("${columnCellName[k]}${6 * i + 3}"),
              CellIndex.indexByString("${columnCellName[k]}${6 * i + 7}"));
          for (int l = 3; l <= 7; l++) {
            sheet
                .cell(
                    CellIndex.indexByString("${columnCellName[k]}${6 * i + l}"))
                .cellStyle = CellStyle(
              leftBorder: Border(borderStyle: BorderStyle.Thin),
              rightBorder: Border(borderStyle: BorderStyle.Thin),
              bottomBorder: Border(borderStyle: BorderStyle.Thin),
              topBorder: Border(borderStyle: BorderStyle.Thin),
            );
          }
        }

        // --------------------------------- Buyer --------------------------------- //

        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["factoryName"]);

        // ------------------------------- Style ---------------------------------------- //
        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["floorName"]);

        // ----------------------------- Buy --------------------------------------------- //
        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["style"]);

        // ----------------------------- Sales Order No ---------------------------------- //
        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["salesOrderNo"]);

        // ----------------------------- Item ------------------------------------------- //

        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 4, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["item"]);

        // ----------------------------- FG(Mat No) -------------------------------- //
        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 5, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["fgMatNo"]);

        // ---------------------------- Our No (FR) -------------------------------- //
        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 6, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["ourNoFR"]);

        // -------------------------- Color Name --------------------------------- //

        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 7, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["colorName"]);

        // ------------------------- Country Name --------------------------------- //

        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 8, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["countryName"]);

        // -------------------------- Production Order No ------------------------ //

        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 9, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["productionOrderNo"]);

        // --------------------------- Order Qty --------------------------------- //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 10, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["purchaseOrderNo"].toString());

        // -------------------------- Issue Qty ----------------------------------- //

        sheet
                .cell(CellIndex.indexByColumnRow(
                    columnIndex: 11, rowIndex: 6 * i + 2))
                .value =
            TextCellValue(data[i]["orderQty"].toString() +
                "/" +
                data[i]["issueQty"].toString());

        // ----------------------- D. Gr Qty ------------------------------------ //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 23, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["dGrQty"].toString());

        // ------------------------- Total Gr Cum Qty ------------------------------------ //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 24, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["totalGrCumQty"].toString());

        // ------------------------ Balance Gr Qty --------------------------------- //

        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 25, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["balanceQty"].toString());

        // ------------------------ Last Day Pack Cum Qty ----------------------------- //

        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 26, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["lastDayPackCumQty"].toString());

        // ------------------- Remarks -----------------------------//

        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 27, rowIndex: 6 * i + 2))
            .value = TextCellValue("");

        // ===================== Size Title ==================== //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 12, rowIndex: 6 * i + 2))
            .value = TextCellValue("Size");

        // -----------------  Different Size Title ---------------- //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 13, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size1"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 14, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size2"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 15, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size3"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 16, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size4"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 17, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size5"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 18, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size6"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 19, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size7"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 20, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size8"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 21, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size9"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 22, rowIndex: 6 * i + 2))
            .value = TextCellValue("");

        // ---------------------------------------------
        // ====================Applying Styles to Column 0-27 and Main Rows =====================
        // ---------------------------------------------

        for (int j = 0; j <= 27; j++) {
          sheet
              .cell(CellIndex.indexByColumnRow(
                  columnIndex: j, rowIndex: 6 * i + 2))
              .cellStyle = commonStyle;
        }
        // ===================== Order Qty ==================== //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 12, rowIndex: 6 * i + 3))
            .value = TextCellValue("Order Qty");
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 13, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size1"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 14, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size2"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 15, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size3"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 16, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size4"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 17, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size5"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 18, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size6"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 19, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size7"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 20, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size8"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 21, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size9"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 22, rowIndex: 6 * i + 3))
            .value = TextCellValue("");
        //////////////////////////////////////////////////////////////////////////////////////

        for (int j = 12; j <= 22; j++) {
          sheet
              .cell(CellIndex.indexByColumnRow(
                  columnIndex: j, rowIndex: 6 * i + 3))
              .cellStyle = centerStyle;
        }
        // ===================== Issue Qty ==================== //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 12, rowIndex: 6 * i + 4))
            .value = TextCellValue("Issue Qty");
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 13, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size1"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 14, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size2"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 15, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size3"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 16, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size4"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 17, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size5"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 18, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size6"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 19, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size7"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 20, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size8"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 21, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size9"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 22, rowIndex: 6 * i + 4))
            .value = TextCellValue("");

        //////////////////////////////////////////////////////////////////////////////////////

        for (int j = 12; j <= 22; j++) {
          sheet
              .cell(CellIndex.indexByColumnRow(
                  columnIndex: j, rowIndex: 6 * i + 4))
              .cellStyle = centerStyle;
        }

        // ===================== Received Qty ==================== //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 12, rowIndex: 6 * i + 5))
            .value = TextCellValue("Received");
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 13, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size1"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 14, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size2"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 15, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size3"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 16, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size4"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 17, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size5"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 18, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size6"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 19, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size7"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 20, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size8"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 21, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size9"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 22, rowIndex: 6 * i + 5))
            .value = TextCellValue("");
        //////////////////////////////////////////////////////////////////////////////////////

        for (int j = 12; j <= 22; j++) {
          sheet
              .cell(CellIndex.indexByColumnRow(
                  columnIndex: j, rowIndex: 6 * i + 5))
              .cellStyle = centerStyle;
        }
        // ===================== Total Cum Qty ==================== //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 12, rowIndex: 6 * i + 6))
            .value = TextCellValue("Total Cum.");
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 13, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size1"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 14, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size2"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 15, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size3"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 16, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size4"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 17, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size5"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 18, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size6"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 19, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size7"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 20, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size8"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 21, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size9"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 22, rowIndex: 6 * i + 6))
            .value = TextCellValue("");
        //////////////////////////////////////////////////////////////////////////////////////

        for (int j = 12; j <= 22; j++) {
          sheet
              .cell(CellIndex.indexByColumnRow(
                  columnIndex: j, rowIndex: 6 * i + 6))
              .cellStyle = centerStyle;
        }

        // Adding a signature field at every 44th, 86th, 128th, etc., row index
      }
    }

    // Call function save() to download the file
    var fileBytes = excel.save(fileName: 'Work_Center_Report.xlsx');
  }

  // Excel Creation for Style Wise Collection ========================================
  // ---------------------------------------------------------------------------------
  void createExcelForStyleWisePO(List data, String styleName) {
    var excel = Excel.createExcel();
    var sheet = excel["Sheet1"];

    // --- Style Setup ---
    var headerStyle = CellStyle(
      textWrapping: TextWrapping.WrapText,
      bold: true,
      italic: true,
      fontSize: 12,
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
      rotation: 90,
      leftBorder: Border(borderStyle: BorderStyle.Thin),
      rightBorder: Border(borderStyle: BorderStyle.Thin),
      bottomBorder: Border(borderStyle: BorderStyle.Thin),
      topBorder: Border(borderStyle: BorderStyle.Thin),
    );

    var centerStyle = CellStyle(
      horizontalAlign: HorizontalAlign.Center,
      leftBorder: Border(borderStyle: BorderStyle.Thin),
      rightBorder: Border(borderStyle: BorderStyle.Thin),
      bottomBorder: Border(borderStyle: BorderStyle.Thin),
      topBorder: Border(borderStyle: BorderStyle.Thin),
    );

    var verCenterStyle = CellStyle(
      verticalAlign: VerticalAlign.Center,
    );
    var pageTitleStyle = CellStyle(
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
      bold: true,
      fontSize: 12,
    );
    var signStyle = CellStyle(
      horizontalAlign: HorizontalAlign.Left,
      verticalAlign: VerticalAlign.Bottom,
      bold: true,
      fontSize: 12,
    );

    var commonStyle = CellStyle(
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
      textWrapping: TextWrapping.WrapText,
      fontSize: 10,
      leftBorder: Border(borderStyle: BorderStyle.Thin),
      rightBorder: Border(borderStyle: BorderStyle.Thin),
      bottomBorder: Border(borderStyle: BorderStyle.Thin),
      topBorder: Border(borderStyle: BorderStyle.Thin),
    );

    var rotateStyle = CellStyle(
      rotation: 90,
      leftBorder: Border(borderStyle: BorderStyle.Thin),
      rightBorder: Border(borderStyle: BorderStyle.Thin),
      bottomBorder: Border(borderStyle: BorderStyle.Thin),
      topBorder: Border(borderStyle: BorderStyle.Thin),
    );

    sheet.setRowHeight(1, 45);
    sheet.setRowHeight(0, 30);

    sheet.setColumnWidth(0, 6);
    sheet.setColumnWidth(1, 7);
    sheet.setColumnWidth(2, 6);
    sheet.setColumnWidth(3, 11);
    sheet.setColumnWidth(4, 6);
    sheet.setColumnWidth(5, 7);
    sheet.setColumnWidth(6, 11);
    sheet.setColumnWidth(7, 7);
    sheet.setColumnWidth(8, 10);
    sheet.setColumnWidth(9, 11);
    sheet.setColumnWidth(10, 6);
    sheet.setColumnWidth(11, 6);
    sheet.setColumnWidth(12, 9);
    sheet.setColumnWidth(13, 6);
    sheet.setColumnWidth(14, 6);
    sheet.setColumnWidth(15, 6);
    sheet.setColumnWidth(16, 6);
    sheet.setColumnWidth(17, 6);
    sheet.setColumnWidth(18, 6);
    sheet.setColumnWidth(19, 6);
    sheet.setColumnWidth(20, 6);
    sheet.setColumnWidth(21, 7);
    sheet.setColumnWidth(22, 4);
    sheet.setColumnWidth(23, 5);
    sheet.setColumnWidth(24, 7);
    sheet.setColumnWidth(25, 7);
    sheet.setColumnWidth(26, 7);
    sheet.setColumnWidth(27, 7);
    sheet.setColumnWidth(28, 7);

    // Row 1: Header Titles
    sheet.merge(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0),
        CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: 0));
    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0)).value =
        TextCellValue("Style: ${styleName}");

    // ----------------------- Title -------------------------- //
    sheet.merge(CellIndex.indexByColumnRow(columnIndex: 6, rowIndex: 0),
        CellIndex.indexByColumnRow(columnIndex: 24, rowIndex: 0));
    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 6, rowIndex: 0)).value =
        TextCellValue("Work Center Report");
    sheet
        .cell(CellIndex.indexByColumnRow(columnIndex: 6, rowIndex: 0))
        .cellStyle = pageTitleStyle;

    List header = [
      "Factory",
      "Floor",
      "Buyer",
      "Sales Order No",
      "Item",
      "FG(Mat No)",
      "Our No(F.R)",
      "Color Name",
      "Country Name",
      "Prod Order No",
      "Purchase No",
      "Order/Issue",
      "Size Information",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "D GR Qty",
      "Total GR Cum Qty",
      "Balance GR Qty",
      "Last Day Pack Qty",
      "Remarks"
    ];

    List columnCellName = [
      "A",
      "B",
      "C",
      "D",
      "E",
      "F",
      "G",
      "H",
      "I",
      "J",
      "K",
      "L",
      "X",
      "Y",
      "Z",
      "AA",
      "AB",
    ];

    List mergedColName = [
      "M",
      "N",
      "O",
      "P",
      "Q",
      "R",
      "S",
      "T",
      "U",
      "V",
      "W"
    ];

    for (int i = 0; i < header.length; i++) {
      if (i == 12 ||
          i == 13 ||
          i == 14 ||
          i == 15 ||
          i == 16 ||
          i == 17 ||
          i == 18 ||
          i == 19 ||
          i == 20 ||
          i == 21 ||
          i == 22) {
        // Merge the cells first
        sheet.merge(
            CellIndex.indexByString("M2"), CellIndex.indexByString("W2"));

// Set the value of the merged cell explicitly
        sheet.cell(CellIndex.indexByString("M2")).value =
            TextCellValue("Size Information");
        // sheet.cell(CellIndex.indexByString("M2")).cellStyle = headerStyle;
        for (int m = 0; m < mergedColName.length; m++) {
          sheet
              .cell(CellIndex.indexByString("${mergedColName[m]}2"))
              .cellStyle = headerStyle;
        }
      } else {
        sheet
            .cell(CellIndex.indexByColumnRow(columnIndex: i, rowIndex: 1))
            .value = TextCellValue(header[i]);
        sheet
            .cell(CellIndex.indexByColumnRow(columnIndex: i, rowIndex: 1))
            .cellStyle = headerStyle;
      }

      for (int i = 0; i < data.length; i++) {
        sheet.setRowHeight(6 * i + 2, 25);

        // Merge Column for Decortaion-----------------

        for (int k = 0; k < columnCellName.length; k++) {
          sheet.merge(
              CellIndex.indexByString("${columnCellName[k]}${6 * i + 3}"),
              CellIndex.indexByString("${columnCellName[k]}${6 * i + 7}"));
          for (int l = 3; l <= 7; l++) {
            sheet
                .cell(
                    CellIndex.indexByString("${columnCellName[k]}${6 * i + l}"))
                .cellStyle = CellStyle(
              leftBorder: Border(borderStyle: BorderStyle.Thin),
              rightBorder: Border(borderStyle: BorderStyle.Thin),
              bottomBorder: Border(borderStyle: BorderStyle.Thin),
              topBorder: Border(borderStyle: BorderStyle.Thin),
            );
          }
        }

        // --------------------------------- Buyer --------------------------------- //

        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["factoryName"]);

        // ------------------------------- Style ---------------------------------------- //
        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["floorName"]);

        // ----------------------------- Buy --------------------------------------------- //
        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["buyer"]);

        // ----------------------------- Sales Order No ---------------------------------- //
        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["salesOrderNo"]);

        // ----------------------------- Item ------------------------------------------- //

        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 4, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["item"]);

        // ----------------------------- FG(Mat No) -------------------------------- //
        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 5, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["fgMatNo"]);

        // ---------------------------- Our No (FR) -------------------------------- //
        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 6, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["ourNoFR"]);

        // -------------------------- Color Name --------------------------------- //

        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 7, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["colorName"]);

        // ------------------------- Country Name --------------------------------- //

        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 8, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["countryName"]);

        // -------------------------- Production Order No ------------------------ //

        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 9, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["productionOrderNo"]);

        // --------------------------- Order Qty --------------------------------- //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 10, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["purchaseOrderNo"].toString());

        // -------------------------- Issue Qty ----------------------------------- //

        sheet
                .cell(CellIndex.indexByColumnRow(
                    columnIndex: 11, rowIndex: 6 * i + 2))
                .value =
            TextCellValue(data[i]["orderQty"].toString() +
                "/" +
                data[i]["issueQty"].toString());

        // ----------------------- D. Gr Qty ------------------------------------ //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 23, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["dGrQty"].toString());

        // ------------------------- Total Gr Cum Qty ------------------------------------ //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 24, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["totalGrCumQty"].toString());

        // ------------------------ Balance Gr Qty --------------------------------- //

        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 25, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["balanceQty"].toString());

        // ------------------------ Last Day Pack Cum Qty ----------------------------- //

        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 26, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["lastDayPackCumQty"].toString());

        // ------------------- Remarks -----------------------------//

        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 27, rowIndex: 6 * i + 2))
            .value = TextCellValue("");

        // ===================== Size Title ==================== //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 12, rowIndex: 6 * i + 2))
            .value = TextCellValue("Size");

        // -----------------  Different Size Title ---------------- //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 13, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size1"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 14, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size2"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 15, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size3"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 16, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size4"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 17, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size5"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 18, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size6"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 19, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size7"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 20, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size8"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 21, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size9"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 22, rowIndex: 6 * i + 2))
            .value = TextCellValue("");

        // ---------------------------------------------
        // ====================Applying Styles to Column 0-27 and Main Rows =====================
        // ---------------------------------------------

        for (int j = 0; j <= 27; j++) {
          sheet
              .cell(CellIndex.indexByColumnRow(
                  columnIndex: j, rowIndex: 6 * i + 2))
              .cellStyle = commonStyle;
        }
        // ===================== Order Qty ==================== //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 12, rowIndex: 6 * i + 3))
            .value = TextCellValue("Order Qty");
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 13, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size1"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 14, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size2"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 15, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size3"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 16, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size4"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 17, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size5"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 18, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size6"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 19, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size7"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 20, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size8"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 21, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size9"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 22, rowIndex: 6 * i + 3))
            .value = TextCellValue("");
        //////////////////////////////////////////////////////////////////////////////////////

        for (int j = 12; j <= 22; j++) {
          sheet
              .cell(CellIndex.indexByColumnRow(
                  columnIndex: j, rowIndex: 6 * i + 3))
              .cellStyle = centerStyle;
        }
        // ===================== Issue Qty ==================== //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 12, rowIndex: 6 * i + 4))
            .value = TextCellValue("Issue Qty");
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 13, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size1"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 14, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size2"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 15, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size3"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 16, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size4"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 17, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size5"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 18, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size6"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 19, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size7"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 20, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size8"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 21, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size9"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 22, rowIndex: 6 * i + 4))
            .value = TextCellValue("");

        //////////////////////////////////////////////////////////////////////////////////////

        for (int j = 12; j <= 22; j++) {
          sheet
              .cell(CellIndex.indexByColumnRow(
                  columnIndex: j, rowIndex: 6 * i + 4))
              .cellStyle = centerStyle;
        }

        // ===================== Received Qty ==================== //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 12, rowIndex: 6 * i + 5))
            .value = TextCellValue("Received");
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 13, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size1"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 14, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size2"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 15, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size3"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 16, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size4"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 17, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size5"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 18, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size6"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 19, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size7"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 20, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size8"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 21, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size9"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 22, rowIndex: 6 * i + 5))
            .value = TextCellValue("");
        //////////////////////////////////////////////////////////////////////////////////////

        for (int j = 12; j <= 22; j++) {
          sheet
              .cell(CellIndex.indexByColumnRow(
                  columnIndex: j, rowIndex: 6 * i + 5))
              .cellStyle = centerStyle;
        }
        // ===================== Total Cum Qty ==================== //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 12, rowIndex: 6 * i + 6))
            .value = TextCellValue("Total Cum.");
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 13, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size1"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 14, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size2"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 15, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size3"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 16, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size4"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 17, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size5"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 18, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size6"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 19, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size7"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 20, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size8"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 21, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size9"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 22, rowIndex: 6 * i + 6))
            .value = TextCellValue("");
        //////////////////////////////////////////////////////////////////////////////////////

        for (int j = 12; j <= 22; j++) {
          sheet
              .cell(CellIndex.indexByColumnRow(
                  columnIndex: j, rowIndex: 6 * i + 6))
              .cellStyle = centerStyle;
        }

        // Adding a signature field at every 44th, 86th, 128th, etc., row index
      }
    }

    // Call function save() to download the file
    var fileBytes =
        excel.save(fileName: 'Work_Center_Report_${styleName}.xlsx');
  }

  // Excel Creation for Production Order Wise Collection ========================================
  // ---------------------------------------------------------------------------------
  void createExcelForPOWise(List data, String poNumber) {
    var excel = Excel.createExcel();
    var sheet = excel["Sheet1"];

    // --- Style Setup ---
    var headerStyle = CellStyle(
      textWrapping: TextWrapping.WrapText,
      bold: true,
      italic: true,
      fontSize: 12,
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
      rotation: 90,
      leftBorder: Border(borderStyle: BorderStyle.Thin),
      rightBorder: Border(borderStyle: BorderStyle.Thin),
      bottomBorder: Border(borderStyle: BorderStyle.Thin),
      topBorder: Border(borderStyle: BorderStyle.Thin),
    );

    var centerStyle = CellStyle(
      horizontalAlign: HorizontalAlign.Center,
      leftBorder: Border(borderStyle: BorderStyle.Thin),
      rightBorder: Border(borderStyle: BorderStyle.Thin),
      bottomBorder: Border(borderStyle: BorderStyle.Thin),
      topBorder: Border(borderStyle: BorderStyle.Thin),
    );

    var verCenterStyle = CellStyle(
      verticalAlign: VerticalAlign.Center,
    );
    var pageTitleStyle = CellStyle(
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
      bold: true,
      fontSize: 12,
    );
    var signStyle = CellStyle(
      horizontalAlign: HorizontalAlign.Left,
      verticalAlign: VerticalAlign.Bottom,
      bold: true,
      fontSize: 12,
    );

    var commonStyle = CellStyle(
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
      textWrapping: TextWrapping.WrapText,
      fontSize: 10,
      leftBorder: Border(borderStyle: BorderStyle.Thin),
      rightBorder: Border(borderStyle: BorderStyle.Thin),
      bottomBorder: Border(borderStyle: BorderStyle.Thin),
      topBorder: Border(borderStyle: BorderStyle.Thin),
    );

    var rotateStyle = CellStyle(
      rotation: 90,
      leftBorder: Border(borderStyle: BorderStyle.Thin),
      rightBorder: Border(borderStyle: BorderStyle.Thin),
      bottomBorder: Border(borderStyle: BorderStyle.Thin),
      topBorder: Border(borderStyle: BorderStyle.Thin),
    );

    sheet.setRowHeight(1, 45);
    sheet.setRowHeight(0, 30);

    sheet.setColumnWidth(0, 6);
    sheet.setColumnWidth(1, 7);
    sheet.setColumnWidth(2, 6);
    sheet.setColumnWidth(3, 11);
    sheet.setColumnWidth(4, 6);
    sheet.setColumnWidth(5, 7);
    sheet.setColumnWidth(6, 11);
    sheet.setColumnWidth(7, 7);
    sheet.setColumnWidth(8, 10);
    sheet.setColumnWidth(9, 11);
    sheet.setColumnWidth(10, 6);
    sheet.setColumnWidth(11, 6);
    sheet.setColumnWidth(12, 9);
    sheet.setColumnWidth(13, 6);
    sheet.setColumnWidth(14, 6);
    sheet.setColumnWidth(15, 6);
    sheet.setColumnWidth(16, 6);
    sheet.setColumnWidth(17, 6);
    sheet.setColumnWidth(18, 6);
    sheet.setColumnWidth(19, 6);
    sheet.setColumnWidth(20, 6);
    sheet.setColumnWidth(21, 7);
    sheet.setColumnWidth(22, 4);
    sheet.setColumnWidth(23, 5);
    sheet.setColumnWidth(24, 7);
    sheet.setColumnWidth(25, 7);
    sheet.setColumnWidth(26, 7);
    sheet.setColumnWidth(27, 7);
    sheet.setColumnWidth(28, 7);

    // Row 1: Header Titles
    sheet.merge(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0),
        CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: 0));
    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0)).value =
        TextCellValue("Prod Order: ${poNumber}");

    // ----------------------- Title -------------------------- //
    sheet.merge(CellIndex.indexByColumnRow(columnIndex: 6, rowIndex: 0),
        CellIndex.indexByColumnRow(columnIndex: 24, rowIndex: 0));
    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 6, rowIndex: 0)).value =
        TextCellValue("Work Center Report Production Order Wise");
    sheet
        .cell(CellIndex.indexByColumnRow(columnIndex: 6, rowIndex: 0))
        .cellStyle = pageTitleStyle;

    List header = [
      "Factory",
      "Floor",
      "Buyer",
      "Sales Order No",
      "Item",
      "FG(Mat No)",
      "Our No(F.R)",
      "Color Name",
      "Country Name",
      "style",
      "Purchase No",
      "Order/Issue",
      "Size Information",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "D GR Qty",
      "Total GR Cum Qty",
      "Balance GR Qty",
      "Last Day Pack Qty",
      "Remarks"
    ];

    List columnCellName = [
      "A",
      "B",
      "C",
      "D",
      "E",
      "F",
      "G",
      "H",
      "I",
      "J",
      "K",
      "L",
      "X",
      "Y",
      "Z",
      "AA",
      "AB",
    ];

    List mergedColName = [
      "M",
      "N",
      "O",
      "P",
      "Q",
      "R",
      "S",
      "T",
      "U",
      "V",
      "W"
    ];

    for (int i = 0; i < header.length; i++) {
      if (i == 12 ||
          i == 13 ||
          i == 14 ||
          i == 15 ||
          i == 16 ||
          i == 17 ||
          i == 18 ||
          i == 19 ||
          i == 20 ||
          i == 21 ||
          i == 22) {
        // Merge the cells first
        sheet.merge(
            CellIndex.indexByString("M2"), CellIndex.indexByString("W2"));

// Set the value of the merged cell explicitly
        sheet.cell(CellIndex.indexByString("M2")).value =
            TextCellValue("Size Information");
        // sheet.cell(CellIndex.indexByString("M2")).cellStyle = headerStyle;
        for (int m = 0; m < mergedColName.length; m++) {
          sheet
              .cell(CellIndex.indexByString("${mergedColName[m]}2"))
              .cellStyle = headerStyle;
        }
      } else {
        sheet
            .cell(CellIndex.indexByColumnRow(columnIndex: i, rowIndex: 1))
            .value = TextCellValue(header[i]);
        sheet
            .cell(CellIndex.indexByColumnRow(columnIndex: i, rowIndex: 1))
            .cellStyle = headerStyle;
      }

      for (int i = 0; i < data.length; i++) {
        sheet.setRowHeight(6 * i + 2, 25);

        // Merge Column for Decortaion-----------------

        for (int k = 0; k < columnCellName.length; k++) {
          sheet.merge(
              CellIndex.indexByString("${columnCellName[k]}${6 * i + 3}"),
              CellIndex.indexByString("${columnCellName[k]}${6 * i + 7}"));
          for (int l = 3; l <= 7; l++) {
            sheet
                .cell(
                    CellIndex.indexByString("${columnCellName[k]}${6 * i + l}"))
                .cellStyle = CellStyle(
              leftBorder: Border(borderStyle: BorderStyle.Thin),
              rightBorder: Border(borderStyle: BorderStyle.Thin),
              bottomBorder: Border(borderStyle: BorderStyle.Thin),
              topBorder: Border(borderStyle: BorderStyle.Thin),
            );
          }
        }

        // --------------------------------- Buyer --------------------------------- //

        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["factoryName"]);

        // ------------------------------- Style ---------------------------------------- //
        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["floorName"]);

        // ----------------------------- Buy --------------------------------------------- //
        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["buyer"]);

        // ----------------------------- Sales Order No ---------------------------------- //
        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["salesOrderNo"]);

        // ----------------------------- Item ------------------------------------------- //

        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 4, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["item"]);

        // ----------------------------- FG(Mat No) -------------------------------- //
        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 5, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["fgMatNo"]);

        // ---------------------------- Our No (FR) -------------------------------- //
        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 6, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["ourNoFR"]);

        // -------------------------- Color Name --------------------------------- //

        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 7, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["colorName"]);

        // ------------------------- Country Name --------------------------------- //

        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 8, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["countryName"]);

        // -------------------------- Production Order No ------------------------ //

        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 9, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["style"]);

        // --------------------------- Order Qty --------------------------------- //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 10, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["purchaseOrderNo"].toString());

        // -------------------------- Issue Qty ----------------------------------- //

        sheet
                .cell(CellIndex.indexByColumnRow(
                    columnIndex: 11, rowIndex: 6 * i + 2))
                .value =
            TextCellValue(data[i]["orderQty"].toString() +
                "/" +
                data[i]["issueQty"].toString());

        // ----------------------- D. Gr Qty ------------------------------------ //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 23, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["dGrQty"].toString());

        // ------------------------- Total Gr Cum Qty ------------------------------------ //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 24, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["totalGrCumQty"].toString());

        // ------------------------ Balance Gr Qty --------------------------------- //

        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 25, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["balanceQty"].toString());

        // ------------------------ Last Day Pack Cum Qty ----------------------------- //

        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 26, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["lastDayPackCumQty"].toString());

        // ------------------- Remarks -----------------------------//

        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 27, rowIndex: 6 * i + 2))
            .value = TextCellValue("");

        // ===================== Size Title ==================== //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 12, rowIndex: 6 * i + 2))
            .value = TextCellValue("Size");

        // -----------------  Different Size Title ---------------- //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 13, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size1"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 14, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size2"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 15, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size3"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 16, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size4"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 17, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size5"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 18, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size6"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 19, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size7"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 20, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size8"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 21, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size9"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 22, rowIndex: 6 * i + 2))
            .value = TextCellValue("");

        // ---------------------------------------------
        // ====================Applying Styles to Column 0-27 and Main Rows =====================
        // ---------------------------------------------

        for (int j = 0; j <= 27; j++) {
          sheet
              .cell(CellIndex.indexByColumnRow(
                  columnIndex: j, rowIndex: 6 * i + 2))
              .cellStyle = commonStyle;
        }
        // ===================== Order Qty ==================== //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 12, rowIndex: 6 * i + 3))
            .value = TextCellValue("Order Qty");
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 13, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size1"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 14, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size2"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 15, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size3"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 16, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size4"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 17, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size5"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 18, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size6"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 19, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size7"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 20, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size8"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 21, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size9"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 22, rowIndex: 6 * i + 3))
            .value = TextCellValue("");
        //////////////////////////////////////////////////////////////////////////////////////

        for (int j = 12; j <= 22; j++) {
          sheet
              .cell(CellIndex.indexByColumnRow(
                  columnIndex: j, rowIndex: 6 * i + 3))
              .cellStyle = centerStyle;
        }
        // ===================== Issue Qty ==================== //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 12, rowIndex: 6 * i + 4))
            .value = TextCellValue("Issue Qty");
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 13, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size1"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 14, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size2"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 15, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size3"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 16, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size4"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 17, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size5"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 18, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size6"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 19, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size7"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 20, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size8"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 21, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size9"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 22, rowIndex: 6 * i + 4))
            .value = TextCellValue("");

        //////////////////////////////////////////////////////////////////////////////////////

        for (int j = 12; j <= 22; j++) {
          sheet
              .cell(CellIndex.indexByColumnRow(
                  columnIndex: j, rowIndex: 6 * i + 4))
              .cellStyle = centerStyle;
        }

        // ===================== Received Qty ==================== //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 12, rowIndex: 6 * i + 5))
            .value = TextCellValue("Received");
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 13, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size1"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 14, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size2"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 15, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size3"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 16, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size4"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 17, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size5"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 18, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size6"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 19, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size7"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 20, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size8"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 21, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size9"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 22, rowIndex: 6 * i + 5))
            .value = TextCellValue("");
        //////////////////////////////////////////////////////////////////////////////////////

        for (int j = 12; j <= 22; j++) {
          sheet
              .cell(CellIndex.indexByColumnRow(
                  columnIndex: j, rowIndex: 6 * i + 5))
              .cellStyle = centerStyle;
        }
        // ===================== Total Cum Qty ==================== //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 12, rowIndex: 6 * i + 6))
            .value = TextCellValue("Total Cum.");
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 13, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size1"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 14, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size2"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 15, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size3"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 16, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size4"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 17, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size5"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 18, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size6"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 19, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size7"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 20, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size8"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 21, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size9"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 22, rowIndex: 6 * i + 6))
            .value = TextCellValue("");
        //////////////////////////////////////////////////////////////////////////////////////

        for (int j = 12; j <= 22; j++) {
          sheet
              .cell(CellIndex.indexByColumnRow(
                  columnIndex: j, rowIndex: 6 * i + 6))
              .cellStyle = centerStyle;
        }

        // Adding a signature field at every 44th, 86th, 128th, etc., row index
      }
    }

    // Call function save() to download the file
    var fileBytes = excel.save(fileName: 'Work_Center_Report_${poNumber}.xlsx');
  }

  // Excel Creation for Sales Order Wise Collection ========================================
  // ---------------------------------------------------------------------------------
  void createExcelForSalesOrderWise(List data, String salesOrderNumber) {
    var excel = Excel.createExcel();
    var sheet = excel["Sheet1"];

    // --- Style Setup ---
    var headerStyle = CellStyle(
      textWrapping: TextWrapping.WrapText,
      bold: true,
      italic: true,
      fontSize: 12,
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
      rotation: 90,
      leftBorder: Border(borderStyle: BorderStyle.Thin),
      rightBorder: Border(borderStyle: BorderStyle.Thin),
      bottomBorder: Border(borderStyle: BorderStyle.Thin),
      topBorder: Border(borderStyle: BorderStyle.Thin),
    );

    var centerStyle = CellStyle(
      horizontalAlign: HorizontalAlign.Center,
      leftBorder: Border(borderStyle: BorderStyle.Thin),
      rightBorder: Border(borderStyle: BorderStyle.Thin),
      bottomBorder: Border(borderStyle: BorderStyle.Thin),
      topBorder: Border(borderStyle: BorderStyle.Thin),
    );

    var verCenterStyle = CellStyle(
      verticalAlign: VerticalAlign.Center,
    );
    var pageTitleStyle = CellStyle(
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
      bold: true,
      fontSize: 12,
    );
    var signStyle = CellStyle(
      horizontalAlign: HorizontalAlign.Left,
      verticalAlign: VerticalAlign.Bottom,
      bold: true,
      fontSize: 12,
    );

    var commonStyle = CellStyle(
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
      textWrapping: TextWrapping.WrapText,
      fontSize: 10,
      leftBorder: Border(borderStyle: BorderStyle.Thin),
      rightBorder: Border(borderStyle: BorderStyle.Thin),
      bottomBorder: Border(borderStyle: BorderStyle.Thin),
      topBorder: Border(borderStyle: BorderStyle.Thin),
    );

    var rotateStyle = CellStyle(
      rotation: 90,
      leftBorder: Border(borderStyle: BorderStyle.Thin),
      rightBorder: Border(borderStyle: BorderStyle.Thin),
      bottomBorder: Border(borderStyle: BorderStyle.Thin),
      topBorder: Border(borderStyle: BorderStyle.Thin),
    );

    sheet.setRowHeight(1, 45);
    sheet.setRowHeight(0, 30);

    sheet.setColumnWidth(0, 6);
    sheet.setColumnWidth(1, 7);
    sheet.setColumnWidth(2, 6);
    sheet.setColumnWidth(3, 11);
    sheet.setColumnWidth(4, 6);
    sheet.setColumnWidth(5, 7);
    sheet.setColumnWidth(6, 11);
    sheet.setColumnWidth(7, 7);
    sheet.setColumnWidth(8, 10);
    sheet.setColumnWidth(9, 11);
    sheet.setColumnWidth(10, 6);
    sheet.setColumnWidth(11, 6);
    sheet.setColumnWidth(12, 9);
    sheet.setColumnWidth(13, 6);
    sheet.setColumnWidth(14, 6);
    sheet.setColumnWidth(15, 6);
    sheet.setColumnWidth(16, 6);
    sheet.setColumnWidth(17, 6);
    sheet.setColumnWidth(18, 6);
    sheet.setColumnWidth(19, 6);
    sheet.setColumnWidth(20, 6);
    sheet.setColumnWidth(21, 7);
    sheet.setColumnWidth(22, 4);
    sheet.setColumnWidth(23, 5);
    sheet.setColumnWidth(24, 7);
    sheet.setColumnWidth(25, 7);
    sheet.setColumnWidth(26, 7);
    sheet.setColumnWidth(27, 7);
    sheet.setColumnWidth(28, 7);

    // Row 1: Header Titles
    sheet.merge(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0),
        CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: 0));
    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0)).value =
        TextCellValue("SO No: ${salesOrderNumber}");

    // ----------------------- Title -------------------------- //
    sheet.merge(CellIndex.indexByColumnRow(columnIndex: 6, rowIndex: 0),
        CellIndex.indexByColumnRow(columnIndex: 24, rowIndex: 0));
    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 6, rowIndex: 0)).value =
        TextCellValue("Work Center Report Sales Order Wise");
    sheet
        .cell(CellIndex.indexByColumnRow(columnIndex: 6, rowIndex: 0))
        .cellStyle = pageTitleStyle;

    List header = [
      "Factory",
      "Floor",
      "Buyer",
      "style",
      "Item",
      "FG(Mat No)",
      "Our No(F.R)",
      "Color Name",
      "Country Name",
      "Prod Order No",
      "Purchase No",
      "Order/Issue",
      "Size Information",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "D GR Qty",
      "Total GR Cum Qty",
      "Balance GR Qty",
      "Last Day Pack Qty",
      "Remarks"
    ];

    List columnCellName = [
      "A",
      "B",
      "C",
      "D",
      "E",
      "F",
      "G",
      "H",
      "I",
      "J",
      "K",
      "L",
      "X",
      "Y",
      "Z",
      "AA",
      "AB",
    ];

    List mergedColName = [
      "M",
      "N",
      "O",
      "P",
      "Q",
      "R",
      "S",
      "T",
      "U",
      "V",
      "W"
    ];

    for (int i = 0; i < header.length; i++) {
      if (i == 12 ||
          i == 13 ||
          i == 14 ||
          i == 15 ||
          i == 16 ||
          i == 17 ||
          i == 18 ||
          i == 19 ||
          i == 20 ||
          i == 21 ||
          i == 22) {
        // Merge the cells first
        sheet.merge(
            CellIndex.indexByString("M2"), CellIndex.indexByString("W2"));

// Set the value of the merged cell explicitly
        sheet.cell(CellIndex.indexByString("M2")).value =
            TextCellValue("Size Information");
        // sheet.cell(CellIndex.indexByString("M2")).cellStyle = headerStyle;
        for (int m = 0; m < mergedColName.length; m++) {
          sheet
              .cell(CellIndex.indexByString("${mergedColName[m]}2"))
              .cellStyle = headerStyle;
        }
      } else {
        sheet
            .cell(CellIndex.indexByColumnRow(columnIndex: i, rowIndex: 1))
            .value = TextCellValue(header[i]);
        sheet
            .cell(CellIndex.indexByColumnRow(columnIndex: i, rowIndex: 1))
            .cellStyle = headerStyle;
      }

      for (int i = 0; i < data.length; i++) {
        sheet.setRowHeight(6 * i + 2, 25);

        // Merge Column for Decortaion-----------------

        for (int k = 0; k < columnCellName.length; k++) {
          sheet.merge(
              CellIndex.indexByString("${columnCellName[k]}${6 * i + 3}"),
              CellIndex.indexByString("${columnCellName[k]}${6 * i + 7}"));
          for (int l = 3; l <= 7; l++) {
            sheet
                .cell(
                    CellIndex.indexByString("${columnCellName[k]}${6 * i + l}"))
                .cellStyle = CellStyle(
              leftBorder: Border(borderStyle: BorderStyle.Thin),
              rightBorder: Border(borderStyle: BorderStyle.Thin),
              bottomBorder: Border(borderStyle: BorderStyle.Thin),
              topBorder: Border(borderStyle: BorderStyle.Thin),
            );
          }
        }

        // --------------------------------- Buyer --------------------------------- //

        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["factoryName"]);

        // ------------------------------- Style ---------------------------------------- //
        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["floorName"]);

        // ----------------------------- Buy --------------------------------------------- //
        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["buyer"]);

        // ----------------------------- Sales Order No ---------------------------------- //
        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["style"]);

        // ----------------------------- Item ------------------------------------------- //

        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 4, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["item"]);

        // ----------------------------- FG(Mat No) -------------------------------- //
        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 5, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["fgMatNo"]);

        // ---------------------------- Our No (FR) -------------------------------- //
        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 6, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["ourNoFR"]);

        // -------------------------- Color Name --------------------------------- //

        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 7, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["colorName"]);

        // ------------------------- Country Name --------------------------------- //

        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 8, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["countryName"]);

        // -------------------------- Production Order No ------------------------ //

        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 9, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["productionOrderNo"]);

        // --------------------------- Order Qty --------------------------------- //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 10, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["purchaseOrderNo"].toString());

        // -------------------------- Issue Qty ----------------------------------- //

        sheet
                .cell(CellIndex.indexByColumnRow(
                    columnIndex: 11, rowIndex: 6 * i + 2))
                .value =
            TextCellValue(data[i]["orderQty"].toString() +
                "/" +
                data[i]["issueQty"].toString());

        // ----------------------- D. Gr Qty ------------------------------------ //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 23, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["dGrQty"].toString());

        // ------------------------- Total Gr Cum Qty ------------------------------------ //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 24, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["totalGrCumQty"].toString());

        // ------------------------ Balance Gr Qty --------------------------------- //

        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 25, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["balanceQty"].toString());

        // ------------------------ Last Day Pack Cum Qty ----------------------------- //

        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 26, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["lastDayPackCumQty"].toString());

        // ------------------- Remarks -----------------------------//

        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 27, rowIndex: 6 * i + 2))
            .value = TextCellValue("");

        // ===================== Size Title ==================== //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 12, rowIndex: 6 * i + 2))
            .value = TextCellValue("Size");

        // -----------------  Different Size Title ---------------- //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 13, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size1"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 14, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size2"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 15, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size3"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 16, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size4"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 17, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size5"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 18, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size6"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 19, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size7"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 20, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size8"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 21, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size9"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 22, rowIndex: 6 * i + 2))
            .value = TextCellValue("");

        // ---------------------------------------------
        // ====================Applying Styles to Column 0-27 and Main Rows =====================
        // ---------------------------------------------

        for (int j = 0; j <= 27; j++) {
          sheet
              .cell(CellIndex.indexByColumnRow(
                  columnIndex: j, rowIndex: 6 * i + 2))
              .cellStyle = commonStyle;
        }
        // ===================== Order Qty ==================== //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 12, rowIndex: 6 * i + 3))
            .value = TextCellValue("Order Qty");
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 13, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size1"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 14, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size2"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 15, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size3"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 16, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size4"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 17, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size5"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 18, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size6"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 19, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size7"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 20, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size8"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 21, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size9"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 22, rowIndex: 6 * i + 3))
            .value = TextCellValue("");
        //////////////////////////////////////////////////////////////////////////////////////

        for (int j = 12; j <= 22; j++) {
          sheet
              .cell(CellIndex.indexByColumnRow(
                  columnIndex: j, rowIndex: 6 * i + 3))
              .cellStyle = centerStyle;
        }
        // ===================== Issue Qty ==================== //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 12, rowIndex: 6 * i + 4))
            .value = TextCellValue("Issue Qty");
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 13, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size1"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 14, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size2"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 15, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size3"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 16, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size4"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 17, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size5"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 18, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size6"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 19, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size7"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 20, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size8"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 21, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size9"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 22, rowIndex: 6 * i + 4))
            .value = TextCellValue("");

        //////////////////////////////////////////////////////////////////////////////////////

        for (int j = 12; j <= 22; j++) {
          sheet
              .cell(CellIndex.indexByColumnRow(
                  columnIndex: j, rowIndex: 6 * i + 4))
              .cellStyle = centerStyle;
        }

        // ===================== Received Qty ==================== //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 12, rowIndex: 6 * i + 5))
            .value = TextCellValue("Received");
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 13, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size1"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 14, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size2"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 15, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size3"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 16, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size4"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 17, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size5"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 18, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size6"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 19, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size7"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 20, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size8"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 21, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size9"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 22, rowIndex: 6 * i + 5))
            .value = TextCellValue("");
        //////////////////////////////////////////////////////////////////////////////////////

        for (int j = 12; j <= 22; j++) {
          sheet
              .cell(CellIndex.indexByColumnRow(
                  columnIndex: j, rowIndex: 6 * i + 5))
              .cellStyle = centerStyle;
        }
        // ===================== Total Cum Qty ==================== //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 12, rowIndex: 6 * i + 6))
            .value = TextCellValue("Total Cum.");
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 13, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size1"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 14, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size2"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 15, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size3"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 16, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size4"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 17, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size5"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 18, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size6"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 19, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size7"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 20, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size8"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 21, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size9"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 22, rowIndex: 6 * i + 6))
            .value = TextCellValue("");
        //////////////////////////////////////////////////////////////////////////////////////

        for (int j = 12; j <= 22; j++) {
          sheet
              .cell(CellIndex.indexByColumnRow(
                  columnIndex: j, rowIndex: 6 * i + 6))
              .cellStyle = centerStyle;
        }

        // Adding a signature field at every 44th, 86th, 128th, etc., row index
      }
    }

    // Call function save() to download the file
    var fileBytes =
        excel.save(fileName: 'Work_Center_Report_${salesOrderNumber}.xlsx');
  }

  // Excel Creation for Date Wise Collection ========================================
  // ---------------------------------------------------------------------------------
  void createExcelForDateWise(List data, String date) {
    var excel = Excel.createExcel();
    var sheet = excel["Sheet1"];

    // --- Style Setup ---
    var headerStyle = CellStyle(
      textWrapping: TextWrapping.WrapText,
      bold: true,
      italic: true,
      fontSize: 12,
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
      rotation: 90,
      leftBorder: Border(borderStyle: BorderStyle.Thin),
      rightBorder: Border(borderStyle: BorderStyle.Thin),
      bottomBorder: Border(borderStyle: BorderStyle.Thin),
      topBorder: Border(borderStyle: BorderStyle.Thin),
    );

    var centerStyle = CellStyle(
      horizontalAlign: HorizontalAlign.Center,
      leftBorder: Border(borderStyle: BorderStyle.Thin),
      rightBorder: Border(borderStyle: BorderStyle.Thin),
      bottomBorder: Border(borderStyle: BorderStyle.Thin),
      topBorder: Border(borderStyle: BorderStyle.Thin),
    );

    var verCenterStyle = CellStyle(
      verticalAlign: VerticalAlign.Center,
    );
    var pageTitleStyle = CellStyle(
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
      bold: true,
      fontSize: 12,
    );
    var signStyle = CellStyle(
      horizontalAlign: HorizontalAlign.Left,
      verticalAlign: VerticalAlign.Bottom,
      bold: true,
      fontSize: 12,
    );

    var commonStyle = CellStyle(
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
      textWrapping: TextWrapping.WrapText,
      fontSize: 10,
      leftBorder: Border(borderStyle: BorderStyle.Thin),
      rightBorder: Border(borderStyle: BorderStyle.Thin),
      bottomBorder: Border(borderStyle: BorderStyle.Thin),
      topBorder: Border(borderStyle: BorderStyle.Thin),
    );

    var rotateStyle = CellStyle(
      rotation: 90,
      leftBorder: Border(borderStyle: BorderStyle.Thin),
      rightBorder: Border(borderStyle: BorderStyle.Thin),
      bottomBorder: Border(borderStyle: BorderStyle.Thin),
      topBorder: Border(borderStyle: BorderStyle.Thin),
    );

    sheet.setRowHeight(1, 45);
    sheet.setRowHeight(0, 30);

    sheet.setColumnWidth(0, 6);
    sheet.setColumnWidth(1, 7);
    sheet.setColumnWidth(2, 6);
    sheet.setColumnWidth(3, 11);
    sheet.setColumnWidth(4, 6);
    sheet.setColumnWidth(5, 7);
    sheet.setColumnWidth(6, 11);
    sheet.setColumnWidth(7, 7);
    sheet.setColumnWidth(8, 10);
    sheet.setColumnWidth(9, 11);
    sheet.setColumnWidth(10, 6);
    sheet.setColumnWidth(11, 6);
    sheet.setColumnWidth(12, 9);
    sheet.setColumnWidth(13, 6);
    sheet.setColumnWidth(14, 6);
    sheet.setColumnWidth(15, 6);
    sheet.setColumnWidth(16, 6);
    sheet.setColumnWidth(17, 6);
    sheet.setColumnWidth(18, 6);
    sheet.setColumnWidth(19, 6);
    sheet.setColumnWidth(20, 6);
    sheet.setColumnWidth(21, 7);
    sheet.setColumnWidth(22, 4);
    sheet.setColumnWidth(23, 5);
    sheet.setColumnWidth(24, 7);
    sheet.setColumnWidth(25, 7);
    sheet.setColumnWidth(26, 7);
    sheet.setColumnWidth(27, 7);
    sheet.setColumnWidth(28, 7);

    // Row 1: Header Titles
    sheet.merge(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0),
        CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: 0));
    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0)).value =
        TextCellValue("Date: ${date}");

    // ----------------------- Title -------------------------- //
    sheet.merge(CellIndex.indexByColumnRow(columnIndex: 6, rowIndex: 0),
        CellIndex.indexByColumnRow(columnIndex: 24, rowIndex: 0));
    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 6, rowIndex: 0)).value =
        TextCellValue("Work Center Report Date Wise");
    sheet
        .cell(CellIndex.indexByColumnRow(columnIndex: 6, rowIndex: 0))
        .cellStyle = pageTitleStyle;

    List header = [
      "Factory",
      "Floor",
      "Buyer",
      "style",
      "Sales Order",
      "FG(Mat No)",
      "Our No(F.R)",
      "Color Name",
      "Country Name",
      "Prod Order No",
      "Purchase No",
      "Order/Issue",
      "Size Information",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "D GR Qty",
      "Total GR Cum Qty",
      "Balance GR Qty",
      "Last Day Pack Qty",
      "Remarks"
    ];

    List columnCellName = [
      "A",
      "B",
      "C",
      "D",
      "E",
      "F",
      "G",
      "H",
      "I",
      "J",
      "K",
      "L",
      "X",
      "Y",
      "Z",
      "AA",
      "AB",
    ];

    List mergedColName = [
      "M",
      "N",
      "O",
      "P",
      "Q",
      "R",
      "S",
      "T",
      "U",
      "V",
      "W"
    ];

    for (int i = 0; i < header.length; i++) {
      if (i == 12 ||
          i == 13 ||
          i == 14 ||
          i == 15 ||
          i == 16 ||
          i == 17 ||
          i == 18 ||
          i == 19 ||
          i == 20 ||
          i == 21 ||
          i == 22) {
        // Merge the cells first
        sheet.merge(
            CellIndex.indexByString("M2"), CellIndex.indexByString("W2"));

// Set the value of the merged cell explicitly
        sheet.cell(CellIndex.indexByString("M2")).value =
            TextCellValue("Size Information");
        // sheet.cell(CellIndex.indexByString("M2")).cellStyle = headerStyle;
        for (int m = 0; m < mergedColName.length; m++) {
          sheet
              .cell(CellIndex.indexByString("${mergedColName[m]}2"))
              .cellStyle = headerStyle;
        }
      } else {
        sheet
            .cell(CellIndex.indexByColumnRow(columnIndex: i, rowIndex: 1))
            .value = TextCellValue(header[i]);
        sheet
            .cell(CellIndex.indexByColumnRow(columnIndex: i, rowIndex: 1))
            .cellStyle = headerStyle;
      }

      for (int i = 0; i < data.length; i++) {
        sheet.setRowHeight(6 * i + 2, 25);

        // Merge Column for Decortaion-----------------

        for (int k = 0; k < columnCellName.length; k++) {
          sheet.merge(
              CellIndex.indexByString("${columnCellName[k]}${6 * i + 3}"),
              CellIndex.indexByString("${columnCellName[k]}${6 * i + 7}"));
          for (int l = 3; l <= 7; l++) {
            sheet
                .cell(
                    CellIndex.indexByString("${columnCellName[k]}${6 * i + l}"))
                .cellStyle = CellStyle(
              leftBorder: Border(borderStyle: BorderStyle.Thin),
              rightBorder: Border(borderStyle: BorderStyle.Thin),
              bottomBorder: Border(borderStyle: BorderStyle.Thin),
              topBorder: Border(borderStyle: BorderStyle.Thin),
            );
          }
        }

        // --------------------------------- Factory --------------------------------- //

        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["factoryName"]);

        // ------------------------------- Floor ---------------------------------------- //
        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["floorName"]);

        // ----------------------------- Buyer --------------------------------------------- //
        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["buyer"]);

        // ----------------------------- Style ---------------------------------- //
        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["style"]);

        // ----------------------------- Sales order No ------------------------------------------- //

        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 4, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["salesOrderNo"]);

        // ----------------------------- FG(Mat No) -------------------------------- //
        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 5, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["fgMatNo"]);

        // ---------------------------- Our No (FR) -------------------------------- //
        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 6, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["ourNoFR"]);

        // -------------------------- Color Name --------------------------------- //

        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 7, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["colorName"]);

        // ------------------------- Country Name --------------------------------- //

        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 8, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["countryName"]);

        // -------------------------- Production Order No ------------------------ //

        sheet
            .cell(
                CellIndex.indexByColumnRow(columnIndex: 9, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["productionOrderNo"]);

        // --------------------------- Order Qty --------------------------------- //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 10, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["purchaseOrderNo"].toString());

        // -------------------------- Issue Qty ----------------------------------- //

        sheet
                .cell(CellIndex.indexByColumnRow(
                    columnIndex: 11, rowIndex: 6 * i + 2))
                .value =
            TextCellValue(data[i]["orderQty"].toString() +
                "/" +
                data[i]["issueQty"].toString());

        // ----------------------- D. Gr Qty ------------------------------------ //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 23, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["dGrQty"].toString());

        // ------------------------- Total Gr Cum Qty ------------------------------------ //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 24, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["totalGrCumQty"].toString());

        // ------------------------ Balance Gr Qty --------------------------------- //

        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 25, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["balanceQty"].toString());

        // ------------------------ Last Day Pack Cum Qty ----------------------------- //

        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 26, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["lastDayPackCumQty"].toString());

        // ------------------- Remarks -----------------------------//

        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 27, rowIndex: 6 * i + 2))
            .value = TextCellValue("");

        // ===================== Size Title ==================== //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 12, rowIndex: 6 * i + 2))
            .value = TextCellValue("Size");

        // -----------------  Different Size Title ---------------- //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 13, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size1"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 14, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size2"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 15, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size3"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 16, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size4"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 17, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size5"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 18, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size6"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 19, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size7"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 20, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size8"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 21, rowIndex: 6 * i + 2))
            .value = TextCellValue(data[i]["size9"]["sizeTitle"]);
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 22, rowIndex: 6 * i + 2))
            .value = TextCellValue("");

        // ---------------------------------------------
        // ====================Applying Styles to Column 0-27 and Main Rows =====================
        // ---------------------------------------------

        for (int j = 0; j <= 27; j++) {
          sheet
              .cell(CellIndex.indexByColumnRow(
                  columnIndex: j, rowIndex: 6 * i + 2))
              .cellStyle = commonStyle;
        }
        // ===================== Order Qty ==================== //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 12, rowIndex: 6 * i + 3))
            .value = TextCellValue("Order Qty");
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 13, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size1"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 14, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size2"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 15, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size3"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 16, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size4"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 17, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size5"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 18, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size6"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 19, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size7"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 20, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size8"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 21, rowIndex: 6 * i + 3))
            .value = TextCellValue(data[i]["size9"]["orderQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 22, rowIndex: 6 * i + 3))
            .value = TextCellValue("");
        //////////////////////////////////////////////////////////////////////////////////////

        for (int j = 12; j <= 22; j++) {
          sheet
              .cell(CellIndex.indexByColumnRow(
                  columnIndex: j, rowIndex: 6 * i + 3))
              .cellStyle = centerStyle;
        }
        // ===================== Issue Qty ==================== //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 12, rowIndex: 6 * i + 4))
            .value = TextCellValue("Issue Qty");
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 13, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size1"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 14, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size2"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 15, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size3"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 16, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size4"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 17, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size5"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 18, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size6"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 19, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size7"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 20, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size8"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 21, rowIndex: 6 * i + 4))
            .value = TextCellValue(data[i]["size9"]["issue"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 22, rowIndex: 6 * i + 4))
            .value = TextCellValue("");

        //////////////////////////////////////////////////////////////////////////////////////

        for (int j = 12; j <= 22; j++) {
          sheet
              .cell(CellIndex.indexByColumnRow(
                  columnIndex: j, rowIndex: 6 * i + 4))
              .cellStyle = centerStyle;
        }

        // ===================== Received Qty ==================== //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 12, rowIndex: 6 * i + 5))
            .value = TextCellValue("Received");
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 13, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size1"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 14, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size2"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 15, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size3"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 16, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size4"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 17, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size5"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 18, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size6"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 19, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size7"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 20, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size8"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 21, rowIndex: 6 * i + 5))
            .value = TextCellValue(data[i]["size9"]["receivedQty"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 22, rowIndex: 6 * i + 5))
            .value = TextCellValue("");
        //////////////////////////////////////////////////////////////////////////////////////

        for (int j = 12; j <= 22; j++) {
          sheet
              .cell(CellIndex.indexByColumnRow(
                  columnIndex: j, rowIndex: 6 * i + 5))
              .cellStyle = centerStyle;
        }
        // ===================== Total Cum Qty ==================== //
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 12, rowIndex: 6 * i + 6))
            .value = TextCellValue("Total Cum.");
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 13, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size1"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 14, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size2"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 15, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size3"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 16, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size4"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 17, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size5"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 18, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size6"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 19, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size7"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 20, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size8"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 21, rowIndex: 6 * i + 6))
            .value = TextCellValue(data[i]["size9"]["totalCum"].toString());
        sheet
            .cell(CellIndex.indexByColumnRow(
                columnIndex: 22, rowIndex: 6 * i + 6))
            .value = TextCellValue("");
        //////////////////////////////////////////////////////////////////////////////////////

        for (int j = 12; j <= 22; j++) {
          sheet
              .cell(CellIndex.indexByColumnRow(
                  columnIndex: j, rowIndex: 6 * i + 6))
              .cellStyle = centerStyle;
        }

        // Adding a signature field at every 44th, 86th, 128th, etc., row index
      }
    }

    // Call function save() to download the file
    var fileBytes = excel.save(fileName: 'Work_Center_Report_${date}.xlsx');
  }
}
