import 'package:excel/excel.dart';

// class ExcelGenerator {
//   void createExcel(List data) {
//     var excel = Excel.createExcel();

//     CellStyle cellStyle = CellStyle(
//       bold: true,
//       italic: true,
//       textWrapping: TextWrapping.WrapText,
//       fontFamily: getFontFamily(FontFamily.Comic_Sans_MS),
//       rotation: 0,
//       leftBorder: Border(),
//       rightBorder: Border(),
//       topBorder: Border(),
//       bottomBorder: Border(),
//     );

//     var sheet = excel['Daily Work Center Report'];

//     //  ------------- Buyer -------------
//     //  --------------------------------------------
//     var cell1 = sheet.cell(CellIndex.indexByString("A1"));
//     cell1.value = TextCellValue("Buyer");
//     cell1.cellStyle = cellStyle;
//     for (int i = 0; i < data.length; i++) {
//       var cell = sheet.cell(CellIndex.indexByString("A${5 * i + 2}"));
//       cell.value = TextCellValue(data[i]["Buyer"]);
//     }

//     //  ------------- Style -------------
//     //  --------------------------------------------
//     var cell2 = sheet.cell(CellIndex.indexByString("B1"));
//     cell2.value = TextCellValue("Style");
//     cell2.cellStyle = cellStyle;
//     for (int i = 0; i < data.length; i++) {
//       var cell = sheet.cell(CellIndex.indexByString("B${5 * i + 2}"));
//       cell.value = TextCellValue(data[i]["Style"]);
//     }

//     //  ------------- Buy -------------
//     //  --------------------------------------------
//     var cell3 = sheet.cell(CellIndex.indexByString("C1"));
//     cell3.value = TextCellValue("Buy");
//     cell3.cellStyle = cellStyle;
//     for (int i = 0; i < data.length; i++) {
//       var cell = sheet.cell(CellIndex.indexByString("C${5 * i + 2}"));
//       cell.value = TextCellValue(data[i]["Buy"]);
//     }

//     //  ------------- Sales Order No -------------
//     //  --------------------------------------------
//     var cell4 = sheet.cell(CellIndex.indexByString("D1"));
//     cell4.value = TextCellValue("Sales Order No");
//     cell4.cellStyle = cellStyle;
//     for (int i = 0; i < data.length; i++) {
//       var cell = sheet.cell(CellIndex.indexByString("D${5 * i + 2}"));
//       cell.value = TextCellValue(data[i]["Sales Order No"]);
//     }

//     //  ------------- Item -------------
//     //  --------------------------------------------
//     var cell5 = sheet.cell(CellIndex.indexByString("E1"));
//     cell5.value = TextCellValue("Item");
//     cell5.cellStyle = cellStyle;
//     for (int i = 0; i < data.length; i++) {
//       var cell = sheet.cell(CellIndex.indexByString("E${5 * i + 2}"));
//       cell.value = TextCellValue(data[i]["Item"]);
//     }

//     //  ------------- FG(Material) No -------------
//     //  --------------------------------------------
//     var cell6 = sheet.cell(CellIndex.indexByString("F1"));
//     cell6.value = TextCellValue("FG(Mat No)");
//     cell6.cellStyle = cellStyle;
//     for (int i = 0; i < data.length; i++) {
//       var cell = sheet.cell(CellIndex.indexByString("F${5 * i + 2}"));
//       cell.value = TextCellValue(data[i]["FG(Mat No)"]);
//     }

//     //  ------------- Our No(F.R) -------------
//     //  --------------------------------------------
//     var cell7 = sheet.cell(CellIndex.indexByString("G1"));
//     cell7.value = TextCellValue("Our No(F.R)");
//     cell7.cellStyle = cellStyle;
//     for (int i = 0; i < data.length; i++) {
//       var cell = sheet.cell(CellIndex.indexByString("G${5 * i + 2}"));
//       cell.value = TextCellValue(data[i]["Our No(FR)"]);
//     }

//     //  ------------- Color Name -------------
//     //  --------------------------------------------
//     var cell8 = sheet.cell(CellIndex.indexByString("H1"));
//     cell8.value = TextCellValue("Color Name");
//     cell8.cellStyle = cellStyle;
//     for (int i = 0; i < data.length; i++) {
//       var cell = sheet.cell(CellIndex.indexByString("H${5 * i + 2}"));
//       cell.value = TextCellValue(data[i]["Color Name"]);
//     }

//     //  ------------- Country Name -------------
//     //  --------------------------------------------
//     var cell9 = sheet.cell(CellIndex.indexByString("I1"));
//     cell9.value = TextCellValue("Country Name");
//     cell9.cellStyle = cellStyle;
//     for (int i = 0; i < data.length; i++) {
//       var cell = sheet.cell(CellIndex.indexByString("I${5 * i + 2}"));
//       cell.value = TextCellValue(data[i]["Country Name"]);
//     }

//     //  ------------- Production Order No -------------
//     //  --------------------------------------------
//     var cell10 = sheet.cell(CellIndex.indexByString("J1"));
//     cell10.value = TextCellValue("Production Order No");
//     cell10.cellStyle = cellStyle;
//     for (int i = 0; i < data.length; i++) {
//       var cell = sheet.cell(CellIndex.indexByString("J${5 * i + 2}"));
//       cell.value = TextCellValue(data[i]["Production Order No"]);
//     }

//     //  ------------- Order Quantity -------------
//     //  --------------------------------------------
//     var cell11 = sheet.cell(CellIndex.indexByString("K1"));
//     cell11.value = TextCellValue("Order Qty");
//     cell11.cellStyle = cellStyle;
//     for (int i = 0; i < data.length; i++) {
//       var cell = sheet.cell(CellIndex.indexByString("K${5 * i + 2}"));
//       cell.value = TextCellValue(data[i]["Order Qty"].toString());
//     }

//     //  ------------- Issue Quantity -------------
//     //  --------------------------------------------
//     var cell12 = sheet.cell(CellIndex.indexByString("L1"));
//     cell12.value = TextCellValue("Issue Qty");
//     cell12.cellStyle = cellStyle;
//     for (int i = 0; i < data.length; i++) {
//       var cell = sheet.cell(CellIndex.indexByString("L${5 * i + 2}"));
//       cell.value = TextCellValue(data[i]["Issue Qty"].toString());
//     }

//     //  ------------- Size Information -------------
//     //  --------------------------------------------
//     var cell13 = sheet.merge(
//         CellIndex.indexByString("M1"), CellIndex.indexByString("W1"),
//         customValue: TextCellValue("Size Information"));
//     for (int i = 0; i < data.length; i++) {
//       var cell13A = sheet.cell(CellIndex.indexByString("M${5 * i + 2}"));
//       cell13A.value = TextCellValue("Size");
//       var cell13B = sheet.cell(CellIndex.indexByString("M${5 * i + 3}"));
//       cell13B.value = TextCellValue("Order Qty");
//       var cell13C = sheet.cell(CellIndex.indexByString("M${5 * i + 4}"));
//       cell13C.value = TextCellValue("Issue");

//       var cell13D = sheet.cell(CellIndex.indexByString("M${5 * i + 5}"));
//       cell13D.value = TextCellValue("Received Qty");

//       var cell13E = sheet.cell(CellIndex.indexByString("M${5 * i + 6}"));
//       cell13E.value = TextCellValue("Total Cum");
//       // ====================== XXS Size and Data ========================= //
//       var cell13AA = sheet.cell(CellIndex.indexByString("N${5 * i + 2}"));
//       cell13AA.value = TextCellValue("XXS");
//       var xxsData = sheet.cell(CellIndex.indexByString("N${5 * i + 3}"));
//       xxsData.value = TextCellValue(data[i]["XXS"]["Order Qty"].toString());
//       var xxsIssueData = sheet.cell(CellIndex.indexByString("N${5 * i + 4}"));
//       xxsIssueData.value = TextCellValue(data[i]["XXS"]["Issue"].toString());
//       var xxsRcvData = sheet.cell(CellIndex.indexByString("N${5 * i + 5}"));
//       xxsRcvData.value =
//           TextCellValue(data[i]["XXS"]["Received Qty"].toString());
//       var xxsCumData = sheet.cell(CellIndex.indexByString("N${5 * i + 6}"));
//       xxsCumData.value = TextCellValue(data[i]["XXS"]["Total Cum"].toString());

//       // ====================== XS Size and Data ========================= //
//       var cell13AB = sheet.cell(CellIndex.indexByString("O${5 * i + 2}"));
//       cell13AB.value = TextCellValue("XS");
//       var xsData = sheet.cell(CellIndex.indexByString("O${5 * i + 3}"));
//       xsData.value = TextCellValue(data[i]["XS"]["Order Qty"].toString());
//       var xsIssueData = sheet.cell(CellIndex.indexByString("O${5 * i + 4}"));
//       xsIssueData.value = TextCellValue(data[i]["XS"]["Issue"].toString());
//       var xsRcvData = sheet.cell(CellIndex.indexByString("O${5 * i + 5}"));
//       xsRcvData.value = TextCellValue(data[i]["XS"]["Received Qty"].toString());
//       var xsCumData = sheet.cell(CellIndex.indexByString("O${5 * i + 6}"));
//       xsCumData.value = TextCellValue(data[i]["XS"]["Total Cum"].toString());

//       // ====================== S Size and Data ========================= //
//       var cell13AC = sheet.cell(CellIndex.indexByString("P${5 * i + 2}"));
//       cell13AC.value = TextCellValue("S");
//       var sData = sheet.cell(CellIndex.indexByString("P${5 * i + 3}"));
//       sData.value = TextCellValue(data[i]["S"]["Order Qty"].toString());
//       var sIssueData = sheet.cell(CellIndex.indexByString("P${5 * i + 4}"));
//       sIssueData.value = TextCellValue(data[i]["S"]["Issue"].toString());
//       var sRcvData = sheet.cell(CellIndex.indexByString("P${5 * i + 5}"));
//       sRcvData.value = TextCellValue(data[i]["S"]["Received Qty"].toString());
//       var sCumData = sheet.cell(CellIndex.indexByString("P${5 * i + 6}"));
//       sCumData.value = TextCellValue(data[i]["S"]["Total Cum"].toString());
//       // ====================== M Size and Data ========================= //
//       var cell13AD = sheet.cell(CellIndex.indexByString("Q${5 * i + 2}"));
//       cell13AD.value = TextCellValue("M");
//       var mData = sheet.cell(CellIndex.indexByString("Q${5 * i + 3}"));
//       mData.value = TextCellValue(data[i]["M"]["Order Qty"].toString());
//       var mIssueData = sheet.cell(CellIndex.indexByString("Q${5 * i + 4}"));
//       mIssueData.value = TextCellValue(data[i]["M"]["Issue"].toString());
//       var mRcvData = sheet.cell(CellIndex.indexByString("Q${5 * i + 5}"));
//       mRcvData.value = TextCellValue(data[i]["M"]["Received Qty"].toString());
//       var mCumData = sheet.cell(CellIndex.indexByString("Q${5 * i + 6}"));
//       mCumData.value = TextCellValue(data[i]["M"]["Total Cum"].toString());
//       // ====================== L Size and Data ========================= //
//       var cell13AE = sheet.cell(CellIndex.indexByString("R${5 * i + 2}"));
//       cell13AE.value = TextCellValue("L");
//       var lData = sheet.cell(CellIndex.indexByString("R${5 * i + 3}"));
//       lData.value = TextCellValue(data[i]["L"]["Order Qty"].toString());
//       var lIssueData = sheet.cell(CellIndex.indexByString("R${5 * i + 4}"));
//       lIssueData.value = TextCellValue(data[i]["L"]["Issue"].toString());
//       var lRcvData = sheet.cell(CellIndex.indexByString("R${5 * i + 5}"));
//       lRcvData.value = TextCellValue(data[i]["L"]["Received Qty"].toString());
//       var lCumData = sheet.cell(CellIndex.indexByString("R${5 * i + 6}"));
//       lCumData.value = TextCellValue(data[i]["L"]["Total Cum"].toString());
//       // ====================== XL Size and Data ========================= //
//       var cell13AF = sheet.cell(CellIndex.indexByString("S${5 * i + 2}"));
//       cell13AF.value = TextCellValue("XL");
//       var xlData = sheet.cell(CellIndex.indexByString("S${5 * i + 3}"));
//       xlData.value = TextCellValue(data[i]["XL"]["Order Qty"].toString());
//       var xlIssueData = sheet.cell(CellIndex.indexByString("S${5 * i + 4}"));
//       xlIssueData.value = TextCellValue(data[i]["XL"]["Issue"].toString());
//       var xlRcvData = sheet.cell(CellIndex.indexByString("S${5 * i + 5}"));
//       xlRcvData.value = TextCellValue(data[i]["XL"]["Received Qty"].toString());
//       var xlCumData = sheet.cell(CellIndex.indexByString("S${5 * i + 6}"));
//       xlCumData.value = TextCellValue(data[i]["XL"]["Total Cum"].toString());
//       // ====================== XXL Size and Data ========================= //
//       var cell13AG = sheet.cell(CellIndex.indexByString("T${5 * i + 2}"));
//       cell13AG.value = TextCellValue("XXL");
//       var xxlData = sheet.cell(CellIndex.indexByString("T${5 * i + 3}"));
//       xxlData.value = TextCellValue(data[i]["XXL"]["Order Qty"].toString());
//       var xxlIssueData = sheet.cell(CellIndex.indexByString("T${5 * i + 4}"));
//       xxlIssueData.value = TextCellValue(data[i]["XXL"]["Issue"].toString());
//       var xxlRcvData = sheet.cell(CellIndex.indexByString("T${5 * i + 5}"));
//       xxlRcvData.value =
//           TextCellValue(data[i]["XXL"]["Received Qty"].toString());
//       var xxlCumData = sheet.cell(CellIndex.indexByString("T${5 * i + 6}"));
//       xxlCumData.value = TextCellValue(data[i]["XXL"]["Total Cum"].toString());
//       // ====================== XXXL Size and Data ========================= //
//       var cell13AH = sheet.cell(CellIndex.indexByString("U${5 * i + 2}"));
//       cell13AH.value = TextCellValue("XXXL");
//       var xxxlData = sheet.cell(CellIndex.indexByString("U${5 * i + 3}"));
//       xxxlData.value = TextCellValue(data[i]["XXXL"]["Order Qty"].toString());
//       var xxxlIssueData = sheet.cell(CellIndex.indexByString("U${5 * i + 4}"));
//       xxxlIssueData.value = TextCellValue(data[i]["XXXL"]["Issue"].toString());
//       var xxxlRcvData = sheet.cell(CellIndex.indexByString("U${5 * i + 5}"));
//       xxxlRcvData.value =
//           TextCellValue(data[i]["XXXL"]["Received Qty"].toString());
//       var xxxlCumData = sheet.cell(CellIndex.indexByString("U${5 * i + 6}"));
//       xxxlCumData.value =
//           TextCellValue(data[i]["XXXL"]["Total Cum"].toString());
//       // ====================== Extra Size and Data ========================= //
//       var cell13AI = sheet.cell(CellIndex.indexByString("V${5 * i + 2}"));
//       cell13AI.value = TextCellValue("");
//     }

//     //  ------------- D. GR Quantity -------------
//     //  --------------------------------------------
//     var cell14 = sheet.cell(CellIndex.indexByString("X1"));
//     cell14.value = TextCellValue("D.Gr Qty");
//     cell14.cellStyle = cellStyle;
//     for (int i = 0; i < data.length; i++) {
//       var cell = sheet.cell(CellIndex.indexByString("X${5 * i + 2}"));
//       cell.value = TextCellValue(data[i]["D Gr Qty"].toString());
//     }

//     //  ------------- Total GR Cum Qty -------------
//     //  --------------------------------------------
//     var cell15 = sheet.cell(CellIndex.indexByString("Y1"));
//     cell15.value = TextCellValue("Total Gr. Cum Qty");
//     cell15.cellStyle = cellStyle;
//     for (int i = 0; i < data.length; i++) {
//       var cell = sheet.cell(CellIndex.indexByString("Y${5 * i + 2}"));
//       cell.value = TextCellValue(data[i]["Total Gr Cum Qty"].toString());
//     }

//     //  -------------Balance Gr. Qty -------------
//     //  --------------------------------------------
//     var cell16 = sheet.cell(CellIndex.indexByString("Z1"));
//     cell16.value = TextCellValue("Balance Gr. Qty");
//     cell16.cellStyle = cellStyle;
//     for (int i = 0; i < data.length; i++) {
//       var cell = sheet.cell(CellIndex.indexByString("Z${5 * i + 2}"));
//       cell.value = TextCellValue(data[i]["Balance Gr Qty"].toString());
//     }

//     //  ------------- Daily Pack Qty -------------
//     //  --------------------------------------------
//     var cell17 = sheet.cell(CellIndex.indexByString("AA1"));
//     cell17.value = TextCellValue("Daily Pack Qty");
//     cell17.cellStyle = cellStyle;
//     for (int i = 0; i < data.length; i++) {
//       var cell = sheet.cell(CellIndex.indexByString("AA${5 * i + 2}"));
//       cell.value = TextCellValue(data[i]["Daily Pack Qty"].toString());
//     }

//     //  ------------- Last Day Pack Cum Qty -------------
//     //  --------------------------------------------
//     var cell18 = sheet.cell(CellIndex.indexByString("AB1"));
//     cell18.value = TextCellValue("Last Day Pack Cum Qty");
//     cell18.cellStyle = cellStyle;
//     for (int i = 0; i < data.length; i++) {
//       var cell = sheet.cell(CellIndex.indexByString("AB${5 * i + 2}"));
//       cell.value = TextCellValue(data[i]["Last Day Pack Cum Qty"].toString());
//     }

//     //  ------------- Total Pack Cum Qty -------------
//     //  --------------------------------------------
//     var cell19 = sheet.cell(CellIndex.indexByString("AC1"));
//     cell19.value = TextCellValue("Total Gr Cum Qty");
//     cell19.cellStyle = cellStyle;
//     for (int i = 0; i < data.length; i++) {
//       var cell = sheet.cell(CellIndex.indexByString("AC${5 * i + 2}"));
//       cell.value = TextCellValue(data[i]["Total Gr Cum Qty"].toString());
//     }

//     //  ------------- Balance Qty -------------
//     //  --------------------------------------------
//     var cell20 = sheet.cell(CellIndex.indexByString("AD1"));
//     cell20.value = TextCellValue("Balance Qty");
//     cell20.cellStyle = cellStyle;
//     for (int i = 0; i < data.length; i++) {
//       var cell = sheet.cell(CellIndex.indexByString("AD${5 * i + 2}"));
//       cell.value = TextCellValue(data[i]["Balance Qty"].toString());
//     }

//     //  ------------- Remarks -------------
//     //  --------------------------------------------
//     var cell21 = sheet.cell(CellIndex.indexByString("AE1"));
//     cell21.value = TextCellValue("Remarks");
//     cell21.cellStyle = cellStyle;
//     for (int i = 0; i < data.length; i++) {
//       var cell = sheet.cell(CellIndex.indexByString("AE${5 * i + 2}"));
//       cell.value = TextCellValue(data[i]["Remarks"].toString());
//     }

//     // Call function save() to download the file
//     var fileBytes = excel.save(fileName: 'My_Excel_File_Name.xlsx');
//   }
// }

class ExcelGenerator {
  void createExcel(List data, factoryName, floorName, date) {
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
    var signStyle = CellStyle(
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
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
        .cellStyle = signStyle;

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
      "D Dr. Qty",
      "Total Gr Cum Qty",
      "Balance Gr Qty",
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
            columnIndex: 0, rowIndex: data.length * 6 + 2),
        CellIndex.indexByColumnRow(
            columnIndex: 6, rowIndex: data.length * 6 + 2));
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

    //
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 0, rowIndex: data.length * 6 + 2))
        .value = TextCellValue("Prepared By");
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 0, rowIndex: data.length * 6 + 2))
        .cellStyle = signStyle;
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 0, rowIndex: data.length * 6 + 3))
        .value = TextCellValue("Name: ");
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 0, rowIndex: data.length * 6 + 4))
        .value = TextCellValue("ID: ");
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 0, rowIndex: data.length * 6 + 5))
        .value = TextCellValue("Designation: ");
    // ----------------- Quality ------------------- //

    sheet.merge(
        CellIndex.indexByColumnRow(
            columnIndex: 7, rowIndex: data.length * 6 + 2),
        CellIndex.indexByColumnRow(
            columnIndex: 13, rowIndex: data.length * 6 + 2));
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

    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 7, rowIndex: data.length * 6 + 2))
        .value = TextCellValue("Floor Chief(Quality)");
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 7, rowIndex: data.length * 6 + 2))
        .cellStyle = signStyle;
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 7, rowIndex: data.length * 6 + 3))
        .value = TextCellValue("Name: ");
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 7, rowIndex: data.length * 6 + 4))
        .value = TextCellValue("ID: ");
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 7, rowIndex: data.length * 6 + 5))
        .value = TextCellValue("Designation: ");
    // ----------------- Packing ------------------- //
    sheet.merge(
        CellIndex.indexByColumnRow(
            columnIndex: 14, rowIndex: data.length * 6 + 2),
        CellIndex.indexByColumnRow(
            columnIndex: 20, rowIndex: data.length * 6 + 2));
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
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 14, rowIndex: data.length * 6 + 2))
        .value = TextCellValue("Floor Chief(Packing)");
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 14, rowIndex: data.length * 6 + 2))
        .cellStyle = signStyle;
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 14, rowIndex: data.length * 6 + 3))
        .value = TextCellValue("Name: ");
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 14, rowIndex: data.length * 6 + 4))
        .value = TextCellValue("ID: ");
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 14, rowIndex: data.length * 6 + 5))
        .value = TextCellValue("Designation: ");

    // ----------------- Production ------------------- //
    sheet.merge(
        CellIndex.indexByColumnRow(
            columnIndex: 21, rowIndex: data.length * 6 + 2),
        CellIndex.indexByColumnRow(
            columnIndex: 27, rowIndex: data.length * 6 + 2));
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
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 21, rowIndex: data.length * 6 + 2))
        .value = TextCellValue("Floor Production In-charge");
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 21, rowIndex: data.length * 6 + 2))
        .cellStyle = signStyle;
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 21, rowIndex: data.length * 6 + 3))
        .value = TextCellValue("Name: ");
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 21, rowIndex: data.length * 6 + 4))
        .value = TextCellValue("ID: ");
    sheet
        .cell(CellIndex.indexByColumnRow(
            columnIndex: 21, rowIndex: data.length * 6 + 5))
        .value = TextCellValue("Designation: ");

    // Call function save() to download the file
    var fileBytes = excel.save(fileName: 'My_Excel_File_Name.xlsx');
  }
}
