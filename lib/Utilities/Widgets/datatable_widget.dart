import 'package:flutter/material.dart';
import 'package:provider/provider.dart';

import '../../State_Management/data_fetch_state_management.dart';
import '../styles.dart';

class DatatableWidget extends StatelessWidget {
  List data;
  DatatableWidget({super.key, required this.data});

  final ScrollController horizontal = ScrollController();
  final ScrollController vertical = ScrollController();

  @override
  Widget build(BuildContext context) {
    return Container(
      width: MediaQuery.of(context).size.width * 0.98,
      child: Padding(
        padding: const EdgeInsets.symmetric(
          horizontal: 10,
          vertical: 8,
        ),
        child: Scrollbar(
            controller: horizontal,
            thumbVisibility: true,
            trackVisibility: true,
            child: SingleChildScrollView(
              controller: horizontal,
              scrollDirection: Axis.horizontal,
              child: Scrollbar(
                controller: vertical,
                thumbVisibility: true,
                trackVisibility: true,
                child: SingleChildScrollView(
                  controller: vertical,
                  scrollDirection: Axis.vertical,
                  child: DataTable(
                    headingRowColor: WidgetStatePropertyAll(Styles().palletes4),
                    headingTextStyle: TextStyle(
                      fontWeight: FontWeight.bold,
                      color: Colors.white,
                    ),
                    dataRowHeight: 125,
                    headingRowHeight: 25,
                    border: TableBorder.all(
                      color: Colors.grey,
                    ),
                    columns: [
                      DataColumn(label: Text("Factory")),
                      DataColumn(label: Text("Floor")),
                      DataColumn(label: Text("Buyer")),
                      DataColumn(label: Text("Style")),
                      DataColumn(label: Text("Buy")),
                      DataColumn(label: Text("SO No")),
                      DataColumn(label: Text("Item")),
                      DataColumn(label: Text("FG (Mat No)")),
                      DataColumn(label: Text("Our No")),
                      DataColumn(label: Text("Color Name")),
                      DataColumn(label: Text("Country Name")),
                      DataColumn(label: Text("Production Order No")),
                      DataColumn(label: Text("Purchase Order")),
                      DataColumn(label: Text("Order/Issue Qty")),
                      DataColumn(
                          label: Text(
                            "Size Information",
                          ),
                          headingRowAlignment: MainAxisAlignment.center),
                      DataColumn(label: Text("D.Gr Qty")),
                      DataColumn(label: Text("Total GR Cum Qty")),
                      DataColumn(label: Text("Balance Qty")),
                      DataColumn(label: Text("Last Day Pack Qty")),
                      DataColumn(label: Text("Remarks")),
                      DataColumn(label: Text("Last Edited")),
                      DataColumn(label: Text("Edited By")),
                    ],
                    rows: [
                      for (int i = 0; i < data.length; i++)
                        DataRow(cells: [
                          DataCell(Text(data[i]["factoryName"])),
                          DataCell(Text(data[i]["floorName"])),
                          DataCell(Text(data[i]["buyer"])),
                          DataCell(Text(data[i]["style"])),
                          DataCell(Text(data[i]["buy"])),
                          DataCell(Text(data[i]["salesOrderNo"])),
                          DataCell(Text(data[i]["item"])),
                          DataCell(Text(data[i]["fgMatNo"])),
                          DataCell(Text(data[i]["ourNoFR"])),
                          DataCell(Text(data[i]["colorName"])),
                          DataCell(Text(data[i]["countryName"])),
                          DataCell(Text(data[i]["productionOrderNo"])),
                          DataCell(Text(data[i]["purchaseOrderNo"].toString())),
                          DataCell(Text(data[i]["orderQty"].toString() +
                              "/" +
                              data[i]["issueQty"].toString())),
                          DataCell(DataTable(
                              dataRowHeight: 25,
                              headingRowHeight: 25,
                              border: TableBorder.all(),
                              columns: [
                                DataColumn(label: Text("Size")),
                                DataColumn(
                                    label: Text(data[i]["size1"]["sizeTitle"])),
                                DataColumn(
                                    label: Text(data[i]["size2"]["sizeTitle"])),
                                DataColumn(
                                    label: Text(data[i]["size3"]["sizeTitle"])),
                                DataColumn(
                                    label: Text(data[i]["size4"]["sizeTitle"])),
                                DataColumn(
                                    label: Text(data[i]["size5"]["sizeTitle"])),
                                DataColumn(
                                    label: Text(data[i]["size6"]["sizeTitle"])),
                                DataColumn(
                                    label: Text(data[i]["size7"]["sizeTitle"])),
                                DataColumn(
                                    label: Text(data[i]["size8"]["sizeTitle"])),
                                DataColumn(
                                    label: Text(data[i]["size9"]["sizeTitle"])),
                                DataColumn(label: Text("")),
                              ],
                              rows: [
                                DataRow(
                                  cells: [
                                    DataCell(Text("O Qty")),
                                    DataCell(Text(data[i]["size1"]["orderQty"]
                                        .toString())),
                                    DataCell(Text(data[i]["size2"]["orderQty"]
                                        .toString())),
                                    DataCell(Text(data[i]["size3"]["orderQty"]
                                        .toString())),
                                    DataCell(Text(data[i]["size4"]["orderQty"]
                                        .toString())),
                                    DataCell(Text(data[i]["size5"]["orderQty"]
                                        .toString())),
                                    DataCell(Text(data[i]["size6"]["orderQty"]
                                        .toString())),
                                    DataCell(Text(data[i]["size7"]["orderQty"]
                                        .toString())),
                                    DataCell(Text(data[i]["size8"]["orderQty"]
                                        .toString())),
                                    DataCell(Text(data[i]["size9"]["orderQty"]
                                        .toString())),
                                    DataCell(Text("")),
                                  ],
                                ),
                                DataRow(
                                  cells: [
                                    DataCell(Text("Issue Qty")),
                                    DataCell(Text(
                                        data[i]["size1"]["issue"].toString())),
                                    DataCell(Text(
                                        data[i]["size2"]["issue"].toString())),
                                    DataCell(Text(
                                        data[i]["size3"]["issue"].toString())),
                                    DataCell(Text(
                                        data[i]["size4"]["issue"].toString())),
                                    DataCell(Text(
                                        data[i]["size5"]["issue"].toString())),
                                    DataCell(Text(
                                        data[i]["size6"]["issue"].toString())),
                                    DataCell(Text(
                                        data[i]["size7"]["issue"].toString())),
                                    DataCell(Text(
                                        data[i]["size8"]["issue"].toString())),
                                    DataCell(Text(
                                        data[i]["size9"]["issue"].toString())),
                                    DataCell(Text("")),
                                  ],
                                ),
                                DataRow(
                                  cells: [
                                    DataCell(Text("Received")),
                                    DataCell(Text(data[i]["size1"]
                                            ["receivedQty"]
                                        .toString())),
                                    DataCell(Text(data[i]["size2"]
                                            ["receivedQty"]
                                        .toString())),
                                    DataCell(Text(data[i]["size3"]
                                            ["receivedQty"]
                                        .toString())),
                                    DataCell(Text(data[i]["size4"]
                                            ["receivedQty"]
                                        .toString())),
                                    DataCell(Text(data[i]["size5"]
                                            ["receivedQty"]
                                        .toString())),
                                    DataCell(Text(data[i]["size6"]
                                            ["receivedQty"]
                                        .toString())),
                                    DataCell(Text(data[i]["size7"]
                                            ["receivedQty"]
                                        .toString())),
                                    DataCell(Text(data[i]["size8"]
                                            ["receivedQty"]
                                        .toString())),
                                    DataCell(Text(data[i]["size9"]
                                            ["receivedQty"]
                                        .toString())),
                                    DataCell(Text("")),
                                  ],
                                ),
                                DataRow(
                                  cells: [
                                    DataCell(Text("Total Cum")),
                                    DataCell(Text(data[i]["size1"]["totalCum"]
                                        .toString())),
                                    DataCell(Text(data[i]["size2"]["totalCum"]
                                        .toString())),
                                    DataCell(Text(data[i]["size3"]["totalCum"]
                                        .toString())),
                                    DataCell(Text(data[i]["size4"]["totalCum"]
                                        .toString())),
                                    DataCell(Text(data[i]["size5"]["totalCum"]
                                        .toString())),
                                    DataCell(Text(data[i]["size6"]["totalCum"]
                                        .toString())),
                                    DataCell(Text(data[i]["size7"]["totalCum"]
                                        .toString())),
                                    DataCell(Text(data[i]["size8"]["totalCum"]
                                        .toString())),
                                    DataCell(Text(data[i]["size9"]["totalCum"]
                                        .toString())),
                                    DataCell(Text("")),
                                  ],
                                ),
                              ])),
                          DataCell(Text(data[i]["dGrQty"].toString())),
                          DataCell(Text(data[i]["totalGrCumQty"].toString())),
                          DataCell(Text(data[i]["balanceQty"].toString())),
                          DataCell(
                              Text(data[i]["lastDayPackCumQty"].toString())),
                          DataCell(Text(data[i]["remarks"])),
                          DataCell(
                              Text(data[i]["date"] + "/" + data[i]["time"])),
                          DataCell(Text(data[i]["userName"])),
                        ]),
                    ],
                  ),
                ),
              ),
            )),
      ),
    );
  }
}
