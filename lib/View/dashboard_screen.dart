import 'package:dropdown_button2/dropdown_button2.dart';
import 'package:flutter/material.dart';
import 'package:google_fonts/google_fonts.dart';
import 'package:pie_chart/pie_chart.dart';
import 'package:provider/provider.dart';
import 'package:tsl_packaging/State_Management/data_fetch_state_management.dart';
import 'package:tsl_packaging/State_Management/selection_state_management.dart';
import 'package:tsl_packaging/Utilities/styles.dart';

class DashboardScreen extends StatefulWidget {
  const DashboardScreen({super.key});

  @override
  State<DashboardScreen> createState() => _DashboardScreenState();
}

class _DashboardScreenState extends State<DashboardScreen> {
  @override
  void initState() {
    // TODO: implement initState
    super.initState();
    Provider.of<DataFetchStateManagement>(context, listen: false)
        .fetchUserData();

    Provider.of<DataFetchStateManagement>(context, listen: false)
        .fetchFactoryWisePOCount();
  }

  @override
  Widget build(BuildContext context) {
    return Container(
      child: Column(
        children: [
          Container(
            height: MediaQuery.of(context).size.height * 0.6,
            width: MediaQuery.of(context).size.width * 0.8,
            child: Row(
              children: [
                Container(
                  padding: EdgeInsets.all(12),
                  width: MediaQuery.of(context).size.width * 0.4,
                  height: MediaQuery.of(context).size.height * 0.6,
                  child: Column(
                    mainAxisAlignment: MainAxisAlignment.center,
                    children: [
                      SelectableText(
                        "Welecome to EasyPack - your packaging partner!",
                        style: GoogleFonts.ptSerif(
                          fontWeight: FontWeight.bold,
                        ),
                      ),
                      SizedBox(
                        height: MediaQuery.of(context).size.height * 0.04,
                      ),
                      SelectableText(
                        "EasyPack is a revolutionary software solution designed to optimize the garment packing process, ensuring seamless operations and real-time tracking. Built with efficiency in mind, it empowers teams to monitor progress, analyze performance, and manage buyer-specific packing data effortlessly. With its intuitive interface and robust features, EasyPack simplifies workflows, reduces manual errors, and accelerates decision-making for packing departments",
                      ),
                      SizedBox(
                        height: MediaQuery.of(context).size.height * 0.02,
                      ),
                      SelectableText(
                        "As a trusted tool developed by TSL R&D, EasyPack reflects our commitment to innovation and excellence in garment manufacturing. Whether you're keeping track of daily targets or evaluating long-term performance, EasyPack is your reliable partner for achieving operational success.",
                      )
                    ],
                  ),
                ),
                Container(
                  width: MediaQuery.of(context).size.width * 0.4,
                  height: MediaQuery.of(context).size.height * 0.6,
                  child: Column(
                    children: [
                      SizedBox(
                          height: MediaQuery.of(context).size.height * 0.3,
                          width: MediaQuery.of(context).size.width * 0.3,
                          child: PieChart(
                            dataMap: {
                              "TSl-1":
                                  Provider.of<DataFetchStateManagement>(context)
                                      .tsl1Count
                                      .length as dynamic,
                              "TSl-2":
                                  Provider.of<DataFetchStateManagement>(context)
                                      .tsl2Count
                                      .length as dynamic,
                              "TSl-3":
                                  Provider.of<DataFetchStateManagement>(context)
                                      .tsl3Count
                                      .length as dynamic,
                              "TSl-4":
                                  Provider.of<DataFetchStateManagement>(context)
                                      .tsl4Count
                                      .length as dynamic,
                            },
                            animationDuration: Duration(milliseconds: 800),
                          )),
                      Container(
                          height: MediaQuery.of(context).size.height * 0.3,
                          child: SingleChildScrollView(
                            child: DataTable(
                              dataTextStyle: GoogleFonts.ptSerif(),
                              border: TableBorder.all(),
                              headingRowHeight: 25,
                              headingRowColor:
                                  WidgetStatePropertyAll(Styles().palletes4),
                              headingTextStyle: TextStyle(
                                fontWeight: FontWeight.bold,
                                fontSize: 10,
                                color: Colors.white,
                              ),
                              columns: [
                                DataColumn(label: Text("Factory")),
                                DataColumn(label: Text("Floor")),
                                DataColumn(label: Text("Name")),
                                DataColumn(label: Text("Last Active")),
                              ],
                              rows: [
                                for (int i = 0;
                                    i <
                                        Provider.of<DataFetchStateManagement>(
                                                context,
                                                listen: false)
                                            .userList
                                            .length;
                                    i++)
                                  DataRow(cells: [
                                    DataCell(Text(
                                        Provider.of<DataFetchStateManagement>(
                                                context,
                                                listen: false)
                                            .userList[i]["factoryName"])),
                                    DataCell(Text(
                                        Provider.of<DataFetchStateManagement>(
                                                context,
                                                listen: false)
                                            .userList[i]["floorName"])),
                                    DataCell(Text(
                                        Provider.of<DataFetchStateManagement>(
                                                context,
                                                listen: false)
                                            .userList[i]["userName"])),
                                    DataCell(Text(
                                        Provider.of<DataFetchStateManagement>(
                                                context,
                                                listen: false)
                                            .userList[i]["lastActive"])),
                                  ])
                              ],
                            ),
                          ))
                    ],
                  ),
                )
              ],
            ),
          ),
          Container(
            padding: EdgeInsets.symmetric(
              horizontal: MediaQuery.of(context).size.width * 0.009,
            ),
            height: MediaQuery.of(context).size.height * 0.4,
            width: MediaQuery.of(context).size.width * 0.8,
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                SelectableText(
                  "Recently added PO",
                  style: GoogleFonts.roboto(
                    fontWeight: FontWeight.bold,
                  ),
                ),
                SizedBox(
                  height: MediaQuery.of(context).size.height * 0.02,
                ),
                Expanded(
                  child: Container(
                    width: MediaQuery.of(context).size.width,
                    child: SingleChildScrollView(
                      child: DataTable(
                        headingRowColor: WidgetStatePropertyAll(
                          Styles().palletes4,
                        ),
                        headingTextStyle: TextStyle(
                          fontWeight: FontWeight.bold,
                          color: Colors.white,
                        ),
                        headingRowHeight: 35,
                        border: TableBorder.all(
                          color: Colors.grey,
                        ),
                        columns: [
                          DataColumn(
                            label: DropdownButtonHideUnderline(
                              child: DropdownButton2(
                                  style: TextStyle(
                                    fontSize: 12,
                                    color: Colors.white,
                                    fontWeight: FontWeight.bold,
                                  ),
                                  iconStyleData: IconStyleData(
                                    iconSize: 16,
                                    iconEnabledColor: Colors.white,
                                  ),
                                  dropdownStyleData: DropdownStyleData(
                                      maxHeight: 120,
                                      decoration: BoxDecoration(
                                        color: Styles().palletes4,
                                      )),
                                  value: Provider.of<SelectionStateManagement>(
                                          context)
                                      .selectedFactory,
                                  onChanged: (value) async {
                                    Provider.of<SelectionStateManagement>(
                                            context,
                                            listen: false)
                                        .factorySelection(value);

                                    await Provider.of<DataFetchStateManagement>(
                                            context,
                                            listen: false)
                                        .fetchData(context);
                                  },
                                  items: [
                                    DropdownMenuItem(
                                      child: Text("TSL-1"),
                                      value: "TSL-1",
                                    ),
                                    DropdownMenuItem(
                                      child: Text("TSL-2"),
                                      value: "TSL-2",
                                    ),
                                    DropdownMenuItem(
                                      child: Text("TSL-3"),
                                      value: "TSL-3",
                                    ),
                                    DropdownMenuItem(
                                      child: Text("TSL-4"),
                                      value: "TSL-4",
                                    ),
                                  ]),
                            ),
                          ),
                          DataColumn(
                              label: Text(
                            "Floor",
                            softWrap: true,
                            textAlign: TextAlign.center,
                          )),
                          DataColumn(
                              label: Text(
                            "Date",
                            softWrap: true,
                            textAlign: TextAlign.center,
                          )),
                          DataColumn(label: Text("Style")),
                          DataColumn(label: Text("PO No")),
                          DataColumn(label: Text("Order Qty")),
                          DataColumn(label: Text("Issue Qty")),
                          DataColumn(label: Text("Balance Qty")),
                        ],
                        rows: [
                          for (int i = 0;
                              i <
                                  Provider.of<DataFetchStateManagement>(context)
                                      .recentlyAddedPO
                                      .length;
                              i++)
                            DataRow(
                              cells: [
                                DataCell(Center(
                                  child: Text(
                                      Provider.of<DataFetchStateManagement>(
                                              context)
                                          .recentlyAddedPO[i]["factoryName"]),
                                )),
                                DataCell(Center(
                                    child: Text(
                                        Provider.of<DataFetchStateManagement>(
                                                context)
                                            .recentlyAddedPO[i]["floorName"]))),
                                DataCell(Center(
                                    child: Text(
                                        Provider.of<DataFetchStateManagement>(
                                                context)
                                            .recentlyAddedPO[i]["date"]))),
                                DataCell(Center(
                                    child: Text(
                                        Provider.of<DataFetchStateManagement>(
                                                context)
                                            .recentlyAddedPO[i]["style"]))),
                                DataCell(Center(
                                    child: Text(
                                        Provider.of<DataFetchStateManagement>(
                                                    context)
                                                .recentlyAddedPO[i]
                                            ["productionOrderNo"]))),
                                DataCell(Center(
                                    child: Text(
                                        Provider.of<DataFetchStateManagement>(
                                                context)
                                            .recentlyAddedPO[i]["orderQty"]
                                            .toString()))),
                                DataCell(Center(
                                    child: Text(
                                        Provider.of<DataFetchStateManagement>(
                                                context)
                                            .recentlyAddedPO[i]["issueQty"]
                                            .toString()))),
                                DataCell(Center(
                                    child: Text(
                                        Provider.of<DataFetchStateManagement>(
                                                context)
                                            .recentlyAddedPO[i]["balanceQty"]
                                            .toString()))),
                              ],
                            ),
                        ],
                      ),
                    ),
                  ),
                ),
              ],
            ),
          ),
        ],
      ),
    );
  }
}




// ListView.builder(
//                             itemCount: Provider.of<DataFetchStateManagement>(
//                                     context,
//                                     listen: false)
//                                 .userList
//                                 .length,
//                             itemBuilder: (context, index) {
//                               return Column(
//                                 children: [
//                                   ListTile(
//                                     leading: Text(
//                                       Provider.of<DataFetchStateManagement>(
//                                               context,
//                                               listen: false)
//                                           .userList[index]["factoryName"],
//                                       style: GoogleFonts.ptSerif(
//                                         fontSize: 10,
//                                       ),
//                                     ),
//                                     title: Text(
//                                       Provider.of<DataFetchStateManagement>(
//                                               context,
//                                               listen: false)
//                                           .userList[index]["userName"],
//                                       style: GoogleFonts.ptSerif(fontSize: 13),
//                                     ),
//                                     subtitle: Text(
//                                       Provider.of<DataFetchStateManagement>(
//                                               context,
//                                               listen: false)
//                                           .userList[index]["floorName"],
//                                       style: GoogleFonts.ptSerif(
//                                         fontSize: 11,
//                                       ),
//                                     ),
//                                     trailing: Text(
//                                         Provider.of<DataFetchStateManagement>(
//                                                 context,
//                                                 listen: false)
//                                             .userList[index]["lastActive"]),
//                                   ),
//                                   Divider(),
//                                 ],
//                               );
//                             }),
