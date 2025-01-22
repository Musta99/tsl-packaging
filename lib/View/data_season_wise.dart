import 'package:dropdown_button2/dropdown_button2.dart';
import 'package:flutter/material.dart';
import 'package:google_fonts/google_fonts.dart';
import 'package:provider/provider.dart';
import 'package:tsl_packaging/State_Management/data_fetch_state_management.dart';
import 'package:tsl_packaging/State_Management/selection_state_management.dart';

import '../Services/excel_generator.dart';
import '../Utilities/Widgets/datatable_widget.dart';

class DataSeasonWise extends StatefulWidget {
  const DataSeasonWise({super.key});

  @override
  State<DataSeasonWise> createState() => _DataSeasonWiseState();
}

class _DataSeasonWiseState extends State<DataSeasonWise> {
  @override
  void initState() {
    // TODO: implement initState
    super.initState();
    Provider.of<DataFetchStateManagement>(context, listen: false)
        .fetchMonthName();
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      // floatingActionButton: FloatingActionButton.small(
      //   onPressed: () {
      //     ExcelGenerator().createExcelForDeletedPO(
      //         Provider.of<DataFetchStateManagement>(context, listen: false)
      //             .packagingDataSeasonWise);
      //   },
      //   child: Center(
      //     child: Image.asset(
      //       "assets/images/excel.png",
      //       height: 30,
      //       width: 30,
      //     ),
      //   ),
      // ),
      body: Container(
        margin: EdgeInsets.symmetric(
          horizontal: 20,
          vertical: 20,
        ),
        child: Column(
          children: [
            SelectableText("Daily Work Center Report"),
            Row(
              mainAxisAlignment: MainAxisAlignment.end,
              children: [
                SizedBox(
                  height: 45,
                  child: Card(
                    shape: RoundedRectangleBorder(
                        borderRadius: BorderRadius.circular(3)),
                    child: DropdownButtonHideUnderline(
                      child: DropdownButton2(
                          style: GoogleFonts.ptSerif(
                            fontSize: 12,
                          ),
                          dropdownStyleData: DropdownStyleData(
                            maxHeight: 150,
                          ),
                          hint: Text("Select Month"),
                          value: Provider.of<SelectionStateManagement>(context)
                              .selectedMonth,
                          onChanged: (value) async {
                            Provider.of<SelectionStateManagement>(context,
                                    listen: false)
                                .selectMonth(value.toString());
                            print(Provider.of<SelectionStateManagement>(context,
                                    listen: false)
                                .selectedMonth);

                            await Provider.of<DataFetchStateManagement>(context,
                                    listen: false)
                                .getDataAdminSeasonWise(
                                    context,
                                    Provider.of<SelectionStateManagement>(
                                            context,
                                            listen: false)
                                        .selectedMonth);
                          },
                          items: [
                            for (int i = 0;
                                i <
                                    Provider.of<DataFetchStateManagement>(
                                            context)
                                        .monthList
                                        .length;
                                i++)
                              DropdownMenuItem(
                                child: Text(
                                    Provider.of<DataFetchStateManagement>(
                                            context)
                                        .monthList[i]),
                                value: Provider.of<DataFetchStateManagement>(
                                        context,
                                        listen: false)
                                    .monthList[i],
                              )
                          ]),
                    ),
                  ),
                ),
              ],
            ),
            Expanded(
              child: Provider.of<DataFetchStateManagement>(context)
                      .packagingDataSeasonWise
                      .isEmpty
                  ? Center(
                      child: Text("No Data Found"),
                    )
                  : Provider.of<DataFetchStateManagement>(context)
                              .isLoadedSeasonWise ==
                          false
                      ? Center(
                          child: CircularProgressIndicator(),
                        )
                      : DatatableWidget(
                          data: Provider.of<DataFetchStateManagement>(context,
                                  listen: false)
                              .packagingDataSeasonWise),
            ),
          ],
        ),
      ),
    );
  }
}
