import 'package:flutter/foundation.dart';
import 'package:flutter/material.dart';
import 'package:flutter/widgets.dart';
import 'package:google_fonts/google_fonts.dart';
import 'package:provider/provider.dart';

import '../Services/excel_generator.dart';
import '../State_Management/data_fetch_state_management.dart';
import '../State_Management/selection_state_management.dart';
import '../Utilities/Widgets/datatable_widget.dart';
import '../Utilities/styles.dart';

class DataDateWiseScreen extends StatelessWidget {
  const DataDateWiseScreen({super.key});

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      floatingActionButton: FloatingActionButton.small(
        onPressed: () {
          ExcelGenerator().createExcelForDateWise(
            Provider.of<DataFetchStateManagement>(context, listen: false)
                .packagingDataDateWise,
            Provider.of<SelectionStateManagement>(context, listen: false)
                .selectedDateReport,
          );
        },
        child: Center(
          child: Image.asset(
            "assets/images/excel.png",
            height: 30,
            width: 30,
          ),
        ),
      ),
      body: Container(
        margin: EdgeInsets.symmetric(
          horizontal: 20,
          vertical: 20,
        ),
        child: Column(
          children: [
            // Name Title and Search Box According PO No: -------------------------------->
            SelectableText("Work Center Report Date Wise"),
            Container(
              child: Row(
                mainAxisAlignment: MainAxisAlignment.end,
                children: [
                  Container(
                    height: 45,
                    width: MediaQuery.of(context).size.width * 0.3,
                    child: Row(
                      mainAxisAlignment: MainAxisAlignment.end,
                      children: [
                        Container(
                          width: 190,
                          child: Card(
                            color: Styles().palletes1,
                            elevation: 2,
                            shape: RoundedRectangleBorder(
                                borderRadius: BorderRadius.circular(5)),
                            child: ListTile(
                              minTileHeight: 35,
                              title: Text(
                                Provider.of<SelectionStateManagement>(context)
                                    .selectedDateReport,
                                style: TextStyle(
                                  fontSize: 12,
                                ),
                              ),
                              trailing: IconButton(
                                onPressed: () async {
                                  DateTime? pickDate = await showDatePicker(
                                    context: context,
                                    firstDate: DateTime(2000),
                                    lastDate: DateTime(2050),
                                  );
                                  if (pickDate != null) {
                                    Provider.of<SelectionStateManagement>(
                                            context,
                                            listen: false)
                                        .dateSelectionReport(pickDate);
                                    await Provider.of<DataFetchStateManagement>(
                                            context,
                                            listen: false)
                                        .getDataAdminDateWise(
                                      context,
                                      Provider.of<SelectionStateManagement>(
                                              context,
                                              listen: false)
                                          .selectedDateReport,
                                    );
                                    print(pickDate.toString());
                                  }
                                },
                                icon: Icon(
                                  Icons.calendar_month,
                                  size: 18,
                                ),
                              ),
                            ),
                          ),
                        ),
                      ],
                    ),
                  ),
                ],
              ),
            ),

            Expanded(
              child: Provider.of<DataFetchStateManagement>(context)
                          .isLoadedDateWise ==
                      false
                  ? Center(
                      child: Text("Select a Date"),
                    )
                  : Provider.of<DataFetchStateManagement>(context)
                          .packagingDataDateWise
                          .isEmpty
                      ? Center(
                          child: Text("No Data Available"),
                        )
                      : DatatableWidget(
                          data: Provider.of<DataFetchStateManagement>(context,
                                  listen: false)
                              .packagingDataDateWise),
            ),
            // SizedBox(
            //   height: 40,
            //   child: Card(
            //     color: Styles().palletes1,
            //     elevation: 2,
            //     shape: RoundedRectangleBorder(
            //         borderRadius: BorderRadius.circular(5)),
            //     child: IconButton(
            //       hoverColor: Colors.transparent,
            //       onPressed: () {},
            //       icon: Icon(
            //         FontAwesomeIcons.file,
            //         size: 18,
            //       ),
            //     ),
            //   ),
            // ),
          ],
        ),
      ),
    );
  }
}
