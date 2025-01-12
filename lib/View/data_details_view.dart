import 'package:cloud_firestore/cloud_firestore.dart';
import 'package:dropdown_button2/dropdown_button2.dart';
import 'package:flutter/foundation.dart';
import 'package:flutter/material.dart';
import 'package:flutter/widgets.dart';
import 'package:font_awesome_flutter/font_awesome_flutter.dart';
import 'package:google_fonts/google_fonts.dart';
import 'package:provider/provider.dart';
import 'package:tsl_packaging/Services/excel_generator.dart';
import 'package:tsl_packaging/State_Management/data_fetch_state_management.dart';
import 'package:tsl_packaging/State_Management/selection_state_management.dart';
import 'package:tsl_packaging/Utilities/Widgets/datatable_widget.dart';
import 'package:tsl_packaging/Utilities/styles.dart';

class DataDetailsView extends StatelessWidget {
  const DataDetailsView({super.key});

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      backgroundColor: Styles().backgroundColor,
      floatingActionButton: FloatingActionButton.small(
        onPressed: () {
          ExcelGenerator().createExcel(
              Provider.of<DataFetchStateManagement>(context, listen: false)
                  .packagingData,
              Provider.of<SelectionStateManagement>(context, listen: false)
                  .selectedFactory,
              Provider.of<SelectionStateManagement>(context, listen: false)
                  .selectedFloor,
              Provider.of<SelectionStateManagement>(context, listen: false)
                  .selectedDate,
              Provider.of<DataFetchStateManagement>(context, listen: false)
                  .countDGrQty);
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
            SelectableText("Daily Work Center Report"),
            Container(
              child: Row(
                mainAxisAlignment: MainAxisAlignment.spaceBetween,
                children: [
                  // Factory Name and Floor Name -------------------------------->
                  Column(
                    children: [
                      Row(
                        children: [
                          SizedBox(
                            height: 40,
                            child: Card(
                              color: Styles().palletes1,
                              elevation: 2,
                              shape: RoundedRectangleBorder(
                                  borderRadius: BorderRadius.circular(5)),
                              child: DropdownButtonHideUnderline(
                                child: DropdownButton2(
                                    style: TextStyle(
                                      fontSize: 12,
                                    ),
                                    iconStyleData: IconStyleData(
                                      iconSize: 16,
                                    ),
                                    dropdownStyleData: DropdownStyleData(
                                      maxHeight: 120,
                                    ),
                                    value:
                                        Provider.of<SelectionStateManagement>(
                                                context)
                                            .selectedFactory,
                                    onChanged: (value) async {
                                      Provider.of<SelectionStateManagement>(
                                              context,
                                              listen: false)
                                          .factorySelection(value);
                                      await Provider.of<
                                                  DataFetchStateManagement>(
                                              context,
                                              listen: false)
                                          .getDataAdmin(context);
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
                          ),
                          SizedBox(width: 10),
                          SizedBox(
                            height: 40,
                            child: Card(
                              color: Styles().palletes1,
                              elevation: 2,
                              shape: RoundedRectangleBorder(
                                  borderRadius: BorderRadius.circular(5)),
                              child: DropdownButtonHideUnderline(
                                child: DropdownButton2(
                                    style: TextStyle(
                                      fontSize: 12,
                                    ),
                                    iconStyleData: IconStyleData(
                                      iconSize: 16,
                                    ),
                                    dropdownStyleData: DropdownStyleData(
                                      maxHeight: 120,
                                    ),
                                    value:
                                        Provider.of<SelectionStateManagement>(
                                                context)
                                            .selectedFloor,
                                    onChanged: (value) async {
                                      Provider.of<SelectionStateManagement>(
                                              context,
                                              listen: false)
                                          .floorSelection(value);
                                      await Provider.of<
                                                  DataFetchStateManagement>(
                                              context,
                                              listen: false)
                                          .getDataAdmin(context);
                                    },
                                    items: [
                                      DropdownMenuItem(
                                        child: Text("Floor-1"),
                                        value: "Floor-1",
                                      ),
                                      DropdownMenuItem(
                                        child: Text("Floor-2"),
                                        value: "Floor-2",
                                      ),
                                      DropdownMenuItem(
                                        child: Text("Floor-3"),
                                        value: "Floor-3",
                                      ),
                                      DropdownMenuItem(
                                        child: Text("Floor-4"),
                                        value: "Floor-4",
                                      ),
                                      DropdownMenuItem(
                                        child: Text("Floor-5"),
                                        value: "Floor-5",
                                      ),
                                    ]),
                              ),
                            ),
                          ),
                        ],
                      ),
                    ],
                  ),

                  // Date -------------------------------->
                  Row(
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
                                  .selectedDate,
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
                                  Provider.of<SelectionStateManagement>(context,
                                          listen: false)
                                      .dateSelection(pickDate);
                                  await Provider.of<DataFetchStateManagement>(
                                          context,
                                          listen: false)
                                      .getDataAdmin(context);
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
                ],
              ),
            ),

            Expanded(
              child: Provider.of<DataFetchStateManagement>(context).isLoaded ==
                      false
                  ? Center(
                      child: Text("Select Factory, Floor and Date"),
                    )
                  : Provider.of<DataFetchStateManagement>(context)
                          .packagingData
                          .isEmpty
                      ? Center(
                          child: Text("No Data Available"),
                        )
                      : DatatableWidget(
                          data: Provider.of<DataFetchStateManagement>(context,
                                  listen: false)
                              .packagingData),
            ),
          ],
        ),
      ),
    );
  }
}
