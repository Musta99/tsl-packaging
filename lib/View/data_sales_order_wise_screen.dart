import 'package:flutter/foundation.dart';
import 'package:flutter/material.dart';
import 'package:flutter/widgets.dart';
import 'package:google_fonts/google_fonts.dart';
import 'package:provider/provider.dart';
import 'package:tsl_packaging/Utilities/styles.dart';

import '../Services/excel_generator.dart';
import '../State_Management/data_fetch_state_management.dart';
import '../Utilities/Widgets/datatable_widget.dart';

class DataSalesOrderWiseScreen extends StatefulWidget {
  const DataSalesOrderWiseScreen({super.key});

  @override
  State<DataSalesOrderWiseScreen> createState() =>
      _DataSalesOrderWiseScreenState();
}

class _DataSalesOrderWiseScreenState extends State<DataSalesOrderWiseScreen> {
  TextEditingController soController = TextEditingController();

  @override
  void initState() {
    // TODO: implement initState
    super.initState();
  }

  @override
  void dispose() {
    soController.dispose();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      floatingActionButton: FloatingActionButton.small(
        onPressed: () {
          ExcelGenerator().createExcelForSalesOrderWise(
            Provider.of<DataFetchStateManagement>(context, listen: false)
                .packagingDataSOWise,
            soController.text.trim().toUpperCase(),
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
            SelectableText("Work Center Report Sales Order"),
            Container(
              child: Row(
                mainAxisAlignment: MainAxisAlignment.end,
                children: [
                  Container(
                    height: 45,
                    width: MediaQuery.of(context).size.width * 0.3,
                    child: Row(
                      children: [
                        Expanded(
                          child: Card(
                            color: Styles().backgroundColor,
                            elevation: 3,
                            shape: RoundedRectangleBorder(
                              borderRadius: BorderRadius.circular(5),
                            ),
                            child: TextFormField(
                              controller: soController,
                              style: TextStyle(
                                fontSize: 14,
                              ),
                              decoration: InputDecoration(
                                hintText: "Provide SO Number",
                                hintStyle: TextStyle(
                                  color: Colors.grey.withOpacity(0.5),
                                ),
                                contentPadding: EdgeInsets.only(
                                  left: 15,
                                  bottom: 10,
                                ),
                                border: InputBorder.none,
                              ),
                            ),
                          ),
                        ),
                        GestureDetector(
                          onTap: () async {
                            await Provider.of<DataFetchStateManagement>(context,
                                    listen: false)
                                .getDataAdminSOWise(context,
                                    soController.text.trim().toUpperCase());

                            print(soController.text);
                          },
                          child: Card(
                            elevation: 3,
                            color: Styles().palletes4,
                            shape: RoundedRectangleBorder(
                                borderRadius: BorderRadius.circular(5)),
                            child: SizedBox(
                              width: MediaQuery.of(context).size.width * 0.045,
                              child: Center(
                                child: Text(
                                  "Search",
                                  style: GoogleFonts.ptSerif(
                                    color: Colors.white,
                                    fontSize: 12,
                                  ),
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
                          .isLoadedSOWise ==
                      false
                  ? Center(
                      child: Text("Provide Sales Order Number"),
                    )
                  : Provider.of<DataFetchStateManagement>(context)
                          .packagingDataSOWise
                          .isEmpty
                      ? Center(
                          child: Text("No Data Available"),
                        )
                      : DatatableWidget(
                          data: Provider.of<DataFetchStateManagement>(context,
                                  listen: false)
                              .packagingDataSOWise),
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
