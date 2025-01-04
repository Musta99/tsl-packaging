import 'package:flutter/material.dart';
import 'package:flutter_easyloading/flutter_easyloading.dart';
import 'package:google_fonts/google_fonts.dart';
import 'package:provider/provider.dart';
import 'package:tsl_packaging/Utilities/styles.dart';

import '../State_Management/data_fetch_state_management.dart';
import '../Utilities/Widgets/datatable_widget.dart';

class DataStyleWise extends StatelessWidget {
  DataStyleWise({super.key});

  TextEditingController styleController = TextEditingController();

  @override
  Widget build(BuildContext context) {
    return Scaffold(
        body: Container(
      margin: EdgeInsets.symmetric(
        horizontal: 20,
        vertical: 20,
      ),
      child: Column(
        children: [
          // Name Title and Search Box According PO No: -------------------------------->
          SelectableText("Daily Work Center Report"),
          Row(
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
                          controller: styleController,
                          style: TextStyle(
                            fontSize: 14,
                          ),
                          decoration: InputDecoration(
                            hintText: "Provide Style Name",
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
                            .fetchDataStyleWise(context,
                                styleController.text.trim().toUpperCase());

                        print(styleController.text);
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

          Expanded(
            child: Provider.of<DataFetchStateManagement>(context)
                    .packagingDataStyleWise
                    .isEmpty
                ? Center(
                    child: Text("No Data Found"),
                  )
                : Provider.of<DataFetchStateManagement>(context).isDataLoaded ==
                        false
                    ? Center(
                        child: CircularProgressIndicator(),
                      )
                    : DatatableWidget(
                        data: Provider.of<DataFetchStateManagement>(context,
                                listen: false)
                            .packagingDataStyleWise),
          ),

          Card(
            shape:
                RoundedRectangleBorder(borderRadius: BorderRadius.circular(5)),
            elevation: 3,
            child: ListTile(
              title: Text("Total Finished Quantity: "),
              trailing: Text(
                Provider.of<DataFetchStateManagement>(context, listen: false)
                    .totalGrCumQty
                    .toString(),
                style: GoogleFonts.ptSerif(
                  fontWeight: FontWeight.bold,
                ),
              ),
            ),
          ),

          Card(
            shape:
                RoundedRectangleBorder(borderRadius: BorderRadius.circular(5)),
            elevation: 3,
            child: ListTile(
              title: Text("Remaining Quantity: "),
              trailing: Text(
                Provider.of<DataFetchStateManagement>(context, listen: false)
                    .totalBalanceQty
                    .toString(),
                style: GoogleFonts.ptSerif(
                  fontWeight: FontWeight.bold,
                ),
              ),
            ),
          ),
        ],
      ),
    ));
  }
}
