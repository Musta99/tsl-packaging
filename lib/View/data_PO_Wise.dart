import 'package:dropdown_button2/dropdown_button2.dart';
import 'package:flutter/material.dart';
import 'package:font_awesome_flutter/font_awesome_flutter.dart';
import 'package:google_fonts/google_fonts.dart';
import 'package:modern_textfield/modern_textfield.dart';
import 'package:provider/provider.dart';

import '../State_Management/data_fetch_state_management.dart';
import '../State_Management/selection_state_management.dart';
import '../Utilities/Widgets/datatable_widget.dart';
import '../Utilities/styles.dart';

class DataDetailsPOWise extends StatefulWidget {
  DataDetailsPOWise({super.key});

  @override
  State<DataDetailsPOWise> createState() => _DataDetailsPOWiseState();
}

class _DataDetailsPOWiseState extends State<DataDetailsPOWise> {
  TextEditingController poController = TextEditingController();

  @override
  void initState() {
    // TODO: implement initState
    super.initState();
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      floatingActionButton: FloatingActionButton.small(
        backgroundColor: Styles().palletes3,
        onPressed: () {},
        child: Image.asset(
          "assets/images/excel.png",
          height: 30,
          width: 30,
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
                              controller: poController,
                              style: TextStyle(
                                fontSize: 14,
                              ),
                              decoration: InputDecoration(
                                hintText: "Provide PO Number",
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
                                .getDataAdminPOWise(context,
                                    poController.text.trim().toUpperCase());

                            print(poController.text);
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
                          .isLoadedPOWise ==
                      false
                  ? Center(
                      child: Text("Select Factory and Provide PO"),
                    )
                  : Provider.of<DataFetchStateManagement>(context)
                          .packagingDataPOWise
                          .isEmpty
                      ? Center(
                          child: Text("No Data Available"),
                        )
                      : DatatableWidget(
                          data: Provider.of<DataFetchStateManagement>(context,
                                  listen: false)
                              .packagingDataPOWise),
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
