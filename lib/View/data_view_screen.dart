import 'package:cloud_firestore/cloud_firestore.dart';
import 'package:flutter/foundation.dart';
import 'package:flutter/material.dart';
import 'package:font_awesome_flutter/font_awesome_flutter.dart';
import 'package:google_fonts/google_fonts.dart';
import 'package:provider/provider.dart';
import 'package:tsl_packaging/State_Management/left_bar_state_management.dart';
import 'package:tsl_packaging/Utilities/Widgets/left_bar_button.dart';
import 'package:tsl_packaging/Utilities/styles.dart';
import 'package:tsl_packaging/View/dashboard_screen.dart';
import 'package:tsl_packaging/View/data_PO_Wise.dart';
import 'package:tsl_packaging/View/data_date_wise_screen.dart';
import 'package:tsl_packaging/View/data_details_view.dart';
import 'package:tsl_packaging/View/data_sales_order_wise_screen.dart';
import 'package:tsl_packaging/View/data_season_wise.dart';
import 'package:tsl_packaging/View/data_style_wise.dart';
import 'package:tsl_packaging/View/deleted_data_screen.dart';

class DataViewScreen extends StatefulWidget {
  const DataViewScreen({super.key});

  @override
  State<DataViewScreen> createState() => _DataViewScreenState();
}

class _DataViewScreenState extends State<DataViewScreen> {
  void fetchData() async {
    QuerySnapshot querySnapshot =
        await FirebaseFirestore.instance.collection("Packaging").get();

    print(querySnapshot.docs[0]["L"]["Order Qty"].toString());
  }

  @override
  void initState() {
    // TODO: implement initState
    super.initState();
    fetchData();
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      backgroundColor: Colors.white,
      body: Container(
        height: MediaQuery.of(context).size.height,
        width: MediaQuery.of(context).size.width,
        child: Row(
          children: [
            // ---------------------- Left Side Bar ------------------

            SingleChildScrollView(
              child: Container(
                width: MediaQuery.of(context).size.width * 0.2,
                height: MediaQuery.of(context).size.height,
                decoration: BoxDecoration(
                  color: Styles().palletes4,
                  borderRadius: BorderRadius.only(
                    topRight: Radius.circular(15),
                    bottomRight: Radius.circular(15),
                  ),
                ),
                child: Container(
                  margin: EdgeInsets.only(
                    top: 15,
                  ),
                  child: Column(
                    mainAxisAlignment: MainAxisAlignment.spaceBetween,
                    children: [
                      Column(
                        crossAxisAlignment: CrossAxisAlignment.start,
                        children: [
                          Align(
                            alignment: Alignment.center,
                            child: SelectableText(
                              "EasyPack",
                              style: GoogleFonts.pacifico(
                                color: Colors.white,
                                fontSize:
                                    MediaQuery.of(context).size.width * 0.015,
                              ),
                            ),
                          ),
                          SizedBox(
                            height: MediaQuery.of(context).size.height * 0.08,
                          ),
                          GestureDetector(
                            onTap: () {
                              Provider.of<LeftBarStateManagement>(context,
                                      listen: false)
                                  .activeHome();
                            },
                            child: LeftBarCommonButton(
                              iconOpacity:
                                  Provider.of<LeftBarStateManagement>(context)
                                              .isHomeActive ==
                                          true
                                      ? 1
                                      : 0.3,
                              textOpacity:
                                  Provider.of<LeftBarStateManagement>(context)
                                              .isHomeActive ==
                                          true
                                      ? 1
                                      : 0.3,
                              icon: FontAwesomeIcons.home,
                              title: "Home",
                            ),
                          ),
                          SizedBox(
                            height: MediaQuery.of(context).size.height * 0.02,
                          ),
                          GestureDetector(
                            onTap: () {
                              Provider.of<LeftBarStateManagement>(context,
                                      listen: false)
                                  .activeData();
                            },
                            child: LeftBarCommonButton(
                              iconOpacity:
                                  Provider.of<LeftBarStateManagement>(context)
                                              .isDataActive ==
                                          true
                                      ? 1
                                      : 0.3,
                              textOpacity:
                                  Provider.of<LeftBarStateManagement>(context)
                                              .isDataActive ==
                                          true
                                      ? 1
                                      : 0.3,
                              icon: FontAwesomeIcons.database,
                              title: "Daily Report",
                            ),
                          ),
                          SizedBox(
                            height: MediaQuery.of(context).size.height * 0.02,
                          ),
                          GestureDetector(
                            onTap: () {
                              Provider.of<LeftBarStateManagement>(context,
                                      listen: false)
                                  .activeDate();
                            },
                            child: LeftBarCommonButton(
                              iconOpacity:
                                  Provider.of<LeftBarStateManagement>(context)
                                              .isDateActive ==
                                          true
                                      ? 1
                                      : 0.3,
                              textOpacity:
                                  Provider.of<LeftBarStateManagement>(context)
                                              .isDateActive ==
                                          true
                                      ? 1
                                      : 0.3,
                              icon: Icons.weekend,
                              title: "Date Wise",
                            ),
                          ),
                          SizedBox(
                            height: MediaQuery.of(context).size.height * 0.02,
                          ),
                          GestureDetector(
                            onTap: () {
                              Provider.of<LeftBarStateManagement>(context,
                                      listen: false)
                                  .activePO();
                            },
                            child: LeftBarCommonButton(
                              iconOpacity:
                                  Provider.of<LeftBarStateManagement>(context)
                                              .isPOWiseActive ==
                                          true
                                      ? 1
                                      : 0.3,
                              textOpacity:
                                  Provider.of<LeftBarStateManagement>(context)
                                              .isPOWiseActive ==
                                          true
                                      ? 1
                                      : 0.3,
                              icon: FontAwesomeIcons.listNumeric,
                              title: "Prod Order Wise",
                            ),
                          ),
                          SizedBox(
                            height: MediaQuery.of(context).size.height * 0.02,
                          ),
                          GestureDetector(
                            onTap: () {
                              Provider.of<LeftBarStateManagement>(context,
                                      listen: false)
                                  .activeSO();
                            },
                            child: LeftBarCommonButton(
                              icon: FontAwesomeIcons.salesforce,
                              title: "Sales Order Wise",
                              iconOpacity:
                                  Provider.of<LeftBarStateManagement>(context)
                                              .isSOWiseActive ==
                                          true
                                      ? 1
                                      : 0.3,
                              textOpacity:
                                  Provider.of<LeftBarStateManagement>(context)
                                              .isSOWiseActive ==
                                          true
                                      ? 1
                                      : 0.3,
                            ),
                          ),
                          SizedBox(
                            height: MediaQuery.of(context).size.height * 0.02,
                          ),
                          GestureDetector(
                            onTap: () {
                              Provider.of<LeftBarStateManagement>(context,
                                      listen: false)
                                  .activeStyle();
                            },
                            child: LeftBarCommonButton(
                              iconOpacity:
                                  Provider.of<LeftBarStateManagement>(context)
                                              .isStyleWiseActive ==
                                          true
                                      ? 1
                                      : 0.3,
                              textOpacity:
                                  Provider.of<LeftBarStateManagement>(context)
                                              .isStyleWiseActive ==
                                          true
                                      ? 1
                                      : 0.3,
                              icon: Icons.style,
                              title: "Style Wise",
                            ),
                          ),
                          SizedBox(
                            height: MediaQuery.of(context).size.height * 0.02,
                          ),
                          GestureDetector(
                            onTap: () {
                              Provider.of<LeftBarStateManagement>(context,
                                      listen: false)
                                  .activeSeason();
                            },
                            child: LeftBarCommonButton(
                              iconOpacity:
                                  Provider.of<LeftBarStateManagement>(context)
                                              .isSeasonWiseActive ==
                                          true
                                      ? 1
                                      : 0.3,
                              textOpacity:
                                  Provider.of<LeftBarStateManagement>(context)
                                              .isSeasonWiseActive ==
                                          true
                                      ? 1
                                      : 0.3,
                              icon: FontAwesomeIcons.calendarAlt,
                              title: "Season Wise",
                            ),
                          ),
                          SizedBox(
                            height: MediaQuery.of(context).size.height * 0.02,
                          ),
                          GestureDetector(
                            onTap: () {
                              Provider.of<LeftBarStateManagement>(context,
                                      listen: false)
                                  .activeDeleted();
                            },
                            child: LeftBarCommonButton(
                              iconOpacity:
                                  Provider.of<LeftBarStateManagement>(context)
                                              .isDeletedPOActive ==
                                          true
                                      ? 1
                                      : 0.3,
                              textOpacity:
                                  Provider.of<LeftBarStateManagement>(context)
                                              .isDeletedPOActive ==
                                          true
                                      ? 1
                                      : 0.3,
                              icon: FontAwesomeIcons.eraser,
                              title: "Deleted PO",
                            ),
                          ),
                        ],
                      ),
                      Align(
                        alignment: Alignment.center,
                        child: SelectableText(
                          "Developed by: Must@fiz",
                          style: GoogleFonts.montserrat(
                            color: Colors.white,
                            fontSize: 10,
                          ),
                        ),
                      ),
                    ],
                  ),
                ),
              ),
            ),

            // ---------------------- Right Side Bar ------------------

            Container(
              width: MediaQuery.of(context).size.width * 0.8,
              height: MediaQuery.of(context).size.height,
              color: Styles().palletes1,
              child: Provider.of<LeftBarStateManagement>(context).isHomeActive
                  ? DashboardScreen()
                  : Provider.of<LeftBarStateManagement>(context).isDataActive
                      ? DataDetailsView()
                      : Provider.of<LeftBarStateManagement>(context)
                              .isPOWiseActive
                          ? DataDetailsPOWise()
                          : Provider.of<LeftBarStateManagement>(context)
                                  .isSeasonWiseActive
                              ? DataSeasonWise()
                              : Provider.of<LeftBarStateManagement>(context)
                                      .isStyleWiseActive
                                  ? DataStyleWise()
                                  : Provider.of<LeftBarStateManagement>(context)
                                          .isSOWiseActive
                                      ? DataSalesOrderWiseScreen()
                                      : Provider.of<LeftBarStateManagement>(
                                                  context)
                                              .isDateActive
                                          ? DataDateWiseScreen()
                                          : DeletedDataScreen(),
            ),
          ],
        ),
      ),
    );
  }
}
