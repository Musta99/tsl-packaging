import 'package:cloud_firestore/cloud_firestore.dart';
import 'package:flutter/material.dart';
import 'package:flutter_easyloading/flutter_easyloading.dart';
import 'package:google_fonts/google_fonts.dart';
import 'package:provider/provider.dart';
import 'package:tsl_packaging/State_Management/selection_state_management.dart';

class DataFetchStateManagement extends ChangeNotifier {
  List<String> factoryList = [
    "TSL-1",
    "TSL-2",
    "TSL-3",
    "TSL-4",
  ];
  // -----------------> Fetch Month List <-----------------------

  List monthList = [];

  void fetchMonthName() async {
    DocumentSnapshot documentSnapshot = await FirebaseFirestore.instance
        .collection("Information")
        .doc("monthName")
        .get();

    monthList = documentSnapshot["month"];
    notifyListeners();
  }

  // -----------------> Recently Added PO for showing in Dashboard <-------------------------
  List recentlyAddedPO = [];

  Future fetchData(BuildContext context) async {
    try {
      QuerySnapshot querySnapshot = await FirebaseFirestore.instance
          .collection(
              Provider.of<SelectionStateManagement>(context, listen: false)
                  .selectedFactory)
          .orderBy("dateOrder")
          .limit(10)
          .get();
      recentlyAddedPO = querySnapshot.docs;
      notifyListeners();
    } catch (err) {
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(
          backgroundColor: Colors.red,
          content: Text(
            err.toString(),
            style: GoogleFonts.roboto(
              fontWeight: FontWeight.bold,
            ),
          ),
        ),
      );
    }
  }

  // -----------------> Fetch Data for Data Details View <-------------------------

  List packagingData = [];
  bool isLoaded = false;
  Future getDataAdmin(BuildContext context) async {
    try {
      isLoaded;
      QuerySnapshot querySnapshot = await FirebaseFirestore.instance
          .collection(
              Provider.of<SelectionStateManagement>(context, listen: false)
                  .selectedFactory)
          .where("floorName",
              isEqualTo:
                  Provider.of<SelectionStateManagement>(context, listen: false)
                      .selectedFloor)
          .where("date",
              isEqualTo:
                  Provider.of<SelectionStateManagement>(context, listen: false)
                      .selectedDate)
          .get();

      packagingData = querySnapshot.docs;
      notifyListeners();

      isLoaded = true;
      notifyListeners();
    } catch (err) {
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(
          backgroundColor: Colors.red,
          content: Text(
            err.toString(),
            style: GoogleFonts.roboto(
              fontWeight: FontWeight.bold,
            ),
          ),
        ),
      );
    }
  }

  // -----------------> Fetch Data for Data PO Wise View <-------------------------

  List packagingDataPOWise = [];
  bool isLoadedPOWise = false;

  Future getDataAdminPOWise(BuildContext context, poNumber) async {
    try {
      isLoadedPOWise;
      for (int i = 0; i < factoryList.length; i++) {
        QuerySnapshot querySnapshot = await FirebaseFirestore.instance
            .collection(factoryList[i])
            .where("productionOrderNo", isEqualTo: poNumber)
            .get();

        packagingDataPOWise.addAll(querySnapshot.docs);
      }

      isLoadedPOWise = true;
      notifyListeners();
    } catch (err) {
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(
          backgroundColor: Colors.red,
          content: Text(
            err.toString(),
            style: GoogleFonts.roboto(
              fontWeight: FontWeight.bold,
            ),
          ),
        ),
      );
    }
  }

  // -----------------> Fetch Data for Data Season Wise View <-------------------------

  List packagingDataSeasonWise = [];
  bool isLoadedSeasonWise = false;

  Future getDataAdminSeasonWise(BuildContext context, monthName) async {
    try {
      isLoadedSeasonWise;
      for (int i = 0; i < factoryList.length; i++) {
        QuerySnapshot querySnapshot = await FirebaseFirestore.instance
            .collection(factoryList[i])
            .where("month", isEqualTo: monthName)
            .get();

        packagingDataSeasonWise.addAll(querySnapshot.docs);
      }

      isLoadedSeasonWise = true;
      notifyListeners();
    } catch (err) {
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(
          backgroundColor: Colors.red,
          content: Text(
            err.toString(),
            style: GoogleFonts.roboto(
              fontWeight: FontWeight.bold,
            ),
          ),
        ),
      );
    }
  }

  // -----------------> Fetch Data for Style Wise View <-------------------------
  List packagingDataStyleWise = [];

  bool isDataLoaded = false;
  int totalGrCumQty = 0;
  int totalBalanceQty = 0;

  int calculateTotalGrCumQty(List packagingDataStyleWise) {
    int totalGr = 0;

    for (var doc in packagingDataStyleWise) {
      // Ensure the field exists and is an integer

      totalGr += (doc['totalGrCumQty'] as num).toInt();
    }

    return totalGr;
  }

  int calculateTotalBalance(List packagingDataStyleWise) {
    int balance = 0;

    for (var doc in packagingDataStyleWise) {
      // Ensure the field exists and is an integer

      balance += (doc['balanceQty'] as num).toInt();
    }

    return balance;
  }

  Future<void> fetchDataStyleWise(
      BuildContext context, String styleName) async {
    packagingDataStyleWise.clear(); // Clear the list to avoid duplicate data
    isDataLoaded;
    try {
      for (int i = 0; i < factoryList.length; i++) {
        QuerySnapshot tslQuery = await FirebaseFirestore.instance
            .collection(factoryList[i])
            .where("style", isEqualTo: styleName)
            .get();

        // Add documents directly to the list
        packagingDataStyleWise.addAll(tslQuery.docs);
      }

      isDataLoaded = true;
      notifyListeners();

      print("Total Records: ${packagingDataStyleWise.length}");
      // Calculate the total of "totalGrCumQty"
      totalGrCumQty = calculateTotalGrCumQty(packagingDataStyleWise);
      totalBalanceQty = calculateTotalBalance(packagingDataStyleWise);
      notifyListeners();

      print(totalGrCumQty.toString());
      print(totalBalanceQty.toString());
    } catch (e) {
      print("Error fetching data: $e");
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text("Error fetching data: $e")),
      );
    }
  }
}
