import 'package:flutter/material.dart';

class SelectionStateManagement extends ChangeNotifier {
  String selectedFactory = "TSL-1";

  void factorySelection(value) {
    selectedFactory = value;
    notifyListeners();
  }

  String selectedFactoryPOWise = "TSL-1";

  void factorySelectionPOWise(value) {
    selectedFactoryPOWise = value;
    notifyListeners();
  }

// ------------------------> <------------------------
  String selectedFloor = "Floor-1";

  void floorSelection(value) {
    selectedFloor = value;
    notifyListeners();
  }

// ------------------------> <------------------------
  String selectedDate =
      "${DateTime.now().day.toString()}-${DateTime.now().month.toString()}-${DateTime.now().year.toString()}";

  void dateSelection(DateTime date) {
    selectedDate =
        "${date.day.toString()}-${date.month.toString()}-${date.year.toString()}";
    notifyListeners();
  }

  // ---------------------> <----------------------------
  String? selectedMonth;

  void selectMonth(String value) {
    selectedMonth = value;
    notifyListeners();
  }
}
