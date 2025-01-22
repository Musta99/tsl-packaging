import 'package:flutter/foundation.dart';

class LeftBarStateManagement extends ChangeNotifier {
  bool isHomeActive = true;
  bool isDataActive = false;
  bool isDateActive = false;
  bool isPOWiseActive = false;
  bool isSOWiseActive = false;
  bool isSeasonWiseActive = false;
  bool isStyleWiseActive = false;
  bool isDeletedPOActive = false;

  void activeHome() {
    isHomeActive = true;
    isDataActive = false;
    isDateActive = false;
    isPOWiseActive = false;
    isSOWiseActive = false;
    isSeasonWiseActive = false;
    isStyleWiseActive = false;
    isDeletedPOActive = false;

    notifyListeners();
  }

  void activeData() {
    isHomeActive = false;
    isDataActive = true;
    isDateActive = false;
    isPOWiseActive = false;
    isSOWiseActive = false;
    isSeasonWiseActive = false;
    isStyleWiseActive = false;
    isDeletedPOActive = false;

    notifyListeners();
  }

  void activeDate() {
    isHomeActive = false;
    isDataActive = false;
    isDateActive = true;
    isPOWiseActive = false;
    isSOWiseActive = false;
    isSeasonWiseActive = false;
    isStyleWiseActive = false;
    isDeletedPOActive = false;

    notifyListeners();
  }

  void activePO() {
    isHomeActive = false;
    isDataActive = false;
    isDateActive = false;
    isPOWiseActive = true;
    isSOWiseActive = false;
    isSeasonWiseActive = false;
    isStyleWiseActive = false;
    isDeletedPOActive = false;

    notifyListeners();
  }

  void activeSO() {
    isHomeActive = false;
    isDataActive = false;
    isDateActive = false;
    isPOWiseActive = false;
    isSOWiseActive = true;
    isSeasonWiseActive = false;
    isStyleWiseActive = false;
    isDeletedPOActive = false;

    notifyListeners();
  }

  void activeSeason() {
    isHomeActive = false;
    isDataActive = false;
    isDateActive = false;
    isPOWiseActive = false;
    isSOWiseActive = false;
    isSeasonWiseActive = true;
    isStyleWiseActive = false;
    isDeletedPOActive = false;

    notifyListeners();
  }

  void activeStyle() {
    isHomeActive = false;
    isDataActive = false;
    isDateActive = false;
    isPOWiseActive = false;
    isSOWiseActive = false;
    isSeasonWiseActive = false;
    isStyleWiseActive = true;
    isDeletedPOActive = false;

    notifyListeners();
  }

  void activeDeleted() {
    isHomeActive = false;
    isDataActive = false;
    isDateActive = false;
    isPOWiseActive = false;
    isSOWiseActive = false;
    isSeasonWiseActive = false;
    isStyleWiseActive = false;
    isDeletedPOActive = true;

    notifyListeners();
  }
}
