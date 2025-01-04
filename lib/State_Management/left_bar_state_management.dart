import 'package:flutter/foundation.dart';

class LeftBarStateManagement extends ChangeNotifier {
  bool isHomeActive = true;
  bool isDataActive = false;
  bool isPOWiseActive = false;
  bool isSeasonWiseActive = false;
  bool isStyleWiseActive = false;
  bool isDeletedPOActive = false;

  void activeHome() {
    isHomeActive = true;
    isDataActive = false;
    isPOWiseActive = false;
    isSeasonWiseActive = false;
    isStyleWiseActive = false;
    isDeletedPOActive = false;

    notifyListeners();
  }

  void activeData() {
    isHomeActive = false;
    isDataActive = true;
    isPOWiseActive = false;
    isSeasonWiseActive = false;
    isStyleWiseActive = false;
    isDeletedPOActive = false;

    notifyListeners();
  }

  void activePO() {
    isHomeActive = false;
    isDataActive = false;
    isPOWiseActive = true;
    isSeasonWiseActive = false;
    isStyleWiseActive = false;
    isDeletedPOActive = false;

    notifyListeners();
  }

  void activeSeason() {
    isHomeActive = false;
    isDataActive = false;
    isPOWiseActive = false;
    isSeasonWiseActive = true;
    isStyleWiseActive = false;
    isDeletedPOActive = false;

    notifyListeners();
  }

  void activeStyle() {
    isHomeActive = false;
    isDataActive = false;
    isPOWiseActive = false;
    isSeasonWiseActive = false;
    isStyleWiseActive = true;
    isDeletedPOActive = false;

    notifyListeners();
  }

  void activeDeleted() {
    isHomeActive = false;
    isDataActive = false;
    isPOWiseActive = false;
    isSeasonWiseActive = false;
    isStyleWiseActive = false;
    isDeletedPOActive = true;

    notifyListeners();
  }
}
