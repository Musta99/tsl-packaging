import 'package:cloud_firestore/cloud_firestore.dart';

class DataModel {
  String factoryName;
  String floorName;
  String lineName;
  String buyer;
  String style;
  String buy;
  String salesOrderNo;
  String item;
  String fGMatNo;
  String ourNoFR;
  String colorName;
  String countryName;
  String productionOrderNo;
  int orderQty;
  int issueQty;
  XXS? xXS;
  XXS? xS;
  XXS? s;
  XXS? m;
  XXS? l;
  XXS? xL;
  XXS? xXL;
  XXS? xXXL;
  int dGrQty;
  int totalGrCumQty;
  int balanceGrQty;
  int dailyPackQty;
  int lastDayPackCumQty;
  int balanceQty;
  String remarks;
  String date;
  int dateOrder;

  DataModel(
      {required this.factoryName,
      required this.floorName,
      required this.lineName,
      required this.buyer,
      required this.style,
      required this.buy,
      required this.salesOrderNo,
      required this.item,
      required this.fGMatNo,
      required this.ourNoFR,
      required this.colorName,
      required this.countryName,
      required this.productionOrderNo,
      required this.orderQty,
      required this.issueQty,
      this.xXS,
      this.xS,
      this.s,
      this.m,
      this.l,
      this.xL,
      this.xXL,
      this.xXXL,
      required this.dGrQty,
      required this.totalGrCumQty,
      required this.balanceGrQty,
      required this.dailyPackQty,
      required this.lastDayPackCumQty,
      required this.balanceQty,
      required this.remarks,
      required this.date,
      required this.dateOrder});

  DataModel.fromJson(Map<String, dynamic> json)
      : factoryName = json['factoryName'],
        floorName = json['floorName'],
        lineName = json['lineName'],
        buyer = json['Buyer'],
        style = json['Style'],
        buy = json['Buy'],
        salesOrderNo = json['Sales Order No'],
        item = json['Item'],
        fGMatNo = json['FG(Mat No)'],
        ourNoFR = json['Our No(FR)'],
        colorName = json['Color Name'],
        countryName = json['Country Name'],
        productionOrderNo = json['Production Order No'],
        orderQty = json['Order Qty'],
        issueQty = json['Issue Qty'],
        xXS = json['XXS'] != null ? XXS.fromJson(json['XXS']) : null,
        xS = json['XS'] != null ? XXS.fromJson(json['XS']) : null,
        s = json['S'] != null ? XXS.fromJson(json['S']) : null,
        m = json['M'] != null ? XXS.fromJson(json['M']) : null,
        l = json['L'] != null ? XXS.fromJson(json['L']) : null,
        xL = json['XL'] != null ? XXS.fromJson(json['XL']) : null,
        xXL = json['XXL'] != null ? XXS.fromJson(json['XXL']) : null,
        xXXL = json['XXXL'] != null ? XXS.fromJson(json['XXXL']) : null,
        dGrQty = json['D Gr Qty'],
        totalGrCumQty = json['Total Gr Cum Qty'],
        balanceGrQty = json['Balance Gr Qty'],
        dailyPackQty = json['Daily Pack Qty'],
        lastDayPackCumQty = json['Last Day Pack Cum Qty'],
        balanceQty = json['Balance Qty'],
        remarks = json['Remarks'],
        date = json['date'],
        dateOrder = json['dateOrder'];

  Map<String, dynamic> toJson() {
    final data = <String, dynamic>{};
    data['factoryName'] = factoryName;
    data['floorName'] = floorName;
    data['lineName'] = lineName;
    data['Buyer'] = buyer;
    data['Style'] = style;
    data['Buy'] = buy;
    data['Sales Order No'] = salesOrderNo;
    data['Item'] = item;
    data['FG(Mat No)'] = fGMatNo;
    data['Our No(FR)'] = ourNoFR;
    data['Color Name'] = colorName;
    data['Country Name'] = countryName;
    data['Production Order No'] = productionOrderNo;
    data['Order Qty'] = orderQty;
    data['Issue Qty'] = issueQty;
    if (xXS != null) data['XXS'] = xXS!.toJson();
    if (xS != null) data['XS'] = xS!.toJson();
    if (s != null) data['S'] = s!.toJson();
    if (m != null) data['M'] = m!.toJson();
    if (l != null) data['L'] = l!.toJson();
    if (xL != null) data['XL'] = xL!.toJson();
    if (xXL != null) data['XXL'] = xXL!.toJson();
    if (xXXL != null) data['XXXL'] = xXXL!.toJson();
    data['D Gr Qty'] = dGrQty;
    data['Total Gr Cum Qty'] = totalGrCumQty;
    data['Balance Gr Qty'] = balanceGrQty;
    data['Daily Pack Qty'] = dailyPackQty;
    data['Last Day Pack Cum Qty'] = lastDayPackCumQty;
    data['Balance Qty'] = balanceQty;
    data['Remarks'] = remarks;
    data['date'] = date;
    data['dateOrder'] = dateOrder;
    return data;
  }
}

class XXS {
  int orderQty;
  int issue;
  int receivedQty;
  int totalCum;

  XXS({
    required this.orderQty,
    required this.issue,
    required this.receivedQty,
    required this.totalCum,
  });

  XXS.fromJson(Map<String, dynamic> json)
      : orderQty = json['Order Qty'],
        issue = json['Issue'] ?? '',
        receivedQty = json['Received Qty'] ?? '',
        totalCum = json['Total Cum'];

  Map<String, dynamic> toJson() => {
        'Order Qty': orderQty,
        'Issue': issue,
        'Received Qty': receivedQty,
        'Total Cum': totalCum,
      };
}
