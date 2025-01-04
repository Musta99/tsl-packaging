import 'package:cloud_firestore/cloud_firestore.dart';
import 'package:flutter/foundation.dart';
import 'package:flutter/material.dart';
import 'package:flutter/widgets.dart';

import 'package:tsl_packaging/Model/data_model.dart';
import 'package:tsl_packaging/Utilities/styles.dart';

class DataInputScreen extends StatefulWidget {
  const DataInputScreen({super.key});

  @override
  State<DataInputScreen> createState() => _DataInputScreenState();
}

class _DataInputScreenState extends State<DataInputScreen> {
  TextEditingController factoryController = TextEditingController();
  TextEditingController floorController = TextEditingController();
  TextEditingController lineController = TextEditingController();
  TextEditingController buyerController = TextEditingController();
  TextEditingController styleController = TextEditingController();
  TextEditingController buyController = TextEditingController();
  TextEditingController salesOrderController = TextEditingController();
  TextEditingController itemController = TextEditingController();
  TextEditingController fgMatController = TextEditingController();
  TextEditingController ourNoController = TextEditingController();
  TextEditingController colorController = TextEditingController();
  TextEditingController countryController = TextEditingController();
  TextEditingController prodOrderController = TextEditingController();
  TextEditingController lastDayPackController = TextEditingController();

  TextEditingController xxsOrderQtyController = TextEditingController();
  TextEditingController xxsIssueQtyController = TextEditingController();
  TextEditingController xxsRcvQtyController = TextEditingController();
  TextEditingController xxsCumQtyController = TextEditingController();

  TextEditingController xsOrderQtyController = TextEditingController();
  TextEditingController xsIssueQtyController = TextEditingController();
  TextEditingController xsRcvQtyController = TextEditingController();
  TextEditingController xsCumQtyController = TextEditingController();

  TextEditingController sOrderQtyController = TextEditingController();
  TextEditingController sIssueQtyController = TextEditingController();
  TextEditingController sRcvQtyController = TextEditingController();
  TextEditingController sCumQtyController = TextEditingController();

  TextEditingController lOrderQtyController = TextEditingController();
  TextEditingController lIssueQtyController = TextEditingController();
  TextEditingController lRcvQtyController = TextEditingController();
  TextEditingController lCumQtyController = TextEditingController();

  TextEditingController mOrderQtyController = TextEditingController();
  TextEditingController mIssueQtyController = TextEditingController();
  TextEditingController mRcvQtyController = TextEditingController();
  TextEditingController mCumQtyController = TextEditingController();

  TextEditingController xlOrderQtyController = TextEditingController();
  TextEditingController xlIssueQtyController = TextEditingController();
  TextEditingController xlRcvQtyController = TextEditingController();
  TextEditingController xlCumQtyController = TextEditingController();

  TextEditingController xxlOrderQtyController = TextEditingController();
  TextEditingController xxlIssueQtyController = TextEditingController();
  TextEditingController xxlRcvQtyController = TextEditingController();
  TextEditingController xxlCumQtyController = TextEditingController();

  TextEditingController xxxlOrderQtyController = TextEditingController();
  TextEditingController xxxlIssueQtyController = TextEditingController();
  TextEditingController xxxlRcvQtyController = TextEditingController();
  TextEditingController xxxlCumQtyController = TextEditingController();

  String setDate =
      "${DateTime.now().day.toString()}-${DateTime.now().month.toString()}-${DateTime.now().year.toString()}";

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      backgroundColor: Styles().palletes2,
      body: Container(
        margin: EdgeInsets.symmetric(
          horizontal: 50,
        ),
        child: SingleChildScrollView(
          child: Column(
            children: [
              Text("Final Packaging Information"),

              // ------------------------ Factory Name ------------------------
              Material(
                elevation: 5,
                child: TextFormField(
                  controller: factoryController,
                  decoration: InputDecoration(
                    filled: true,
                    fillColor: Colors.grey.withOpacity(0.5),
                    labelText: "Factory Name",
                    contentPadding: EdgeInsets.only(
                      left: 20,
                    ),
                    border: InputBorder.none,
                  ),
                ),
              ),
              SizedBox(
                height: 15,
              ),

              // ------------------------ Floor Name ------------------------
              Material(
                elevation: 5,
                child: TextFormField(
                  controller: floorController,
                  decoration: InputDecoration(
                    filled: true,
                    fillColor: Colors.grey.withOpacity(0.5),
                    labelText: "Floor Name",
                    contentPadding: EdgeInsets.only(
                      left: 20,
                    ),
                    border: InputBorder.none,
                  ),
                ),
              ),
              SizedBox(
                height: 15,
              ),

              // ------------------------ Line Name ------------------------
              Material(
                elevation: 5,
                child: TextFormField(
                  controller: lineController,
                  decoration: InputDecoration(
                    filled: true,
                    fillColor: Colors.grey.withOpacity(0.5),
                    labelText: "Line Name",
                    contentPadding: EdgeInsets.only(
                      left: 20,
                    ),
                    border: InputBorder.none,
                  ),
                ),
              ),
              SizedBox(
                height: 15,
              ),
              // ------------------------ Date ------------------------

              Material(
                elevation: 5,
                child: ListTile(
                  tileColor: Styles().palletes1,
                  title: Text(setDate),
                  trailing: GestureDetector(
                    onTap: () async {
                      DateTime? pickedDate = await showDatePicker(
                          context: context,
                          firstDate: DateTime(2000),
                          lastDate: DateTime(3000));

                      if (pickedDate != null) {
                        setState(() {
                          setDate =
                              "${pickedDate.day.toString()}-${pickedDate.month.toString()}-${pickedDate.year.toString()}";
                        });
                      }
                    },
                    child: Icon(
                      Icons.calendar_month,
                    ),
                  ),
                ),
              ),

              SizedBox(
                height: 15,
              ),

              // ------------------------ Buyer Name ------------------------
              Material(
                elevation: 5,
                child: TextFormField(
                  controller: buyerController,
                  decoration: InputDecoration(
                    filled: true,
                    fillColor: Colors.grey.withOpacity(0.5),
                    labelText: "Buyer Name",
                    contentPadding: EdgeInsets.only(
                      left: 20,
                    ),
                    border: InputBorder.none,
                  ),
                ),
              ),
              SizedBox(
                height: 15,
              ),
              // ------------------------ Style Name ------------------------
              Material(
                elevation: 5,
                child: TextFormField(
                  controller: styleController,
                  decoration: InputDecoration(
                    filled: true,
                    fillColor: Colors.grey.withOpacity(0.5),
                    labelText: "Style Name",
                    contentPadding: EdgeInsets.only(
                      left: 20,
                    ),
                    border: InputBorder.none,
                  ),
                ),
              ),
              SizedBox(
                height: 15,
              ),
              // ------------------------ Buy Name ------------------------
              Material(
                elevation: 5,
                child: TextFormField(
                  controller: buyController,
                  decoration: InputDecoration(
                    filled: true,
                    fillColor: Colors.grey.withOpacity(0.5),
                    labelText: "Buy Name",
                    contentPadding: EdgeInsets.only(
                      left: 20,
                    ),
                    border: InputBorder.none,
                  ),
                ),
              ),

              SizedBox(
                height: 15,
              ),

              // ------------------------ Sales Order Number ------------------------
              Material(
                elevation: 5,
                child: TextFormField(
                  controller: salesOrderController,
                  decoration: InputDecoration(
                    filled: true,
                    fillColor: Colors.grey.withOpacity(0.5),
                    labelText: "Sales Order Number",
                    contentPadding: EdgeInsets.only(
                      left: 20,
                    ),
                    border: InputBorder.none,
                  ),
                ),
              ),

              SizedBox(
                height: 15,
              ),

              // ------------------------ Item Number ------------------------
              Material(
                elevation: 5,
                child: TextFormField(
                  controller: itemController,
                  decoration: InputDecoration(
                    filled: true,
                    fillColor: Colors.grey.withOpacity(0.5),
                    labelText: "Item Number",
                    contentPadding: EdgeInsets.only(
                      left: 20,
                    ),
                    border: InputBorder.none,
                  ),
                ),
              ),
              SizedBox(
                height: 15,
              ),

              // ------------------------ FG Mat No ------------------------
              Material(
                elevation: 5,
                child: TextFormField(
                  controller: fgMatController,
                  decoration: InputDecoration(
                    filled: true,
                    fillColor: Colors.grey.withOpacity(0.5),
                    labelText: "FG(Mat.) No",
                    contentPadding: EdgeInsets.only(
                      left: 20,
                    ),
                    border: InputBorder.none,
                  ),
                ),
              ),
              SizedBox(
                height: 15,
              ),

              // ------------------------ our No FR ------------------------
              Material(
                elevation: 5,
                child: TextFormField(
                  controller: ourNoController,
                  decoration: InputDecoration(
                    filled: true,
                    fillColor: Colors.grey.withOpacity(0.5),
                    labelText: "Our No",
                    contentPadding: EdgeInsets.only(
                      left: 20,
                    ),
                    border: InputBorder.none,
                  ),
                ),
              ),
              SizedBox(
                height: 15,
              ),

              // ------------------------ Color Name ------------------------
              Material(
                elevation: 5,
                child: TextFormField(
                  controller: colorController,
                  decoration: InputDecoration(
                    filled: true,
                    fillColor: Colors.grey.withOpacity(0.5),
                    labelText: "Color Name",
                    contentPadding: EdgeInsets.only(
                      left: 20,
                    ),
                    border: InputBorder.none,
                  ),
                ),
              ),
              SizedBox(
                height: 15,
              ),

              // ------------------------ Country Name ------------------------
              Material(
                elevation: 5,
                child: TextFormField(
                  controller: countryController,
                  decoration: InputDecoration(
                    filled: true,
                    fillColor: Colors.grey.withOpacity(0.5),
                    labelText: "Country Name",
                    contentPadding: EdgeInsets.only(
                      left: 20,
                    ),
                    border: InputBorder.none,
                  ),
                ),
              ),
              SizedBox(
                height: 15,
              ),

              // ------------------------ Production Order No ------------------------
              Material(
                elevation: 5,
                child: TextFormField(
                  controller: prodOrderController,
                  decoration: InputDecoration(
                    filled: true,
                    fillColor: Colors.grey.withOpacity(0.5),
                    labelText: "Prod Order No",
                    contentPadding: EdgeInsets.only(
                      left: 20,
                    ),
                    border: InputBorder.none,
                  ),
                ),
              ),
              SizedBox(
                height: 15,
              ),

              // ------------------------ Last Day Pack Cum Qty ------------------------
              Material(
                elevation: 5,
                child: TextFormField(
                  controller: lastDayPackController,
                  decoration: InputDecoration(
                    filled: true,
                    fillColor: Colors.grey.withOpacity(0.5),
                    labelText: "Last Day Pack Cum Quantity",
                    contentPadding: EdgeInsets.only(
                      left: 20,
                    ),
                    border: InputBorder.none,
                  ),
                ),
              ),
              SizedBox(
                height: 15,
              ),

              // ------------------------ Size Wise Information ------------------------

              // ============================  XXS =================================
              Row(
                children: [
                  Card(child: Text("XXS : ")),
                  Expanded(
                    child: Material(
                      elevation: 5,
                      child: TextFormField(
                        controller: xxsOrderQtyController,
                        decoration: InputDecoration(
                          filled: true,
                          fillColor: Colors.grey.withOpacity(0.5),
                          labelText: "XXS Order Qty",
                          contentPadding: EdgeInsets.only(
                            left: 20,
                          ),
                          border: InputBorder.none,
                        ),
                      ),
                    ),
                  ),
                  SizedBox(
                    width: 10,
                  ),
                  Expanded(
                    child: Material(
                      elevation: 5,
                      child: TextFormField(
                        controller: xxsIssueQtyController,
                        decoration: InputDecoration(
                          filled: true,
                          fillColor: Colors.grey.withOpacity(0.5),
                          labelText: "XXS Issue Qty",
                          contentPadding: EdgeInsets.only(
                            left: 20,
                          ),
                          border: InputBorder.none,
                        ),
                      ),
                    ),
                  ),
                  SizedBox(
                    width: 10,
                  ),
                  Expanded(
                    child: Material(
                      elevation: 5,
                      child: TextFormField(
                        controller: xxsRcvQtyController,
                        decoration: InputDecoration(
                          filled: true,
                          fillColor: Colors.grey.withOpacity(0.5),
                          labelText: "XXS Rcv Qty",
                          contentPadding: EdgeInsets.only(
                            left: 20,
                          ),
                          border: InputBorder.none,
                        ),
                      ),
                    ),
                  ),
                  SizedBox(
                    width: 10,
                  ),
                  Expanded(
                    child: Material(
                      elevation: 5,
                      child: TextFormField(
                        controller: xxsCumQtyController,
                        decoration: InputDecoration(
                          filled: true,
                          fillColor: Colors.grey.withOpacity(0.5),
                          labelText: "XXS GR Cum Qty",
                          contentPadding: EdgeInsets.only(
                            left: 20,
                          ),
                          border: InputBorder.none,
                        ),
                      ),
                    ),
                  ),
                ],
              ),
              SizedBox(
                height: 15,
              ),

              // ============================  XS =================================
              Row(
                children: [
                  Card(child: Text("XS : ")),
                  Expanded(
                    child: Material(
                      elevation: 5,
                      child: TextFormField(
                        controller: xsOrderQtyController,
                        decoration: InputDecoration(
                          filled: true,
                          fillColor: Colors.grey.withOpacity(0.5),
                          labelText: "XS Order Qty",
                          contentPadding: EdgeInsets.only(
                            left: 20,
                          ),
                          border: InputBorder.none,
                        ),
                      ),
                    ),
                  ),
                  SizedBox(
                    width: 10,
                  ),
                  Expanded(
                    child: Material(
                      elevation: 5,
                      child: TextFormField(
                        controller: xsIssueQtyController,
                        decoration: InputDecoration(
                          filled: true,
                          fillColor: Colors.grey.withOpacity(0.5),
                          labelText: "XS Issue Qty",
                          contentPadding: EdgeInsets.only(
                            left: 20,
                          ),
                          border: InputBorder.none,
                        ),
                      ),
                    ),
                  ),
                  SizedBox(
                    width: 10,
                  ),
                  Expanded(
                    child: Material(
                      elevation: 5,
                      child: TextFormField(
                        controller: xsRcvQtyController,
                        decoration: InputDecoration(
                          filled: true,
                          fillColor: Colors.grey.withOpacity(0.5),
                          labelText: "XS Rcv Qty",
                          contentPadding: EdgeInsets.only(
                            left: 20,
                          ),
                          border: InputBorder.none,
                        ),
                      ),
                    ),
                  ),
                  SizedBox(
                    width: 10,
                  ),
                  Expanded(
                    child: Material(
                      elevation: 5,
                      child: TextFormField(
                        controller: xsCumQtyController,
                        decoration: InputDecoration(
                          filled: true,
                          fillColor: Colors.grey.withOpacity(0.5),
                          labelText: "XS GR Cum Qty",
                          contentPadding: EdgeInsets.only(
                            left: 20,
                          ),
                          border: InputBorder.none,
                        ),
                      ),
                    ),
                  ),
                ],
              ),
              SizedBox(
                height: 15,
              ),

              // ============================  S =================================
              Row(
                children: [
                  Card(child: Text("S : ")),
                  Expanded(
                    child: Material(
                      elevation: 5,
                      child: TextFormField(
                        controller: sOrderQtyController,
                        decoration: InputDecoration(
                          filled: true,
                          fillColor: Colors.grey.withOpacity(0.5),
                          labelText: "S Order Qty",
                          contentPadding: EdgeInsets.only(
                            left: 20,
                          ),
                          border: InputBorder.none,
                        ),
                      ),
                    ),
                  ),
                  SizedBox(
                    width: 10,
                  ),
                  Expanded(
                    child: Material(
                      elevation: 5,
                      child: TextFormField(
                        controller: sIssueQtyController,
                        decoration: InputDecoration(
                          filled: true,
                          fillColor: Colors.grey.withOpacity(0.5),
                          labelText: "S Issue Qty",
                          contentPadding: EdgeInsets.only(
                            left: 20,
                          ),
                          border: InputBorder.none,
                        ),
                      ),
                    ),
                  ),
                  SizedBox(
                    width: 10,
                  ),
                  Expanded(
                    child: Material(
                      elevation: 5,
                      child: TextFormField(
                        controller: sRcvQtyController,
                        decoration: InputDecoration(
                          filled: true,
                          fillColor: Colors.grey.withOpacity(0.5),
                          labelText: "Q Rcv Qty",
                          contentPadding: EdgeInsets.only(
                            left: 20,
                          ),
                          border: InputBorder.none,
                        ),
                      ),
                    ),
                  ),
                  SizedBox(
                    width: 10,
                  ),
                  Expanded(
                    child: Material(
                      elevation: 5,
                      child: TextFormField(
                        controller: sCumQtyController,
                        decoration: InputDecoration(
                          filled: true,
                          fillColor: Colors.grey.withOpacity(0.5),
                          labelText: "S GR Cum Qty",
                          contentPadding: EdgeInsets.only(
                            left: 20,
                          ),
                          border: InputBorder.none,
                        ),
                      ),
                    ),
                  ),
                ],
              ),
              SizedBox(
                height: 15,
              ),

              // ============================  M =================================
              Row(
                children: [
                  Card(child: Text("M : ")),
                  Expanded(
                    child: Material(
                      elevation: 5,
                      child: TextFormField(
                        controller: mOrderQtyController,
                        decoration: InputDecoration(
                          filled: true,
                          fillColor: Colors.grey.withOpacity(0.5),
                          labelText: "M Order Qty",
                          contentPadding: EdgeInsets.only(
                            left: 20,
                          ),
                          border: InputBorder.none,
                        ),
                      ),
                    ),
                  ),
                  SizedBox(
                    width: 10,
                  ),
                  Expanded(
                    child: Material(
                      elevation: 5,
                      child: TextFormField(
                        controller: mIssueQtyController,
                        decoration: InputDecoration(
                          filled: true,
                          fillColor: Colors.grey.withOpacity(0.5),
                          labelText: "M Issue Qty",
                          contentPadding: EdgeInsets.only(
                            left: 20,
                          ),
                          border: InputBorder.none,
                        ),
                      ),
                    ),
                  ),
                  SizedBox(
                    width: 10,
                  ),
                  Expanded(
                    child: Material(
                      elevation: 5,
                      child: TextFormField(
                        controller: mRcvQtyController,
                        decoration: InputDecoration(
                          filled: true,
                          fillColor: Colors.grey.withOpacity(0.5),
                          labelText: "M Rcv Qty",
                          contentPadding: EdgeInsets.only(
                            left: 20,
                          ),
                          border: InputBorder.none,
                        ),
                      ),
                    ),
                  ),
                  SizedBox(
                    width: 10,
                  ),
                  Expanded(
                    child: Material(
                      elevation: 5,
                      child: TextFormField(
                        controller: mCumQtyController,
                        decoration: InputDecoration(
                          filled: true,
                          fillColor: Colors.grey.withOpacity(0.5),
                          labelText: "M GR Cum Qty",
                          contentPadding: EdgeInsets.only(
                            left: 20,
                          ),
                          border: InputBorder.none,
                        ),
                      ),
                    ),
                  ),
                ],
              ),
              SizedBox(
                height: 15,
              ),

              // ============================  L =================================
              Row(
                children: [
                  Card(child: Text("L : ")),
                  Expanded(
                    child: Material(
                      elevation: 5,
                      child: TextFormField(
                        controller: lOrderQtyController,
                        decoration: InputDecoration(
                          filled: true,
                          fillColor: Colors.grey.withOpacity(0.5),
                          labelText: "L Order Qty",
                          contentPadding: EdgeInsets.only(
                            left: 20,
                          ),
                          border: InputBorder.none,
                        ),
                      ),
                    ),
                  ),
                  SizedBox(
                    width: 10,
                  ),
                  Expanded(
                    child: Material(
                      elevation: 5,
                      child: TextFormField(
                        controller: lIssueQtyController,
                        decoration: InputDecoration(
                          filled: true,
                          fillColor: Colors.grey.withOpacity(0.5),
                          labelText: "L Issue Qty",
                          contentPadding: EdgeInsets.only(
                            left: 20,
                          ),
                          border: InputBorder.none,
                        ),
                      ),
                    ),
                  ),
                  SizedBox(
                    width: 10,
                  ),
                  Expanded(
                    child: Material(
                      elevation: 5,
                      child: TextFormField(
                        controller: lRcvQtyController,
                        decoration: InputDecoration(
                          filled: true,
                          fillColor: Colors.grey.withOpacity(0.5),
                          labelText: "L Rcv Qty",
                          contentPadding: EdgeInsets.only(
                            left: 20,
                          ),
                          border: InputBorder.none,
                        ),
                      ),
                    ),
                  ),
                  SizedBox(
                    width: 10,
                  ),
                  Expanded(
                    child: Material(
                      elevation: 5,
                      child: TextFormField(
                        controller: lCumQtyController,
                        decoration: InputDecoration(
                          filled: true,
                          fillColor: Colors.grey.withOpacity(0.5),
                          labelText: "L GR Cum Qty",
                          contentPadding: EdgeInsets.only(
                            left: 20,
                          ),
                          border: InputBorder.none,
                        ),
                      ),
                    ),
                  ),
                ],
              ),
              SizedBox(
                height: 15,
              ),

              // ============================  XL =================================
              Row(
                children: [
                  Card(child: Text("XL : ")),
                  Expanded(
                    child: Material(
                      elevation: 5,
                      child: TextFormField(
                        controller: xlOrderQtyController,
                        decoration: InputDecoration(
                          filled: true,
                          fillColor: Colors.grey.withOpacity(0.5),
                          labelText: "XL Order Qty",
                          contentPadding: EdgeInsets.only(
                            left: 20,
                          ),
                          border: InputBorder.none,
                        ),
                      ),
                    ),
                  ),
                  SizedBox(
                    width: 10,
                  ),
                  Expanded(
                    child: Material(
                      elevation: 5,
                      child: TextFormField(
                        controller: xlIssueQtyController,
                        decoration: InputDecoration(
                          filled: true,
                          fillColor: Colors.grey.withOpacity(0.5),
                          labelText: "XL Issue Qty",
                          contentPadding: EdgeInsets.only(
                            left: 20,
                          ),
                          border: InputBorder.none,
                        ),
                      ),
                    ),
                  ),
                  SizedBox(
                    width: 10,
                  ),
                  Expanded(
                    child: Material(
                      elevation: 5,
                      child: TextFormField(
                        controller: xlRcvQtyController,
                        decoration: InputDecoration(
                          filled: true,
                          fillColor: Colors.grey.withOpacity(0.5),
                          labelText: "XL Rcv Qty",
                          contentPadding: EdgeInsets.only(
                            left: 20,
                          ),
                          border: InputBorder.none,
                        ),
                      ),
                    ),
                  ),
                  SizedBox(
                    width: 10,
                  ),
                  Expanded(
                    child: Material(
                      elevation: 5,
                      child: TextFormField(
                        controller: xlCumQtyController,
                        decoration: InputDecoration(
                          filled: true,
                          fillColor: Colors.grey.withOpacity(0.5),
                          labelText: "XL GR Cum Qty",
                          contentPadding: EdgeInsets.only(
                            left: 20,
                          ),
                          border: InputBorder.none,
                        ),
                      ),
                    ),
                  ),
                ],
              ),
              SizedBox(
                height: 15,
              ),

              // ============================  XXL =================================
              Row(
                children: [
                  Card(child: Text("XXL : ")),
                  Expanded(
                    child: Material(
                      elevation: 5,
                      child: TextFormField(
                        controller: xxlOrderQtyController,
                        decoration: InputDecoration(
                          filled: true,
                          fillColor: Colors.grey.withOpacity(0.5),
                          labelText: "XXL Order Qty",
                          contentPadding: EdgeInsets.only(
                            left: 20,
                          ),
                          border: InputBorder.none,
                        ),
                      ),
                    ),
                  ),
                  SizedBox(
                    width: 10,
                  ),
                  Expanded(
                    child: Material(
                      elevation: 5,
                      child: TextFormField(
                        controller: xxlIssueQtyController,
                        decoration: InputDecoration(
                          filled: true,
                          fillColor: Colors.grey.withOpacity(0.5),
                          labelText: "XXL Issue Qty",
                          contentPadding: EdgeInsets.only(
                            left: 20,
                          ),
                          border: InputBorder.none,
                        ),
                      ),
                    ),
                  ),
                  SizedBox(
                    width: 10,
                  ),
                  Expanded(
                    child: Material(
                      elevation: 5,
                      child: TextFormField(
                        controller: xxlRcvQtyController,
                        decoration: InputDecoration(
                          filled: true,
                          fillColor: Colors.grey.withOpacity(0.5),
                          labelText: "XXL Rcv Qty",
                          contentPadding: EdgeInsets.only(
                            left: 20,
                          ),
                          border: InputBorder.none,
                        ),
                      ),
                    ),
                  ),
                  SizedBox(
                    width: 10,
                  ),
                  Expanded(
                    child: Material(
                      elevation: 5,
                      child: TextFormField(
                        controller: xxlCumQtyController,
                        decoration: InputDecoration(
                          filled: true,
                          fillColor: Colors.grey.withOpacity(0.5),
                          labelText: "XXL GR Cum Qty",
                          contentPadding: EdgeInsets.only(
                            left: 20,
                          ),
                          border: InputBorder.none,
                        ),
                      ),
                    ),
                  ),
                ],
              ),
              SizedBox(
                height: 15,
              ),

              // ============================  XXXL =================================
              Row(
                children: [
                  Card(child: Text("XXXL : ")),
                  Expanded(
                    child: Material(
                      elevation: 5,
                      child: TextFormField(
                        controller: xxxlOrderQtyController,
                        decoration: InputDecoration(
                          filled: true,
                          fillColor: Colors.grey.withOpacity(0.5),
                          labelText: "XXXL Order Qty",
                          contentPadding: EdgeInsets.only(
                            left: 20,
                          ),
                          border: InputBorder.none,
                        ),
                      ),
                    ),
                  ),
                  SizedBox(
                    width: 10,
                  ),
                  Expanded(
                    child: Material(
                      elevation: 5,
                      child: TextFormField(
                        controller: xxxlIssueQtyController,
                        decoration: InputDecoration(
                          filled: true,
                          fillColor: Colors.grey.withOpacity(0.5),
                          labelText: "XXXL Issue Qty",
                          contentPadding: EdgeInsets.only(
                            left: 20,
                          ),
                          border: InputBorder.none,
                        ),
                      ),
                    ),
                  ),
                  SizedBox(
                    width: 10,
                  ),
                  Expanded(
                    child: Material(
                      elevation: 5,
                      child: TextFormField(
                        controller: xxxlRcvQtyController,
                        decoration: InputDecoration(
                          filled: true,
                          fillColor: Colors.grey.withOpacity(0.5),
                          labelText: "XXXL Rcv Qty",
                          contentPadding: EdgeInsets.only(
                            left: 20,
                          ),
                          border: InputBorder.none,
                        ),
                      ),
                    ),
                  ),
                  SizedBox(
                    width: 10,
                  ),
                  Expanded(
                    child: Material(
                      elevation: 5,
                      child: TextFormField(
                        controller: xxxlCumQtyController,
                        decoration: InputDecoration(
                          filled: true,
                          fillColor: Colors.grey.withOpacity(0.5),
                          labelText: "XXXL GR Cum Qty",
                          contentPadding: EdgeInsets.only(
                            left: 20,
                          ),
                          border: InputBorder.none,
                        ),
                      ),
                    ),
                  ),
                ],
              ),
              SizedBox(
                height: 15,
              ),
              ElevatedButton(
                style: ElevatedButton.styleFrom(
                  backgroundColor: Styles().palletes4,
                ),
                onPressed: () async {
                  int totalIssueQty = (xxsIssueQtyController.text.isEmpty
                          ? 0
                          : int.parse(xxsIssueQtyController.text)) +
                      (xsIssueQtyController.text.isEmpty
                          ? 0
                          : int.parse(xsIssueQtyController.text)) +
                      (sIssueQtyController.text.isEmpty
                          ? 0
                          : int.parse(sIssueQtyController.text)) +
                      (mIssueQtyController.text.isEmpty
                          ? 0
                          : int.parse(mIssueQtyController.text)) +
                      (lIssueQtyController.text.isEmpty
                          ? 0
                          : int.parse(lIssueQtyController.text)) +
                      (xlIssueQtyController.text.isEmpty
                          ? 0
                          : int.parse(xlIssueQtyController.text)) +
                      (xxlIssueQtyController.text.isEmpty
                          ? 0
                          : int.parse(xxlIssueQtyController.text)) +
                      (xxxlIssueQtyController.text.isEmpty
                          ? 0
                          : int.parse(xxxlIssueQtyController.text));

                  int totalOrderQty = (xxsOrderQtyController.text.isEmpty
                          ? 0
                          : int.parse(xxsOrderQtyController.text)) +
                      (xsOrderQtyController.text.isEmpty
                          ? 0
                          : int.parse(xsOrderQtyController.text)) +
                      (sOrderQtyController.text.isEmpty
                          ? 0
                          : int.parse(sOrderQtyController.text)) +
                      (mOrderQtyController.text.isEmpty
                          ? 0
                          : int.parse(mOrderQtyController.text)) +
                      (lOrderQtyController.text.isEmpty
                          ? 0
                          : int.parse(lOrderQtyController.text)) +
                      (xlOrderQtyController.text.isEmpty
                          ? 0
                          : int.parse(xlOrderQtyController.text)) +
                      (xxlOrderQtyController.text.isEmpty
                          ? 0
                          : int.parse(xxlOrderQtyController.text)) +
                      (xxxlOrderQtyController.text.isEmpty
                          ? 0
                          : int.parse(xxxlOrderQtyController.text));

                  int totalRcvQty = (xxsRcvQtyController.text.isEmpty
                          ? 0
                          : int.parse(xxsRcvQtyController.text)) +
                      (xsRcvQtyController.text.isEmpty
                          ? 0
                          : int.parse(xsRcvQtyController.text)) +
                      (sRcvQtyController.text.isEmpty
                          ? 0
                          : int.parse(sRcvQtyController.text)) +
                      (mRcvQtyController.text.isEmpty
                          ? 0
                          : int.parse(mRcvQtyController.text)) +
                      (lRcvQtyController.text.isEmpty
                          ? 0
                          : int.parse(lRcvQtyController.text)) +
                      (xlRcvQtyController.text.isEmpty
                          ? 0
                          : int.parse(xlRcvQtyController.text)) +
                      (xxlRcvQtyController.text.isEmpty
                          ? 0
                          : int.parse(xxlRcvQtyController.text)) +
                      (xxxlRcvQtyController.text.isEmpty
                          ? 0
                          : int.parse(xxxlRcvQtyController.text));

                  int totalCumQty = (xxsCumQtyController.text.isEmpty
                          ? 0
                          : int.parse(xxsCumQtyController.text)) +
                      (xsCumQtyController.text.isEmpty
                          ? 0
                          : int.parse(xsCumQtyController.text)) +
                      (sCumQtyController.text.isEmpty
                          ? 0
                          : int.parse(sCumQtyController.text)) +
                      (mCumQtyController.text.isEmpty
                          ? 0
                          : int.parse(mCumQtyController.text)) +
                      (lCumQtyController.text.isEmpty
                          ? 0
                          : int.parse(lCumQtyController.text)) +
                      (xlCumQtyController.text.isEmpty
                          ? 0
                          : int.parse(xlCumQtyController.text)) +
                      (xxlCumQtyController.text.isEmpty
                          ? 0
                          : int.parse(xxlCumQtyController.text)) +
                      (xxxlCumQtyController.text.isEmpty
                          ? 0
                          : int.parse(xxxlCumQtyController.text));

                  try {
                    DataModel dataModel = DataModel(
                      factoryName: factoryController.text.trim(),
                      floorName: floorController.text.trim(),
                      lineName: lineController.text.trim(),
                      buyer: buyerController.text.trim(),
                      style: styleController.text.trim(),
                      buy: buyController.text.trim(),
                      salesOrderNo: salesOrderController.text.trim(),
                      item: itemController.text.trim(),
                      fGMatNo: fgMatController.text.trim(),
                      ourNoFR: ourNoController.text.trim(),
                      colorName: colorController.text.trim(),
                      countryName: countryController.text.trim(),
                      productionOrderNo: prodOrderController.text.trim(),
                      orderQty: totalOrderQty,
                      issueQty: totalIssueQty,
                      xXS: XXS(
                          orderQty: xxsOrderQtyController.text.isEmpty
                              ? 0
                              : int.parse(xxsOrderQtyController.text),
                          issue: xxsIssueQtyController.text.isEmpty
                              ? 0
                              : int.parse(xxsIssueQtyController.text),
                          receivedQty: xxsRcvQtyController.text.isEmpty
                              ? 0
                              : int.parse(xxsRcvQtyController.text),
                          totalCum: xxsCumQtyController.text.isEmpty
                              ? 0
                              : int.parse(xxsCumQtyController.text)),
                      xS: XXS(
                          orderQty: xsOrderQtyController.text.isEmpty
                              ? 0
                              : int.parse(xsOrderQtyController.text),
                          issue: xsIssueQtyController.text.isEmpty
                              ? 0
                              : int.parse(xsIssueQtyController.text),
                          receivedQty: xsRcvQtyController.text.isEmpty
                              ? 0
                              : int.parse(xsRcvQtyController.text),
                          totalCum: xsCumQtyController.text.isEmpty
                              ? 0
                              : int.parse(xsCumQtyController.text)),
                      s: XXS(
                          orderQty: sOrderQtyController.text.isEmpty
                              ? 0
                              : int.parse(sOrderQtyController.text),
                          issue: sIssueQtyController.text.isEmpty
                              ? 0
                              : int.parse(sIssueQtyController.text),
                          receivedQty: sRcvQtyController.text.isEmpty
                              ? 0
                              : int.parse(sRcvQtyController.text),
                          totalCum: sCumQtyController.text.isEmpty
                              ? 0
                              : int.parse(sCumQtyController.text)),
                      m: XXS(
                          orderQty: mOrderQtyController.text.isEmpty
                              ? 0
                              : int.parse(mOrderQtyController.text),
                          issue: mIssueQtyController.text.isEmpty
                              ? 0
                              : int.parse(mIssueQtyController.text),
                          receivedQty: mRcvQtyController.text.isEmpty
                              ? 0
                              : int.parse(mRcvQtyController.text),
                          totalCum: mCumQtyController.text.isEmpty
                              ? 0
                              : int.parse(mCumQtyController.text)),
                      l: XXS(
                          orderQty: lOrderQtyController.text.isEmpty
                              ? 0
                              : int.parse(lOrderQtyController.text),
                          issue: lIssueQtyController.text.isEmpty
                              ? 0
                              : int.parse(lIssueQtyController.text),
                          receivedQty: lRcvQtyController.text.isEmpty
                              ? 0
                              : int.parse(lRcvQtyController.text),
                          totalCum: lCumQtyController.text.isEmpty
                              ? 0
                              : int.parse(lCumQtyController.text)),
                      xL: XXS(
                          orderQty: xlOrderQtyController.text.isEmpty
                              ? 0
                              : int.parse(xlOrderQtyController.text),
                          issue: xlIssueQtyController.text.isEmpty
                              ? 0
                              : int.parse(xlIssueQtyController.text),
                          receivedQty: xlRcvQtyController.text.isEmpty
                              ? 0
                              : int.parse(xlRcvQtyController.text),
                          totalCum: xlCumQtyController.text.isEmpty
                              ? 0
                              : int.parse(xlCumQtyController.text)),
                      xXL: XXS(
                          orderQty: xxlOrderQtyController.text.isEmpty
                              ? 0
                              : int.parse(xxlOrderQtyController.text),
                          issue: xxlIssueQtyController.text.isEmpty
                              ? 0
                              : int.parse(xxlIssueQtyController.text),
                          receivedQty: xxlRcvQtyController.text.isEmpty
                              ? 0
                              : int.parse(xxlRcvQtyController.text),
                          totalCum: xxlCumQtyController.text.isEmpty
                              ? 0
                              : int.parse(xxlCumQtyController.text)),
                      xXXL: XXS(
                          orderQty: xxxlOrderQtyController.text.isEmpty
                              ? 0
                              : int.parse(xxxlOrderQtyController.text),
                          issue: xxxlIssueQtyController.text.isEmpty
                              ? 0
                              : int.parse(xxxlIssueQtyController.text),
                          receivedQty: xxxlRcvQtyController.text.isEmpty
                              ? 0
                              : int.parse(xxxlRcvQtyController.text),
                          totalCum: xxxlCumQtyController.text.isEmpty
                              ? 0
                              : int.parse(xxxlCumQtyController.text)),
                      dGrQty: totalRcvQty,
                      totalGrCumQty:
                          totalRcvQty + int.parse(lastDayPackController.text),
                      balanceGrQty: totalIssueQty -
                          (totalRcvQty + int.parse(lastDayPackController.text)),
                      dailyPackQty: totalRcvQty,
                      lastDayPackCumQty: int.parse(lastDayPackController.text),
                      balanceQty: totalIssueQty -
                          (totalRcvQty + int.parse(lastDayPackController.text)),
                      remarks: "",
                      date: setDate,
                      dateOrder: DateTime.now().millisecondsSinceEpoch,
                    );

                    ScaffoldMessenger.of(context).showSnackBar(
                      SnackBar(
                        backgroundColor: Colors.green,
                        content: Text(
                          "Saving Data....",
                          style: TextStyle(color: Colors.white),
                        ),
                      ),
                    );

                    await FirebaseFirestore.instance
                        .collection(factoryController.text)
                        .doc(prodOrderController.text)
                        .set(
                          dataModel.toJson(),
                        );

                    ScaffoldMessenger.of(context).showSnackBar(
                      SnackBar(
                        backgroundColor: Colors.green,
                        content: Text(
                          "Successfully Saved",
                          style: TextStyle(color: Colors.white),
                        ),
                      ),
                    );
                  } catch (err) {
                    print(err.toString());
                    ScaffoldMessenger.of(context).showSnackBar(
                      SnackBar(
                        backgroundColor: Colors.red,
                        content: Text(
                          err.toString(),
                          style: TextStyle(color: Colors.white),
                        ),
                      ),
                    );
                  }
                },
                child: Text(
                  "Submit",
                  style: TextStyle(
                    color: Colors.white,
                  ),
                ),
              )
            ],
          ),
        ),
      ),
    );
  }
}
