import 'package:flutter/material.dart';
import 'package:flutter_easyloading/flutter_easyloading.dart';
import 'package:provider/provider.dart';
import 'package:tsl_packaging/State_Management/data_fetch_state_management.dart';
import 'package:tsl_packaging/Utilities/Widgets/datatable_widget.dart';

class DeletedDataScreen extends StatefulWidget {
  const DeletedDataScreen({super.key});

  @override
  State<DeletedDataScreen> createState() => _DeletedDataScreenState();
}

class _DeletedDataScreenState extends State<DeletedDataScreen> {
  @override
  void initState() {
    // TODO: implement initState
    super.initState();

    Provider.of<DataFetchStateManagement>(context, listen: false)
        .fetchDeletedPO();
  }

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
            SelectableText("Deleted PO"),
            Expanded(
              child:
                  Provider.of<DataFetchStateManagement>(context, listen: false)
                              .deletedPOList
                              .length ==
                          0
                      ? Center(
                          child: Text("No Data found"),
                        )
                      : Provider.of<DataFetchStateManagement>(context)
                                  .isLoading ==
                              true
                          ? Center(
                              child: CircularProgressIndicator(),
                            )
                          : GridView.builder(
                              itemCount: Provider.of<DataFetchStateManagement>(
                                      context,
                                      listen: false)
                                  .deletedPOList
                                  .length,
                              gridDelegate:
                                  SliverGridDelegateWithFixedCrossAxisCount(
                                      crossAxisCount: 4,
                                      crossAxisSpacing: 10,
                                      mainAxisSpacing: 10,
                                      mainAxisExtent: 180),
                              itemBuilder: (context, index) {
                                return Card(
                                  shape: RoundedRectangleBorder(
                                      borderRadius: BorderRadius.circular(5)),
                                  child: Container(
                                    child: Padding(
                                      padding: const EdgeInsets.all(12.0),
                                      child: Column(
                                        children: [
                                          Row(
                                            crossAxisAlignment:
                                                CrossAxisAlignment.start,
                                            children: [
                                              Text("Production Order No: "),
                                              Expanded(
                                                child: Text(Provider.of<
                                                                DataFetchStateManagement>(
                                                            context,
                                                            listen: false)
                                                        .deletedPOList[index]
                                                    ["productionOrderNo"]),
                                              )
                                            ],
                                          ),
                                          Row(
                                            children: [
                                              Text("Purchase Order No: "),
                                              Text(
                                                  Provider.of<DataFetchStateManagement>(
                                                              context,
                                                              listen: false)
                                                          .deletedPOList[index]
                                                      ["purchaseOrderNo"])
                                            ],
                                          ),
                                          Row(
                                            children: [
                                              Text("Style Name: "),
                                              Text(Provider.of<
                                                          DataFetchStateManagement>(
                                                      context,
                                                      listen: false)
                                                  .deletedPOList[index]["style"])
                                            ],
                                          ),
                                          Row(
                                            children: [
                                              Text("Factory Name: "),
                                              Text(
                                                  Provider.of<DataFetchStateManagement>(
                                                              context,
                                                              listen: false)
                                                          .deletedPOList[index]
                                                      ["factoryName"])
                                            ],
                                          ),
                                          Row(
                                            children: [
                                              Text("Floor Name: "),
                                              Text(
                                                  Provider.of<DataFetchStateManagement>(
                                                              context,
                                                              listen: false)
                                                          .deletedPOList[index]
                                                      ["floorName"])
                                            ],
                                          ),
                                          Row(
                                            children: [
                                              Text("Concerned Person: "),
                                              Text(
                                                  Provider.of<DataFetchStateManagement>(
                                                              context,
                                                              listen: false)
                                                          .deletedPOList[index]
                                                      ["deletedBy"])
                                            ],
                                          ),
                                        ],
                                      ),
                                    ),
                                  ),
                                );
                              }),
            )
          ],
        ),
      ),
    );
  }
}
