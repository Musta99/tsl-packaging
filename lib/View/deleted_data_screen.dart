import 'package:flutter/material.dart';
import 'package:flutter_easyloading/flutter_easyloading.dart';
import 'package:provider/provider.dart';
import 'package:tsl_packaging/Services/excel_generator.dart';
import 'package:tsl_packaging/State_Management/data_fetch_state_management.dart';
import 'package:tsl_packaging/Utilities/Widgets/datatable_widget.dart';

class DeletedDataScreen extends StatefulWidget {
  const DeletedDataScreen({super.key});

  @override
  State<DeletedDataScreen> createState() => _DeletedDataScreenState();
}

class _DeletedDataScreenState extends State<DeletedDataScreen> {
  ScrollController _scrollController = ScrollController();
  @override
  void initState() {
    // TODO: implement initState
    super.initState();

    Provider.of<DataFetchStateManagement>(context, listen: false)
        .fetchDeletedPO();
    Provider.of<DataFetchStateManagement>(context, listen: false)
        .fetchDeletedPOforExcel();

    _scrollController.addListener(() {
      if (_scrollController.position.pixels ==
          _scrollController.position.maxScrollExtent) {
        Provider.of<DataFetchStateManagement>(context, listen: false)
            .fetchNextBatch();
      }
    });
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      floatingActionButton: FloatingActionButton.small(
        onPressed: () {
          ExcelGenerator().createExcelForDeletedPO(
              Provider.of<DataFetchStateManagement>(context, listen: false)
                  .deletedPOExcel);
        },
        child: Center(
          child: Image.asset(
            "assets/images/excel.png",
            height: 30,
            width: 30,
          ),
        ),
      ),
      body: Container(
        margin: EdgeInsets.symmetric(
          horizontal: 20,
          vertical: 20,
        ),
        child: Column(
          children: [
            SelectableText("Deleted PO"),
            Expanded(
              child: Provider.of<DataFetchStateManagement>(context,
                              listen: false)
                          .deletedPOList
                          .length ==
                      0
                  ? Center(
                      child: Text("No Data found"),
                    )
                  : Provider.of<DataFetchStateManagement>(context).isLoading ==
                          true
                      ? Center(
                          child: CircularProgressIndicator(),
                        )
                      : GridView.builder(
                          controller: _scrollController,
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
                            if (index ==
                                Provider.of<DataFetchStateManagement>(context)
                                    .deletedPOList
                                    .length) {
                              if (Provider.of<DataFetchStateManagement>(context)
                                  .isLoading) {
                                return CircularProgressIndicator();
                              } else if (!Provider.of<DataFetchStateManagement>(
                                      context)
                                  .hasMore) {
                                return Text("No More Data to Show!");
                              } else {
                                return SizedBox.shrink();
                              }
                            } else {
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
                                              child: Text(
                                                  Provider.of<DataFetchStateManagement>(
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
                                            Text(Provider.of<
                                                        DataFetchStateManagement>(
                                                    context,
                                                    listen: false)
                                                .deletedPOList[index]["floorName"])
                                          ],
                                        ),
                                        Row(
                                          children: [
                                            Text("Concerned Person: "),
                                            Text(Provider.of<
                                                        DataFetchStateManagement>(
                                                    context,
                                                    listen: false)
                                                .deletedPOList[index]["userName"])
                                          ],
                                        ),
                                      ],
                                    ),
                                  ),
                                ),
                              );
                            }
                          }),
            )
          ],
        ),
      ),
    );
  }
}
