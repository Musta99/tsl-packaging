import 'package:firebase_core/firebase_core.dart';
import 'package:flutter/material.dart';
import 'package:google_fonts/google_fonts.dart';
import 'package:provider/provider.dart';
import 'package:tsl_packaging/State_Management/data_fetch_state_management.dart';
import 'package:tsl_packaging/State_Management/left_bar_state_management.dart';
import 'package:tsl_packaging/State_Management/selection_state_management.dart';
import 'package:tsl_packaging/Utilities/styles.dart';
import 'package:tsl_packaging/View/data_details_view.dart';
import 'package:tsl_packaging/View/data_input_screen.dart';
import 'package:tsl_packaging/View/data_view_screen.dart';
import 'package:tsl_packaging/firebase_options.dart';

void main() async {
  WidgetsFlutterBinding.ensureInitialized();
  await Firebase.initializeApp(options: DefaultFirebaseOptions.currentPlatform);
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  // This widget is the root of your application.
  @override
  Widget build(BuildContext context) {
    return MultiProvider(
      providers: [
        ChangeNotifierProvider(create: (context) => LeftBarStateManagement()),
        ChangeNotifierProvider(create: (context) => SelectionStateManagement()),
        ChangeNotifierProvider(create: (context) => DataFetchStateManagement()),
      ],
      builder: (context, child) {
        return MaterialApp(
          debugShowCheckedModeBanner: false,
          title: 'EasyPack-Your packaging partner',
          theme: ThemeData(
            scaffoldBackgroundColor: Styles().backgroundColor,
            textTheme: GoogleFonts.ptSerifTextTheme(),
            colorScheme: ColorScheme.fromSeed(seedColor: Colors.deepPurple),
            useMaterial3: true,
          ),
          home: DataViewScreen(),
        );
      },
    );
  }
}
