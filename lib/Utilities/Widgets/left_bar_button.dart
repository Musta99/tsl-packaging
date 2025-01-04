import 'package:flutter/foundation.dart';
import 'package:flutter/material.dart';
import 'package:font_awesome_flutter/font_awesome_flutter.dart';
import 'package:google_fonts/google_fonts.dart';
import 'package:tsl_packaging/Utilities/styles.dart';

class LeftBarCommonButton extends StatelessWidget {
  String title;
  IconData icon;
  double textOpacity;
  double iconOpacity;
  LeftBarCommonButton({
    super.key,
    required this.icon,
    required this.title,
    required this.iconOpacity,
    required this.textOpacity,
  });

  @override
  Widget build(BuildContext context) {
    return Container(
      width: MediaQuery.of(context).size.width * 0.2 * 0.8,
      height: MediaQuery.of(context).size.height * 0.08,
      decoration: BoxDecoration(
        color: Styles().palletes1,
        borderRadius: BorderRadius.only(
          topRight: Radius.circular(5),
          bottomRight: Radius.circular(5),
        ),
      ),
      child: Padding(
        padding: EdgeInsets.only(
          left: MediaQuery.of(context).size.width * 0.015,
        ),
        child: Row(
          children: [
            Opacity(
              opacity: iconOpacity,
              child: Icon(
                icon,
                size: MediaQuery.of(context).size.width * 0.012,
                color: Styles().palletes4,
              ),
            ),
            SizedBox(
              width: MediaQuery.of(context).size.width * 0.01,
            ),
            Opacity(
              opacity: textOpacity,
              child: Text(
                title,
                style: GoogleFonts.ptSerif(
                  color: Styles().palletes4,
                  fontWeight: FontWeight.bold,
                  fontSize: MediaQuery.of(context).size.width * 0.01,
                ),
              ),
            ),
          ],
        ),
      ),
    );
  }
}
