package com.angripa.report.util;

import java.math.BigDecimal;
import java.math.RoundingMode;

public class CurrenyUtil {
   public static String parse(BigDecimal value) {
      value = value.setScale(2, RoundingMode.UP);
      String strValue = value.toString();

      String commaDecimal = "";
      if (strValue.contains(".")) {
         commaDecimal = strValue.substring(strValue.indexOf(".") + 1);
         strValue = strValue.substring(0, strValue.indexOf("."));
      }
      if (!commaDecimal.isEmpty()) {
         return terbilang(new BigDecimal(strValue)).concat("Koma ").concat(terbilang(new BigDecimal(commaDecimal)));
      } else {
         return terbilang(new BigDecimal(strValue));
      }
   }

   public static String terbilang(BigDecimal value) {
      String strValue = value.toString();

      int lenValue = strValue.length();
      int x = 0;
      int y = 0;
      int z;
      String bil1 = "";
      String bil2 = "";
      String urai = "";
      while (x < lenValue) {
         x = x + 1;
         int strTot = Integer.valueOf(strValue.substring(x - 1, x));
         y = y + strTot;
         z = lenValue - x + 1;
         switch (strTot) {
            case 1:
               if (z == 1 || z == 7 || z == 10 || z == 13) {
                  bil1 = "Satu ";
               } else if (z == 4) {
                  if (x == 1) {
                     bil1 = "Se";
                  } else {
                     bil1 = "Satu ";
                  }
               } else if (z == 2 || z == 5 || z == 8 || z == 11 || z == 14) {
                  x = x + 1;
                  int newStrTot = Integer.valueOf(strValue.substring(x - 1, x));
                  z = lenValue - x + 1;
                  bil2 = "";
                  switch (newStrTot) {
                     case 0:
                        bil1 = "Sepuluh ";
                        break;
                     case 1:
                        bil1 = "Sebelas ";
                        break;
                     case 2:
                        bil1 = "Dua belas ";
                        break;
                     case 3:
                        bil1 = "Tiga belas ";
                        break;
                     case 4:
                        bil1 = "Empat belas ";
                        break;
                     case 5:
                        bil1 = "Lima belas ";
                        break;
                     case 6:
                        bil1 = "Enam belas ";
                        break;
                     case 7:
                        bil1 = "Tujuh belas ";
                        break;
                     case 8:
                        bil1 = "Delapan belas ";
                        break;
                     case 9:
                        bil1 = "Sembilan belas ";
                        break;
                     default:
                        break;
                  }
               } else {
                  bil1 = "Se";
               }
               break;
            case 2:
               bil1 = "Dua ";
               break;
            case 3:
               bil1 = "Tiga ";
               break;
            case 4:
               bil1 = "Empat ";
               break;
            case 5:
               bil1 = "Lima ";
               break;
            case 6:
               bil1 = "Enam ";
               break;
            case 7:
               bil1 = "Tujuh ";
               break;
            case 8:
               bil1 = "Delapan ";
               break;
            case 9:
               bil1 = "Sembilan ";
               break;
            default:
               bil1 = "";
               break;
         }

         if (strTot > 0) {
            if (z == 2 || z == 5 || z == 8 || z == 11 || z == 14) {
               bil2 = "Puluh ";
            } else if (z == 3 || z == 6 || z == 9 || z == 12 || z == 15) {
               bil2 = "Ratus ";
            } else {
               bil2 = "";
            }
         } else {
            bil2 = "";
         }

         if (y > 0) {
            switch (z) {
               case 4:
                  bil2 = bil2 + "Ribu ";
                  y = 0;
                  break;
               case 7:
                  bil2 = bil2 + "Juta ";
                  y = 0;
                  break;
               case 10:
                  bil2 = bil2 + "Milyar ";
                  y = 0;
                  break;
               case 13:
                  bil2 = bil2 + "Trilyun ";
                  y = 0;
                  break;
               default:
                  break;
            }
         }

         if (bil1.equals("Se")) {
            String pre = bil2.substring(0, 1);
            urai = urai + bil1 + bil2.replace(pre, pre.toLowerCase());
         } else {
            urai = urai + bil1 + bil2;
         }
      }

      return urai;
   }
}
