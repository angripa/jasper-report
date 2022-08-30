package com.angripa.report.controller;


import com.angripa.report.domain.XlsData;
import com.angripa.report.util.CurrenyUtil;
import net.sf.jasperreports.engine.*;
import net.sf.jasperreports.engine.data.JRBeanCollectionDataSource;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.ApplicationContext;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.InputStream;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.*;


@Controller
public class FileUploadController {

   @Autowired
   ApplicationContext context;

   static final String SHEET = "Sheet1";
   static final SimpleDateFormat simpleDateFormat = new SimpleDateFormat("dd/MM/yy hh:mm:ss");

   @GetMapping("")
   public String hello() {
      return "uploader";
   }

   @PostMapping("/upload")
   public ResponseEntity<?> handleFileUpload(@RequestParam("file") final MultipartFile file, HttpServletResponse response) {
      String[] split = file.getOriginalFilename().split("\\.");
      String ext =  split[split.length-1] ;
      String fileName = file.getOriginalFilename().replace(ext,"pdf");
      System.out.println(ext);
      if(!"xls".equals(ext) && !"xlsx".equals(ext)){
         return ResponseEntity.ok("Invalid Extension");
      }
      try {
         Resource resource = context.getResource("classpath:reports/sample.jrxml");
         Resource resourceLogo = context.getResource("classpath:reports/logo.jpeg");

         //Compile to jasperReport
         InputStream inputStream = resource.getInputStream();
         JasperReport report = JasperCompileManager.compileReport(inputStream);
         //Parameters Set
         Map<String, Object> params = new HashMap<>();

         List<XlsData> dataList = new ArrayList<>();
         readData(file.getInputStream(), dataList, params);
         //XlsData source Set
         JRDataSource dataSource = new JRBeanCollectionDataSource(dataList);
         params.put("datasource", dataSource);
         params.put("logo", resourceLogo.getURI().getPath());

         //Make jasperPrint
         JasperPrint jasperPrint = JasperFillManager.fillReport(report, params, dataSource);
         //Media Type
         response.setContentType(MediaType.APPLICATION_PDF_VALUE);

         response.addHeader("Content-Disposition", "attachment; filename=" + fileName);
         //Export PDF Stream
         JasperExportManager.exportReportToPdfStream(jasperPrint, response.getOutputStream());
      } catch (Exception e) {
         e.printStackTrace();
         return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
      }
      return ResponseEntity.ok("File uploaded successfully.");
   }


   public static void readData(InputStream is, List<XlsData> dataList, Map<String, Object> headers) {
      try {
         Workbook workbook = new XSSFWorkbook(is);
         Sheet sheet = workbook.getSheet(SHEET);
         Iterator<Row> rows = sheet.iterator();
         int rowNumber = 0;
         while (rows.hasNext()) {
            Row currentRow = rows.next();
            Iterator<Cell> cellsInRow = currentRow.iterator();
            int cellIdx = 0;
            // skip header
            if (rowNumber < 8) {
               rowNumber++;
               if (rowNumber == 3) {
                  while (cellsInRow.hasNext()) {
                     Cell currentCell = cellsInRow.next();
                     switch (cellIdx) {
                        case 1:
                           headers.put("accountNo", currentCell.getStringCellValue());
                           break;
                        case 3:
                           headers.put("inBalance", currentCell.getNumericCellValue());
                           break;
                        default:
                           break;
                     }
                     cellIdx++;
                  }
               } else if (rowNumber == 4) {
                  while (cellsInRow.hasNext()) {
                     Cell currentCell = cellsInRow.next();
                     switch (cellIdx) {
                        case 1:
                           headers.put("currency", currentCell.getStringCellValue());
                           break;
                        case 3:
                           headers.put("debtBalance", currentCell.getNumericCellValue());
                           break;
                        default:
                           break;
                     }
                     cellIdx++;
                  }
               } else if (rowNumber == 5) {
                  while (cellsInRow.hasNext()) {
                     Cell currentCell = cellsInRow.next();
                     switch (cellIdx) {
                        case 1:
                           headers.put("trxPeriod", currentCell.getStringCellValue());
                           break;
                        case 3:
                           headers.put("creditBalance", currentCell.getNumericCellValue());
                           break;
                        default:
                           break;
                     }
                     cellIdx++;
                  }
               } else if (rowNumber == 6) {
                  while (cellsInRow.hasNext()) {
                     Cell currentCell = cellsInRow.next();
                     switch (cellIdx) {
                        case 1:
                           headers.put("reportDate", currentCell.getStringCellValue());
                           break;
                        case 3:
                           headers.put("lastBalance", currentCell.getNumericCellValue());
                           headers.put("lastBalanceStr", CurrenyUtil.parse(BigDecimal.valueOf(currentCell.getNumericCellValue())).toUpperCase());
                           break;
                        default:
                           break;
                     }
                     cellIdx++;
                  }
               }
               continue;
            }

            XlsData tutorial = new XlsData();


            while (cellsInRow.hasNext()) {
               Cell currentCell = cellsInRow.next();
               switch (cellIdx) {
                  case 0:
                     Date d = simpleDateFormat.parse(currentCell.getStringCellValue());
                     tutorial.setDate(simpleDateFormat.format(d));
                     break;
                  case 1:
                     tutorial.setDetailTrx(currentCell.getStringCellValue());
                     break;
                  case 2:
                     tutorial.setTeller(currentCell.getCellType().getCode() == 0 ? new BigDecimal(currentCell.getNumericCellValue()).setScale(0).toPlainString() : currentCell.getStringCellValue());
                     break;
                  case 3:
                     tutorial.setDebit(currentCell.getNumericCellValue());
                     break;
                  case 4:
                     tutorial.setCredit(currentCell.getNumericCellValue());
                     break;
                  case 5:
                     tutorial.setBalance(currentCell.getNumericCellValue());
                     break;
                  default:
                     break;
               }
               cellIdx++;
            }
            dataList.add(tutorial);
         }
         workbook.close();
      } catch (Exception e) {
         throw new RuntimeException("fail to parse Excel file: " + e.getMessage());
      }
   }

}