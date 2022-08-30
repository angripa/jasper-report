package com.angripa.report.controller;

import com.angripa.report.domain.XlsData;
import com.angripa.report.util.CurrenyUtil;
import net.sf.jasperreports.engine.*;
import net.sf.jasperreports.engine.data.JRBeanCollectionDataSource;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.ApplicationContext;
import org.springframework.core.io.Resource;
import org.springframework.http.MediaType;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.InputStream;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * Created by IntelliJ IDEA.
 * Project : spring-boot-mysql-report
 * User: hendisantika
 * Email: hendisantika@gmail.com
 * Telegram : @hendisantika34
 * Date: 25/02/18
 * Time: 19.17
 * To change this template use File | Settings | File Templates.
 */

@RestController
@RequestMapping("/")
public class PdfController {

//    private Logger logger = LogManager.getLogManager(PdfController.class);

   @Autowired
   ApplicationContext context;

   static final String SHEET = "Sheet1";
   static final SimpleDateFormat simpleDateFormat = new SimpleDateFormat("dd/MM/yy hh:mm:ss");

//    @Autowired
//    CarRepository carRepository;

   //    @GetMapping(path = "pdf/{jrxml}")
   @PostMapping(path = "/pdf")
   @ResponseBody
//    public void getPdf(@PathVariable String jrxml, HttpServletResponse response) throws Exception {
   public void convert(@RequestParam("file") final MultipartFile file, HttpServletResponse response) throws Exception {

      String[] split = file.getOriginalFilename().split("\\.");
      String ext =  split[split.length-1] ;
      String fileName = file.getOriginalFilename().replace(ext,"pdf");
      if(!"xls".equals(ext) && !"xlsx".equals(ext)){
         return;
      }
      //Get JRXML template from resources folder
//        Resource resource = context.getResource("classpath:reports/" + jrxml + ".jrxml");
      Resource resource = context.getResource("classpath:reports/sample.jrxml");
      Resource resourceLogo = context.getResource("classpath:reports/logo.png");

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
