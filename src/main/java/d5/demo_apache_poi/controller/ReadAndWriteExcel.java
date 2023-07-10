package d5.demo_apache_poi.controller;


import d5.demo_apache_poi.model.Sales;
import d5.demo_apache_poi.service.SaleService;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@RestController
@RequestMapping("/api")
@RequiredArgsConstructor
public class ReadAndWriteExcel {
//    public static void main(String[] args) {
//        readFromFileExcel();
//    }

    private final SaleService saleService;

    @PostMapping("/readFile")
    public String readFile() {
        String pathFile = "D:\\thuc_tap\\demo_Apache_Poi\\BaoCao.xlsx";
        try (FileInputStream fis = new FileInputStream(pathFile);
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            List<Sales> dataSales = new ArrayList<>();
            // Duyệt qua từng dòng trong sheet
            DataFormatter formatter = new DataFormatter();
//             for (Row row : sheet) {
//                 row= sheet.getRow(5);
            for (int i = 4; i < sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);

                // Đọc dữ liệu từ các ô trong dòng
                Cell cell1 = row.getCell(0); // Ô đầu tiên
                Cell cell2 = row.getCell(1); // Ô thứ hai
                Cell cell3 = row.getCell(2); // Ô thứ ba
                Cell cell4 = row.getCell(3); // Ô thứ ba
                Cell cell5 = row.getCell(4); // Ô thứ ba

                // Lấy giá trị từ ô

                String nganhHang = formatter.formatCellValue(cell1);
                if (nganhHang.isEmpty()) {
                    break;
                }
                int quanAo = Integer.parseInt(formatter.formatCellValue(cell2));
                int giayDep = Integer.parseInt(formatter.formatCellValue(cell3));
                int tuiSach = Integer.parseInt(formatter.formatCellValue(cell4));
                int muNon = Integer.parseInt(formatter.formatCellValue(cell5));
                // Tạo một đối tượng ExcelData từ dữ liệu đọc được
                Sales excelData = new Sales(nganhHang, quanAo, giayDep, tuiSach, muNon);

                dataSales.add(excelData);
            }
            // Lưu danh sách ExcelData vào cơ sở dữ liệu
            for (Sales s : dataSales) {
                saleService.save(s);
            }

            return "Dữ liệu đã được lưu vào cơ sở dữ liệu thành công!";
        } catch (IOException e) {
            e.printStackTrace();
            return "Lỗi khi đọc file Excel: " + e.getMessage();
        }
    }

    @PostMapping("/writeFile")
    public String writeFile() {
        try {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet("DanhSach");

            Row row = null;
            Cell cell = null;
            row = sheet.createRow(0);
            cell = row.createCell(1);
            cell.setCellValue("Bảng doanh số bán hàng Quý 1 - năm 2023");

            row = sheet.createRow(1);
            cell = row.createCell(2);
            cell.setCellValue("(Đơn vị: triệu đồng)");

            row = sheet.createRow(3);
            cell = row.createCell(0);
            cell.setCellValue("Ngành hàng");
            cell = row.createCell(1);
            cell.setCellValue("Quần Áo");
            cell = row.createCell(2);
            cell.setCellValue("Giày Dép");
            cell = row.createCell(3);
            cell.setCellValue("Túi Xách");
            cell = row.createCell(4);
            cell.setCellValue("Mũ Nón");

            List<Sales> list = saleService.findAll();
            for (int i = 0; i < list.size() ; i++) {
                row = sheet.createRow(4+i);
                cell=row.createCell(0);
                cell.setCellValue(list.get(i).getNganhHang());
                cell=row.createCell(1);
                cell.setCellValue(list.get(i).getQuanAo());
                cell=row.createCell(2);
                cell.setCellValue(list.get(i).getGiayDep());
                cell=row.createCell(3);
                cell.setCellValue(list.get(i).getTuiSach());
                cell=row.createCell(4);
                cell.setCellValue(list.get(i).getMuNon());
            }
            String pathFile = "D:\\TH_BT_MD5_SpringBoot\\Demo-ApachePOI-Excel.xlsx";
            FileOutputStream fileOutputStream= new FileOutputStream(pathFile);
            workbook.write(fileOutputStream);
            fileOutputStream.close();
            return "Xuất File thành công!!!";

        } catch (IOException e) {
            e.printStackTrace();
            return "Lỗi khi xuất file Excel: " + e.getMessage();
        }
    }
    private CTChart createDefaultBarChart(XSSFChart chart, CellReference firstDataCell, CellReference lastDataCell, boolean seriesInCols) {

        CTChart ctChart = chart.getCTChart();
        CTPlotArea ctPlotArea = ctChart.getPlotArea();
        CTBarChart ctBarChart = ctPlotArea.addNewBarChart();
        CTBoolean ctBoolean = ctBarChart.addNewVaryColors();
        ctBoolean.setVal(true);
        ctBarChart.addNewBarDir().setVal(STBarDir.COL);

        int firstDataRow = firstDataCell.getRow();
        int lastDataRow = lastDataCell.getRow();
        int firstDataCol = firstDataCell.getCol();
        int lastDataCol = lastDataCell.getCol();
        String dataSheet = firstDataCell.getSheetName();

        int idx = 0;

        if (seriesInCols) { //the series are in the columns of the data cells

            for (int c = firstDataCol + 1; c < lastDataCol + 1; c++) {
                CTBarSer ctBarSer = ctBarChart.addNewSer();
                CTSerTx ctSerTx = ctBarSer.addNewTx();
                CTStrRef ctStrRef = ctSerTx.addNewStrRef();
                ctStrRef.setF(new CellReference(dataSheet, firstDataRow, c, true, true).formatAsString());

                ctBarSer.addNewIdx().setVal(idx++);
                CTAxDataSource cttAxDataSource = ctBarSer.addNewCat();
                ctStrRef = cttAxDataSource.addNewStrRef();

                ctStrRef.setF(new AreaReference(
                        new CellReference(dataSheet, firstDataRow + 1, firstDataCol, true, true),
                        new CellReference(dataSheet, lastDataRow, firstDataCol, true, true),
                        SpreadsheetVersion.EXCEL2007).formatAsString());

                CTNumDataSource ctNumDataSource = ctBarSer.addNewVal();
                CTNumRef ctNumRef = ctNumDataSource.addNewNumRef();

                ctNumRef.setF(new AreaReference(
                        new CellReference(dataSheet, firstDataRow + 1, c, true, true),
                        new CellReference(dataSheet, lastDataRow, c, true, true),
                        SpreadsheetVersion.EXCEL2007).formatAsString());

                //at least the border lines in Libreoffice Calc ;-)
                ctBarSer.addNewSpPr().addNewLn().addNewSolidFill().addNewSrgbClr().setVal(new byte[] {0,0,0});

            }
        } else { //the series are in the rows of the data cells

            for (int r = firstDataRow + 1; r < lastDataRow + 1; r++) {
                CTBarSer ctBarSer = ctBarChart.addNewSer();
                CTSerTx ctSerTx = ctBarSer.addNewTx();
                CTStrRef ctStrRef = ctSerTx.addNewStrRef();
                ctStrRef.setF(new CellReference(dataSheet, r, firstDataCol, true, true).formatAsString());

                ctBarSer.addNewIdx().setVal(idx++);
                CTAxDataSource cttAxDataSource = ctBarSer.addNewCat();
                ctStrRef = cttAxDataSource.addNewStrRef();

                ctStrRef.setF(new AreaReference(
                        new CellReference(dataSheet, firstDataRow, firstDataCol + 1, true, true),
                        new CellReference(dataSheet, firstDataRow, lastDataCol, true, true),
                        SpreadsheetVersion.EXCEL2007).formatAsString());

                CTNumDataSource ctNumDataSource = ctBarSer.addNewVal();
                CTNumRef ctNumRef = ctNumDataSource.addNewNumRef();

                ctNumRef.setF(new AreaReference(
                        new CellReference(dataSheet, r, firstDataCol + 1, true, true),
                        new CellReference(dataSheet, r, lastDataCol, true, true),
                        SpreadsheetVersion.EXCEL2007).formatAsString());

                //at least the border lines in Libreoffice Calc ;-)
                ctBarSer.addNewSpPr().addNewLn().addNewSolidFill().addNewSrgbClr().setVal(new byte[] {0,0,0});

            }
        }

        //telling the BarChart that it has axes and giving them Ids
        ctBarChart.addNewAxId().setVal(123456);
        ctBarChart.addNewAxId().setVal(123457);

        //cat axis
        CTCatAx ctCatAx = ctPlotArea.addNewCatAx();
        ctCatAx.addNewAxId().setVal(123456); //id of the cat axis
        CTScaling ctScaling = ctCatAx.addNewScaling();
        ctScaling.addNewOrientation().setVal(STOrientation.MIN_MAX);
        ctCatAx.addNewDelete().setVal(false);
        ctCatAx.addNewAxPos().setVal(STAxPos.B);
        ctCatAx.addNewCrossAx().setVal(123457); //id of the val axis
        ctCatAx.addNewTickLblPos().setVal(STTickLblPos.NEXT_TO);

        //val axis
        CTValAx ctValAx = ctPlotArea.addNewValAx();
        ctValAx.addNewAxId().setVal(123457); //id of the val axis
        ctScaling = ctValAx.addNewScaling();
        ctScaling.addNewOrientation().setVal(STOrientation.MIN_MAX);
        ctValAx.addNewDelete().setVal(false);
        ctValAx.addNewAxPos().setVal(STAxPos.L);
        ctValAx.addNewCrossAx().setVal(123456); //id of the cat axis
        ctValAx.addNewTickLblPos().setVal(STTickLblPos.NEXT_TO);

        //legend
        CTLegend ctLegend = ctChart.addNewLegend();
        ctLegend.addNewLegendPos().setVal(STLegendPos.B);
        ctLegend.addNewOverlay().setVal(false);

        return ctChart;

    }

    @PostMapping("/drawBarChart")
    public String drawBarChart() throws Exception {
        String file = "D:\\TH_BT_MD5_SpringBoot\\Demo-ApachePOI-Excel.xlsx";

        FileInputStream fis = new FileInputStream(file);
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet sheet = wb.getSheetAt(0);



        //create empty chart in the sheet
        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 7, 3, 19, 21);

        XSSFChart chart = drawing.createChart(anchor);

        //create the references to the chart data
        CellReference firstDataCell = new CellReference(sheet.getSheetName(), 3, 0, true, true);
        CellReference lastDataCell = new CellReference(sheet.getSheetName(), 6, 4, true, true);

        //create a default bar chart from the data
        CTChart ctBarChart = createDefaultBarChart(chart, firstDataCell, lastDataCell, true);

        //now we can customizing the chart

        //legend position:
        ctBarChart.getLegend().unsetLegendPos();
        ctBarChart.getLegend().addNewLegendPos().setVal(STLegendPos.R);

        //data labels:
        CTBoolean ctboolean = CTBoolean.Factory.newInstance();
        ctboolean.setVal(true);
        ctBarChart.getPlotArea().getBarChartArray(0).addNewDLbls().setShowVal(ctboolean);
        ctboolean.setVal(false);
        ctBarChart.getPlotArea().getBarChartArray(0).getDLbls().setShowSerName(ctboolean);
        ctBarChart.getPlotArea().getBarChartArray(0).getDLbls().setShowPercent(ctboolean);
        ctBarChart.getPlotArea().getBarChartArray(0).getDLbls().setShowLegendKey(ctboolean);
        ctBarChart.getPlotArea().getBarChartArray(0).getDLbls().setShowCatName(ctboolean);
        ctBarChart.getPlotArea().getBarChartArray(0).getDLbls().setShowLeaderLines(ctboolean);
        ctBarChart.getPlotArea().getBarChartArray(0).getDLbls().setShowBubbleSize(ctboolean);

        //val axis maximum:
        ctBarChart.getPlotArea().getValAxArray(0).getScaling().addNewMax().setVal(100);

        //cat axis title:
        ctBarChart.getPlotArea().getCatAxArray(0).addNewTitle().addNewOverlay().setVal(false);
        ctBarChart.getPlotArea().getCatAxArray(0).getTitle().addNewTx().addNewRich().addNewBodyPr();
        ctBarChart.getPlotArea().getCatAxArray(0).getTitle().getTx().getRich().addNewP().addNewR().setT("Bảng doanh số bán hàng Quý 1 - năm 2023");

        //series colors:
        ctBarChart.getPlotArea().getBarChartArray(0).getSerArray(0).getSpPr().addNewSolidFill().addNewSrgbClr().setVal(new byte[]{0, 0, (byte) 255});
        ctBarChart.getPlotArea().getBarChartArray(0).getSerArray(1).getSpPr().addNewSolidFill().addNewSrgbClr().setVal(new byte[]{0, (byte) 255, 0});
//        String pathFile = "D:\\TH_BT_MD5_SpringBoot\\BarChartOutput.xlsx";
        FileOutputStream fileOut = new FileOutputStream(file);
        wb.write(fileOut);
        wb.close();
        fileOut.close();
        return "draw barChart success!!";
    }
}
