package com.hwolf;

import com.hwolf.utils.ReadFile;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import java.io.*;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import java.util.stream.Collectors;

import static sun.misc.MessageUtils.out;

/**
 * @author hwolf
 * @email h.wolf@qq.com
 * @date 2017/11/25.
 */
public class POIExcel {
    // root
    private static String ROOT_FILEPATH = "/Users/hwolf/Downloads";
    // format date
    private static SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");

    // format int
    private static DecimalFormat df = new DecimalFormat("0");
    // format float
    private static DecimalFormat nf = new DecimalFormat("0.00");

    /**
     * according to extension return Workbook type
     * @param file
     * @return
     * @throws Exception
     */
    public static Workbook getWorkbook(File file) throws Exception {
        String fileName = file.getName();
        // print the file name
        System.out.println(fileName);
        System.out.println(file.getAbsoluteFile());
        // get extension
        String extension = fileName.lastIndexOf(".") == -1 ? "" : fileName
                .substring(fileName.lastIndexOf(".") + 1);
        System.out.println(extension);
        FileInputStream fis = new FileInputStream(file);
        if (extension.equals("xls")) {
            return new HSSFWorkbook(fis);
        } else if (extension.equals("xlsx")) {
            return new XSSFWorkbook(fis);
        } else {
//            throw new Exception("no support！");
            return null;
        }
    }

    /**
     *
     * @param startRow
     * @param city your city
     * @param columnHead
     * @param keys other keys
     * @return
     * @throws IOException
     */
    public static List<List<Object>> readExcel(int startRow,String city,String columnHead,String[] keys) throws IOException {
        ReadFile readFile = new ReadFile();
        List<File> fileList = new LinkedList<>();

        List<File> files = readFile.fileToList(new File(ROOT_FILEPATH + "/city"),fileList);

        System.out.println(files.size());

        List<List<Object>> list = new LinkedList<>();
        try {

            for (int i = 0; i < files.size(); i++) {
                Workbook wb = null;
                //catch some workbook error information
                try{
                    wb = getWorkbook(files.get(i));
                } catch (Exception e){
                    throw new IOException("Reading the : "+files.get(i).getAbsolutePath() +" error！");
                }
                //if workbook isn't exist continue loop
                if (wb == null){
                    continue;
                }
                //getNumberOfSheets
                int sheetsSize = wb.getNumberOfSheets();
                //foreach all sheets
                for (int startSheet = 0; startSheet < sheetsSize; startSheet++) {
                    Sheet sheet = wb.getSheetAt(startSheet);
                    Object value = null;
                    Row row = null;
                    Cell cell = null;
                    CellStyle cs = null;
                    String csStr = null;
                    Double numval = null;
                    Iterator<Row> rows = sheet.rowIterator();
                    //foreach rows
                    while (rows.hasNext()) {
                        row = (Row) rows.next();
                        if (row.getRowNum() >= startRow) {
                            //use LinkedList
                            List<Object> cellList = new LinkedList<Object>();
                            Iterator<Cell> cells = row.cellIterator();
                            //foreach cell
                            while (cells.hasNext()) {
                                cell = (Cell) cells.next();
                                switch (cell.getCellType()) {
                                    case Cell.CELL_TYPE_NUMERIC:
                                        cs = cell.getCellStyle();
                                        csStr = cs.getDataFormatString();
                                        numval = cell.getNumericCellValue();
                                        if ("@".equals(csStr)) {
                                            value = df.format(numval);
                                        } else if ("General".equals(csStr)) {
                                            value = nf.format(numval);
                                        } else {
                                            try {
                                                // TODO
//                                                value = sdf.format(HSSFDateUtil.getJavaDate(numval));
                                                value = df.format(numval);
                                            } catch (Exception e) {
                                                //捕获空指针异常
                                                throw new NullPointerException(
                                                        files.get(i).getCanonicalPath() + "path, " +
                                                                sheet.getSheetName() + "sheet name, " + row.getRowNum() + "row , "
                                                                + cell.getCellType() + "cell's type, "
                                                                + cell.getColumnIndex() + "column.");
                                            }
                                        }
                                        break;
                                    case Cell.CELL_TYPE_STRING:
                                        value = cell.getStringCellValue();
                                        break;
                                    case Cell.CELL_TYPE_FORMULA:
                                        cell.setCellType(Cell.CELL_TYPE_STRING);
                                        if (!cell.getStringCellValue().equals("")) {
                                            value = cell.getStringCellValue();
                                        } else {
                                            value = cell.getNumericCellValue() + "";
                                        }
                                        break;
                                    case Cell.CELL_TYPE_BLANK:
                                        value = "";
                                        break;
                                    case Cell.CELL_TYPE_BOOLEAN:
                                        value = cell.getBooleanCellValue();
                                        break;
                                    default:
                                        value = cell.toString();
                                }
                                cellList.add(value + "|");
                            }
                            String cellString = cellList.toString();
                            // judge city and head
                            if (cellString.contains(city) || cellString.contains(columnHead)) {
                                cellList.add(files.get(i).getName() + "|" + sheet.getSheetName());
                                list.add(cellList);
                            }
                            // judge other keys
                            for (String key : keys) {
                                if (cellString.contains(key)){
                                    list.add(cellList);
                                    break;
                                }
                            }
                        }
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return list;
    }


    /**
     * gather up all data to one workbook
     * @param list cell list
     * @param targetXlsxPath xlsx file
     * @throws IOException
     */
    public static void writeExcel(List<List<Object>> list, String targetXlsxPath) throws IOException {
        // 判断文件路径是否为空
        if (targetXlsxPath == null || targetXlsxPath.equals("")) {
            throw new IOException("file path not null");
        }

        // 判断列表是否有数据，如果没有数据，则返回
        if (list == null || list.size() == 0) {
            out("list was null");
            return;
        }

        try {
            // xlsx
            XSSFWorkbook wb = null;
            // judge file exist
            File file = new File(targetXlsxPath);
            if (file.exists()) {
                wb = new XSSFWorkbook(new FileInputStream(targetXlsxPath));
            } else {
                throw new IOException(targetXlsxPath+" isn't exist");
            }
            // get the sheet column
            int t = wb.getSheetAt(0).getPhysicalNumberOfRows();

            FileOutputStream outputStream = new FileOutputStream(targetXlsxPath);

            // get row from list
            for (List row : list) {
                if (row == null) {
                    continue;
                }
                // write to first sheet
                Sheet sheet = wb.getSheetAt(0);
                // add data not rewrite
                Row r = null;
                r = sheet.createRow(t++);

                // CellStyle newstyle = wb.createCellStyle();

                // String[] rowSplitString = row.toString().split("[\\u007c]");

                // split condition is '|'
                String[] rowSplitString = row.stream().collect(Collectors.joining("")).toString().split("[\\u007c]");
                for (int i = 0; i < rowSplitString.length; i++) {
                    // createCell
                    Cell cell = r.createCell(i);

                    if(rowSplitString[i] == null || rowSplitString[i] == ""){
                        continue;
                    }

                    cell.setCellValue(rowSplitString[i]);
                }
            }
            // write out
            wb.write(outputStream);
            outputStream.flush();
            outputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) throws IOException {
        // your keys
        String[] keys = {"","","",""};
        List<List<Object>> list = readExcel(0,"","",keys);
        System.out.println(list.size());
        writeExcel(list,ROOT_FILEPATH + "/yourTarget.xlsx");
    }
}
