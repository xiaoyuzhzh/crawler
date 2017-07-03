package com.example;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;

import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;

/**
 * Created by xiaoy on 2017/7/3.
 */
public class CheckInTimeExcelGenerator {

    private static List<Integer> weekRowNumbers = new LinkedList<>();

    public static void main(String[] args) throws Exception {
        String startDateTemplate = "2017-0#mm#-01";
        String endDateTemplate = "2017-0#mm#-01";
        for(int i = 6; i< 13 ; i ++){
            generate("d:/"+i+"月模板.xlsx",startDateTemplate.replaceAll("#mm#",String.valueOf(i)),endDateTemplate.replaceAll("#mm#",String.valueOf(i+1)));
        }
    }

    /**
     * 生成文件
     * @param path
     * @throws Exception
     */
    public static void generate(String path,String startDateStr,String endDateStr) throws Exception {
        XSSFWorkbook wb = new XSSFWorkbook();
        FileOutputStream fileOut = new FileOutputStream(path);


        Sheet sheet1 = wb.createSheet("sheet1");
        weekRowNumbers.clear();

        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
        Date startDate = simpleDateFormat.parse(startDateStr);
        Date endDate = simpleDateFormat.parse(endDateStr);
        XSSFCellStyle defaultStyle = wb.createCellStyle();//默认格式
        getDefaultBorderStyle(defaultStyle);
        XSSFCellStyle headerStyle = wb.createCellStyle();//表头格式
        headerStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(244, 176, 132, 50)));
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        getDefaultBorderStyle(headerStyle);

        int rowNumber = 0;

        //设置表头
        rowNumber = createRow(sheet1, rowNumber, headerStyle, "日期", "星期","上班签到", "下班签退", "时间", "工作开始", "工作结束", "工作时长", "加班时长", "餐贴","工作日","周末");

        //设置日期
        rowNumber = setDate(wb, sheet1, rowNumber, startDate, endDate, defaultStyle);

        //设置底部cell
        rowNumber = setBottomCells(wb,sheet1,rowNumber,defaultStyle);


        wb.write(fileOut);
        fileOut.close();
    }



    private static int setBottomCells(XSSFWorkbook wb, Sheet sheet, int rowNumber, XSSFCellStyle defaultStyle) {

        List<CreateCellParam> cs = new LinkedList<>();
        for(int i = 0 ; i<7; i++){
            cs.add(new CreateCellParam(defaultStyle,null,null));
        }

        CreateCellParam workTimeSum = new CreateCellParam(defaultStyle,null,getWorkTimeSumFormula(rowNumber)); cs.add(workTimeSum);//工作时间汇总
        CreateCellParam extraTimeSum = new CreateCellParam(defaultStyle,null,getExtraTimeSumFormula(rowNumber));cs.add(extraTimeSum);//加班时间汇总
        CreateCellParam dietFeeSum = new CreateCellParam(defaultStyle,null,getDietFeeSumFormula(rowNumber));cs.add(dietFeeSum);//餐补汇总

        CreateCellParam normalExtraTimeSum = new CreateCellParam(defaultStyle,null,getNormalExtraTimeFormula(rowNumber));cs.add(normalExtraTimeSum);//平常加班汇总
        CreateCellParam weekendExtraTime = new CreateCellParam(defaultStyle,null,getWeekendExtraTimeFormula(rowNumber));cs.add(weekendExtraTime);//平常加班汇总
        rowNumber = createRow2(sheet,rowNumber,cs);
        return rowNumber;
    }

    private static String getWeekendExtraTimeFormula(int rowNumber) {
        String template = "ROUND(#cells#,0)";
        StringBuilder builder = new StringBuilder();
        for (Integer weekRowNumber : weekRowNumbers) {
            builder.append("H"+String.valueOf(weekRowNumber)+"+");
        }
        builder.deleteCharAt(builder.length()-1);
        String cells = builder.toString();
        String result = template.replaceAll("#cells#",cells);
        return result;
    }

    private static String getNormalExtraTimeFormula(int rowNumber) {
        String template = "ROUND(I#rownum#-L#rownum#,0)";
        String result = template.replaceAll("#rownum#",String.valueOf(rowNumber+1));
        return result;
    }

    private static String getDietFeeSumFormula(int rowNumber) {
        String template = "SUM(I2:I#end#)";
        String end = String.valueOf(rowNumber);
        String result = template.replaceAll("#end#",end);
        return result;
    }

    private static String getExtraTimeSumFormula(int rowNumber) {
        String template = "SUM(J2:J#end#)";
        String end = String.valueOf(rowNumber);
        String result = template.replaceAll("#end#",end);
        return result;
    }

    private static String getWorkTimeSumFormula(int rowNumber) {
        String template = "SUM(H2:H#end#)";
        String end = String.valueOf(rowNumber);
        String result = template.replaceAll("#end#",end);
        return result;
    }

    /**
     * 设置默认边框格式
     * @param defaultStyle
     */
    private static void getDefaultBorderStyle(XSSFCellStyle defaultStyle) {
        defaultStyle.setBorderColor(XSSFCellBorder.BorderSide.TOP, new XSSFColor(new java.awt.Color(0, 0, 0)));
        defaultStyle.setBorderColor(XSSFCellBorder.BorderSide.LEFT, new XSSFColor(new java.awt.Color(0, 0, 0)));
        defaultStyle.setBorderColor(XSSFCellBorder.BorderSide.RIGHT, new XSSFColor(new java.awt.Color(0, 0, 0)));
        defaultStyle.setBorderColor(XSSFCellBorder.BorderSide.BOTTOM, new XSSFColor(new java.awt.Color(0, 0, 0)));
        defaultStyle.setBorderTop(BorderStyle.THIN);
        defaultStyle.setBorderLeft(BorderStyle.THIN);
        defaultStyle.setBorderRight(BorderStyle.THIN);
        defaultStyle.setBorderBottom(BorderStyle.THIN);
    }

    /**
     * 设置考勤时间数据
     * @param wb
     * @param sheet
     * @param rowNumber
     * @param startDate
     * @param endDate
     * @param defaultStyle
     * @return
     */
    private static int setDate(XSSFWorkbook wb, Sheet sheet, int rowNumber, Date startDate, Date endDate, XSSFCellStyle defaultStyle) {
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(startDate);
        String dateStr = "";//日期
        String weekStr = "";//星期
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
        XSSFCellStyle weekEndStyle = wb.createCellStyle();//周末格式
        weekEndStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(153, 204, 255, 100)));
        weekEndStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        getDefaultBorderStyle(weekEndStyle);

        //设置列格式
        for(int i = 0; i<10 ; i++){
            sheet.setColumnWidth(i,18*256);
        }

        while (calendar.getTime().getTime() < endDate.getTime()) {
            dateStr = simpleDateFormat.format(calendar.getTime());
            weekStr = getWeekStr(calendar.get(Calendar.DAY_OF_WEEK));
            List<CreateCellParam> cs = new LinkedList<>();
            CreateCellParam dateC = new CreateCellParam(defaultStyle,dateStr,null);cs.add(dateC);
            CreateCellParam weekEndC = new CreateCellParam(defaultStyle,weekStr,null);cs.add(weekEndC);
            CreateCellParam startCheckIn = new CreateCellParam(defaultStyle,null,null);cs.add(startCheckIn);
            CreateCellParam endCheckIn = new CreateCellParam(defaultStyle,null,null);cs.add(endCheckIn);
            CreateCellParam time = new CreateCellParam(defaultStyle,null,getTimeFormula(rowNumber));cs.add(time);
            CreateCellParam workStart = new CreateCellParam(defaultStyle,null,getWorkStartFormula(rowNumber));cs.add(workStart);
            CreateCellParam workEnd = new CreateCellParam(defaultStyle,null,getWorkEndFormula(rowNumber));cs.add(workEnd);
            CreateCellParam workTime = new CreateCellParam(defaultStyle,null,getWorkTimeFormula(rowNumber));cs.add(workTime);
            CreateCellParam extraTime = new CreateCellParam(defaultStyle,null,getExtraTimeFormula(rowNumber));cs.add(extraTime);
            CreateCellParam dietFee = new CreateCellParam(defaultStyle,null,getDietFeeFormula(rowNumber));cs.add(dietFee);
            for(int i = 0 ; i< 2 ; i++){
                cs.add(new CreateCellParam(defaultStyle,null,null));
            }
            if ("星期六".equals(weekStr) || "星期日".equals(weekStr)) {
                for (CreateCellParam c : cs) {
                    c.cellStyle = weekEndStyle;
                }
                weekRowNumbers.add(rowNumber+1);
            }
            rowNumber = createRow2(sheet, rowNumber,cs);
            calendar.add(Calendar.DAY_OF_MONTH, 1);
        }

        return rowNumber;
    }

    private static String getDietFeeFormula(int rowNumber) {
        String template = "IF(OR(B#rownum#=\"星期六\",B#rownum#=\"星期日\"),IF(H#rownum#-8>=0,30,IF(H#rownum#-4>=0,15,0)),IF(H#rownum#-11>=0,15,0))";
        String result = template.replaceAll("#rownum#",String.valueOf(rowNumber+1));
        return result;
    }

    private static String getExtraTimeFormula(int rowNumber) {
        String template = "(IF(OR(B#rownum#=\"星期六\",B#rownum#=\"星期日\"),H#rownum#+0,IF(H#rownum#>0,IF(H#rownum#-12>=0,H#rownum#-9,0),0)))";
        String result = template.replaceAll("#rownum#",String.valueOf(rowNumber+1));
        return result;
    }

    private static String getWorkTimeFormula(int rowNumber) {
        String template = "ROUND(IF(AND(F#rownum#<>\"\",G#rownum#<>\"\"),TEXT(INT(G#rownum#-F#rownum#)*24+MOD(G#rownum#-F#rownum#,1)*24,\"0.0\"),0),1)";
        String result = template.replaceAll("#rownum#",String.valueOf(rowNumber+1));
        return result;
    }

    private static String getWorkEndFormula(int rowNumber) {
        String template = "IF(D#rownum#<>\"\",TEXT(D#rownum#,\"yyyy/m/d hh:mm:ss\"),\"\")";
        String result = template.replaceAll("#rownum#",String.valueOf(rowNumber+1));
        return result;
    }

    private static String getWorkStartFormula(int rowNumber) {
        String template = "IF(C#rownum#<>\"\",TEXT(A#rownum#,\"yyyy/m/d\")&\" \"&TEXT(E#rownum#,\"hh:mm:ss\"),\"\")";
        String result = template.replaceAll("#rownum#",String.valueOf(rowNumber+1));
        return result;
    }

    private static String getTimeFormula(int rowNumber) {
        String template = "IF(C#rownum#<>\"\",TEXT(IF(MINUTE(C#rownum#)=0,HOUR(C#rownum#)&\":00:00\",IF(MINUTE(C#rownum#)<30,HOUR(C#rownum#)&\":30:00\",IF(MINUTE(C#rownum#)>30,(HOUR(C#rownum#)+1)&\":00:00\",HOUR(C#rownum#)&\":\"&MINUTE(C#rownum#)&\":00\"))),\"hh:mm:ss\"),\"\")";
        String result = template.replaceAll("#rownum#",String.valueOf(rowNumber+1));
        return result;
    }

    private static int createRow2(Sheet sheet, int rowNum,List<CreateCellParam> createCellParams) {
        Row row = sheet.createRow(rowNum);
        int cellNumber = 0;//cellNumber
        Cell cell;
        for (CreateCellParam c : createCellParams) {
            cell = row.createCell(cellNumber);
            setCell2(cell, c);
            cellNumber++;
        }
        return ++ rowNum;
    }

    private static void setCell2(Cell cell, CreateCellParam c) {
        if(c == null) {
            return;
        }

        Object o = c.cellValue;
        if(o!=null){
            if (o instanceof Number) {//数字
                cell.setCellValue((double) o);
            } else if (o instanceof String) {//字符串
                cell.setCellValue((String) o);
            }
        }

        if(c.cellStyle!=null){
            cell.setCellStyle(c.cellStyle);
        }

        if(c.formula!=null){
            cell.setCellFormula(c.formula);
        }
    }


    /**
     * 获取中文星期数
     * @param i
     * @return
     */
    private static String getWeekStr(int i) {
        switch (i) {
            case 1:
                return "星期日";
            case 2:
                return "星期一";
            case 3:
                return "星期二";
            case 4:
                return "星期三";
            case 5:
                return "星期四";
            case 6:
                return "星期五";
            case 7:
                return "星期六";
        }
        return null;
    }

    /**
     * 创建一行记录
     * @param sheet
     * @param rowNum
     * @param cellStyle
     * @param object
     * @return
     */
    private static int createRow(Sheet sheet, int rowNum, CellStyle cellStyle, Object... object) {
        Row row = sheet.createRow(rowNum);
        int cellNumber = 0;//cellNumber
        Cell cell;
        for (Object o : object) {
            cell = row.createCell(cellNumber);
            setCell(cellStyle, cell, o);
            cellNumber++;
        }
        return ++rowNum;
    }

    /**
     * 设置cell
     * @param cellStyle
     * @param cell
     * @param o
     */
    private static void setCell(CellStyle cellStyle, Cell cell, Object o) {
        if (o instanceof Number) {//数字
            cell.setCellValue((double) o);
        } else if (o instanceof String) {//字符串
            cell.setCellValue((String) o);
        }
        if (cellStyle != null) {
            cell.setCellStyle(cellStyle);
        }
    }

    static class CreateCellParam{
        public CellStyle cellStyle;//单元格格式
        public Object cellValue;//单元格的值
        public String formula;//函数

        public CreateCellParam(CellStyle cellStyle,Object cellValue,String formula){
            this.cellStyle = cellStyle;
            this.cellValue = cellValue;
            this.formula = formula;
        }

    }
}
