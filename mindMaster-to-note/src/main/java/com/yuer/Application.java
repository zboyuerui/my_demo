package com.yuer;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStreamWriter;
import java.util.HashMap;
import java.util.Map;

public class Application {

    public static void main(String[] args) {
        if (args.length < 1) {
            System.out.println("请指定当前目录中的 Excel 文件名");
            return;
        }
        String fileName = args[0];
        // 当前jar包所在路径
        String thisPath = System.getProperty("user.dir");
        
        excelToNote(thisPath, thisPath, fileName);
        
        //批量
        // ...

    }

    /*
     * Excel 导出  html
     * @param inPath 输入Excel路径
     * @param outPath 输出html路径
     */
    @SuppressWarnings("resource")
    public static void excelToNote(String inPath, String outPath, String fileName) {

        // excel文件路径
        String excelPath = inPath + "\\" + fileName;
        // 输出 html文件路径
        String htmlPath = outPath + "\\" + fileName.replace(".xlsx", ".html").replace(".xls", ".html");

        System.out.println("输入Excel:" + excelPath);
        System.out.println("输出html :" + htmlPath);

        try {
            File excel = new File(excelPath);
            // 判断文件是否存在
            if (!excel.exists() || !excel.isFile()) {
                System.out.println("找不到指定的文件");
            }

            Workbook wb;
            // 根据文件后缀（xls/xlsx）进行判断
            FileInputStream fis = new FileInputStream(excel); // 文件流对象
            if (fileName.endsWith(".xls")) {
                wb = new HSSFWorkbook(fis);
            } else if (fileName.endsWith(".xlsx")) {
                wb = new XSSFWorkbook(fis);
            } else {
                System.out.println("错误:不是Excel文件类型!");
                return;
            }

            StringBuilder sb = new StringBuilder();

            // 开始解析Excel
            Sheet sheet = wb.getSheetAt(0); // 读取sheet 0

            int firstRowIndex = sheet.getFirstRowNum() + 1; // 第一行是列名，所以不读
            int lastRowIndex = sheet.getLastRowNum(); // 最后一行编号

            for (int rIndex = firstRowIndex; rIndex <= lastRowIndex; rIndex++) { // 遍历行
                Row row = sheet.getRow(rIndex);
                if (row != null) {
                    int firstCellIndex = row.getFirstCellNum(); // 当前行有值的第一列
                    // int lastCellIndex = row.getLastCellNum();
                    // for (int cIndex = firstCellIndex; cIndex < lastCellIndex;
                    // cIndex++) { // 遍历列

                    Cell cell = row.getCell(firstCellIndex); // 这里每行 只有一列有数据
                    if (cell != null) {
                        // 操作不为空的每一个单元格
                        // System.out.println("第
                        // "+rIndex+"行,第"+cIndex+"列,内容为:\t"+cell.toString());
                        int nextIndex = rIndex == lastRowIndex ? 0 : sheet.getRow(rIndex + 1).getFirstCellNum();
                        String str = handler(cell.toString(), firstCellIndex, nextIndex);
                        sb.append(str);
                    }
                }
            }

            // 优化-去掉多余的标签 (合并相邻的无序列表,之前每一行是单独为一个无序列表)
            String body = sb.toString();
            while (body.contains("</ul><ul>")) {
                body = body.replace("</ul><ul>", "");
            }

            // 输出文件
            File outFile = new File(htmlPath);
            if (outFile.exists())
                outFile.delete();
            outFile.createNewFile();
            BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(outFile), "utf-8"));
            writer.write(head);
            writer.write(body);
            writer.write(tail);
            writer.close();

            System.out.println("----OK----");

        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    // 文件头--默认字体 10号
    private static String head = "<html><head><title>Evernote Export</title><basefont face=\"微软雅黑\" size=\"2\" />"
            + "<meta http-equiv=\"Content-Type\" content=\"text/html;charset=utf-8\" />"
            + "<meta name=\"exporter-version\" content=\"YXBJ Windows/601302 (zh-CN, DDL); Windows/10.0.0 (Win64);"
            + "EDAMVersion=V2;\"/><style>body, td {font-family: 微软雅黑;font-size: 10pt;}</style></head><body><div>";
    // 文件尾缀
    private static String tail = "</div></body></html>";
    // 标签
    private static Map<String, String> tags = new HashMap<String, String>();
    static {
        tags.put("ul", "<ul>*</ul>"); // 无序列表
        tags.put("li", "<li>*</li>"); // 列表中的元素
        tags.put("div2", "<div style=\"margin-left: ?px;\">*</div>"); // 有缩进的内容
        tags.put("div", "<div>*</div>"); // 内容
        tags.put("b", "<b>*</b>"); // 加粗
        tags.put("font", "<font style=\"font-size: ?pt;\">*</font>"); // 设置字体大小
    }

    /*
     * 组装标签
     * 
     * @param cell 单元格内容
     * 
     * @param cellIndex 单元格列标
     * 
     * @param nextIndex 下一行单元格列标（每一行只有一个单元格有数据）
     */
    private static String handler(String cell, int cellIndex, int nextIndex) {
        // 叶子节点,有换行符时,改为缩进
        if (cell.contains("\n") && cellIndex >= nextIndex) {
            String result = "";
            String[] ceList = cell.split("\n");
            for (String ce : ceList) {
                String ces = tags.get("div2").replace("*", ce).replace("?", String.valueOf(cellIndex * 40));
                result = result + ces;
            }
            return result;
        } else {
            // 不是叶子节点,需要去掉换行符
            if (cell.contains("\n")) {
                cell = cell.replace("\n", "");
            }
            // 总结点--16号字体,加粗
            if (cellIndex == 0) {
                cell = tags.get("font").replace("*", cell).replace("?", String.valueOf(16));
                cell = tags.get("b").replace("*", cell);
            }
            // 一级结点--14号字体,加粗
            if (cellIndex == 1) {
                cell = tags.get("font").replace("*", cell).replace("?", String.valueOf(14));
                cell = tags.get("b").replace("*", cell);
            }
            // 二级结点--12号字体
            if (cellIndex == 2) {
                cell = tags.get("font").replace("*", cell).replace("?", String.valueOf(12));
            }

            cell = tags.get("div").replace("*", cell);
            cell = tags.get("li").replace("*", cell);
            for (int i = 0; i <= cellIndex; i++) {
                cell = tags.get("ul").replace("*", cell);
            }
            return cell;
        }
    }

}
