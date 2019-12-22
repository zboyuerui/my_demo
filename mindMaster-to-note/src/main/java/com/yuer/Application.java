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
            System.out.println("��ָ����ǰĿ¼�е� Excel �ļ���");
            return;
        }
        String fileName = args[0];
        // ��ǰjar������·��
        String thisPath = System.getProperty("user.dir");
        
        excelToNote(thisPath, thisPath, fileName);
        
        //����
        // ...

    }

    /*
     * Excel ����  html
     * @param inPath ����Excel·��
     * @param outPath ���html·��
     */
    @SuppressWarnings("resource")
    public static void excelToNote(String inPath, String outPath, String fileName) {

        // excel�ļ�·��
        String excelPath = inPath + "\\" + fileName;
        // ��� html�ļ�·��
        String htmlPath = outPath + "\\" + fileName.replace(".xlsx", ".html").replace(".xls", ".html");

        System.out.println("����Excel:" + excelPath);
        System.out.println("���html :" + htmlPath);

        try {
            File excel = new File(excelPath);
            // �ж��ļ��Ƿ����
            if (!excel.exists() || !excel.isFile()) {
                System.out.println("�Ҳ���ָ�����ļ�");
            }

            Workbook wb;
            // �����ļ���׺��xls/xlsx�������ж�
            FileInputStream fis = new FileInputStream(excel); // �ļ�������
            if (fileName.endsWith(".xls")) {
                wb = new HSSFWorkbook(fis);
            } else if (fileName.endsWith(".xlsx")) {
                wb = new XSSFWorkbook(fis);
            } else {
                System.out.println("����:����Excel�ļ�����!");
                return;
            }

            StringBuilder sb = new StringBuilder();

            // ��ʼ����Excel
            Sheet sheet = wb.getSheetAt(0); // ��ȡsheet 0

            int firstRowIndex = sheet.getFirstRowNum() + 1; // ��һ�������������Բ���
            int lastRowIndex = sheet.getLastRowNum(); // ���һ�б��

            for (int rIndex = firstRowIndex; rIndex <= lastRowIndex; rIndex++) { // ������
                Row row = sheet.getRow(rIndex);
                if (row != null) {
                    int firstCellIndex = row.getFirstCellNum(); // ��ǰ����ֵ�ĵ�һ��
                    // int lastCellIndex = row.getLastCellNum();
                    // for (int cIndex = firstCellIndex; cIndex < lastCellIndex;
                    // cIndex++) { // ������

                    Cell cell = row.getCell(firstCellIndex); // ����ÿ�� ֻ��һ��������
                    if (cell != null) {
                        // ������Ϊ�յ�ÿһ����Ԫ��
                        // System.out.println("��
                        // "+rIndex+"��,��"+cIndex+"��,����Ϊ:\t"+cell.toString());
                        int nextIndex = rIndex == lastRowIndex ? 0 : sheet.getRow(rIndex + 1).getFirstCellNum();
                        String str = handler(cell.toString(), firstCellIndex, nextIndex);
                        sb.append(str);
                    }
                }
            }

            // �Ż�-ȥ������ı�ǩ (�ϲ����ڵ������б�,֮ǰÿһ���ǵ���Ϊһ�������б�)
            String body = sb.toString();
            while (body.contains("</ul><ul>")) {
                body = body.replace("</ul><ul>", "");
            }

            // ����ļ�
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

    // �ļ�ͷ--Ĭ������ 10��
    private static String head = "<html><head><title>Evernote Export</title><basefont face=\"΢���ź�\" size=\"2\" />"
            + "<meta http-equiv=\"Content-Type\" content=\"text/html;charset=utf-8\" />"
            + "<meta name=\"exporter-version\" content=\"YXBJ Windows/601302 (zh-CN, DDL); Windows/10.0.0 (Win64);"
            + "EDAMVersion=V2;\"/><style>body, td {font-family: ΢���ź�;font-size: 10pt;}</style></head><body><div>";
    // �ļ�β׺
    private static String tail = "</div></body></html>";
    // ��ǩ
    private static Map<String, String> tags = new HashMap<String, String>();
    static {
        tags.put("ul", "<ul>*</ul>"); // �����б�
        tags.put("li", "<li>*</li>"); // �б��е�Ԫ��
        tags.put("div2", "<div style=\"margin-left: ?px;\">*</div>"); // ������������
        tags.put("div", "<div>*</div>"); // ����
        tags.put("b", "<b>*</b>"); // �Ӵ�
        tags.put("font", "<font style=\"font-size: ?pt;\">*</font>"); // ���������С
    }

    /*
     * ��װ��ǩ
     * 
     * @param cell ��Ԫ������
     * 
     * @param cellIndex ��Ԫ���б�
     * 
     * @param nextIndex ��һ�е�Ԫ���б꣨ÿһ��ֻ��һ����Ԫ�������ݣ�
     */
    private static String handler(String cell, int cellIndex, int nextIndex) {
        // Ҷ�ӽڵ�,�л��з�ʱ,��Ϊ����
        if (cell.contains("\n") && cellIndex >= nextIndex) {
            String result = "";
            String[] ceList = cell.split("\n");
            for (String ce : ceList) {
                String ces = tags.get("div2").replace("*", ce).replace("?", String.valueOf(cellIndex * 40));
                result = result + ces;
            }
            return result;
        } else {
            // ����Ҷ�ӽڵ�,��Ҫȥ�����з�
            if (cell.contains("\n")) {
                cell = cell.replace("\n", "");
            }
            // �ܽ��--16������,�Ӵ�
            if (cellIndex == 0) {
                cell = tags.get("font").replace("*", cell).replace("?", String.valueOf(16));
                cell = tags.get("b").replace("*", cell);
            }
            // һ�����--14������,�Ӵ�
            if (cellIndex == 1) {
                cell = tags.get("font").replace("*", cell).replace("?", String.valueOf(14));
                cell = tags.get("b").replace("*", cell);
            }
            // �������--12������
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
