package com.timmy.demo;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;

public class TransToExcel {
	 /**
     * ����excel�ļ�
     * @param title  ��sheet������
     * @param headers  ��ͷ
     * @param dataList  ���ĵ�Ԫ��
     * @param out  �����
     */
    public void exporteExcel(String title,String[] headers,List<Map<String, String>> dataList,OutputStream out){
        HSSFWorkbook workBook = new HSSFWorkbook();
        createSheet(title, headers, dataList, workBook);
        createSheet(title+"2", headers, dataList, workBook);
        try {
            workBook.write(out);
        }catch (IOException e){
            System.out.println("д���ļ�ʧ��"+e.getMessage());
        }
    }

    /**
     * ����sheet
     * @param title  sheet������
     * @param headers  ��ͷ
     * @param dataList  ���ĵ�Ԫ��
     */
    private void createSheet(String title, String[] headers, List<Map<String, String>> dataList, HSSFWorkbook workBook) {
        HSSFSheet sheet = workBook.createSheet(title);
//        sheet.setDefaultColumnWidth(15);
        //���ñ�ͷ����ͨ��Ԫ��ĸ�ʽ
        HSSFCellStyle headStyle = setHeaderStyle(workBook);
        HSSFCellStyle bodyStyle = setBodyStyle(workBook);

        createBody(dataList, sheet, bodyStyle);
        createHeader(headers, sheet, headStyle);
    }
    
    /**
     * �������ĵ�Ԫ��
     * @param dataList ��������
     * @param sheet ��
     * @param bodyStyle ��Ԫ���ʽ
     */
    private void createBody(List<Map<String, String>> dataList, HSSFSheet sheet, HSSFCellStyle bodyStyle) {
        for (int a=0;a<dataList.size();a++){
            HSSFRow row = sheet.createRow(a+1);
            Map<String, String>map=dataList.get(a);
            String []data=new String[7];
            data[0]=map.get("�˿ں�");
            data[1]=map.get("��������");
            data[2]=map.get("��������");
            data[3]=map.get("��������");
            data[4]=map.get("������ַ");
            data[5]=map.get("��¼ֵ");
            data[6]=map.get("��¼ʱ��");

            for(int j=0;j<data.length;j++){
           // int j=0;
           // for (String key : map.keySet()) {
//��ȡmap
           // for (String v : map.values()) {

                HSSFCell cell = row.createCell(j);
                cell.setCellStyle(bodyStyle);
                HSSFRichTextString textString = new HSSFRichTextString(data[j]);
                cell.setCellValue(textString);
                //j++;
            }
        }
    }

    /**
     * ������ͷ
     * @param headers  ��ͷ
     * @param sheet ��
     * @param headStyle ��ͷ��ʽ
     */
    private void createHeader(String[] headers, HSSFSheet sheet, HSSFCellStyle headStyle) {
        HSSFRow row = sheet.createRow(0);
        for (int i=0;i<headers.length;i++){
            HSSFCell cell = row.createCell(i);
            cell.setCellStyle(headStyle);
            HSSFRichTextString textString = new HSSFRichTextString(headers[i]);
            cell.setCellValue(textString);
            sheet.autoSizeColumn((short)i);
        }
    }

    /**
     * �������ĵ�Ԫ���ʽ
     * @param workBook
     * @return
     */
    private HSSFCellStyle setBodyStyle(HSSFWorkbook workBook) {
        HSSFCellStyle style2 = workBook.createCellStyle();
        style2.setFillForegroundColor(HSSFColor.WHITE.index);
        style2.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        style2.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style2.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        style2.setBorderRight(HSSFCellStyle.BORDER_THIN);
        style2.setBorderTop(HSSFCellStyle.BORDER_THIN);
        style2.setAlignment(HSSFCellStyle.ALIGN_LEFT);

        HSSFFont font2 = workBook.createFont();
        font2.setFontName("΢���ź�");
        font2.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
        style2.setFont(font2);
        return style2;
    }

    /**
     * ���ñ�ͷ��ʽ
     * @param workBook
     * @return
     */
    private HSSFCellStyle setHeaderStyle(HSSFWorkbook workBook) {
        HSSFCellStyle style = workBook.createCellStyle();
        style.setFillForegroundColor(HSSFColor.LIGHT_YELLOW.index);
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
        style.setAlignment(HSSFCellStyle.ALIGN_LEFT);

        HSSFFont font = workBook.createFont();
        font.setFontName("΢���ź�");
        font.setFontHeightInPoints((short)12);
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        style.setFont(font);
        return style;
    }

    public static void main(String[] args) {
        TransToExcel transToExcel = new TransToExcel();
        try {
            String path = System.getProperty("user.dir");
            OutputStream os = new FileOutputStream(path+"/���ݼ�¼��.xls");
            String[] headers = {"�˿ں�","��������","��������","��������","������ַ","��¼ֵ","��¼ʱ��"};
            Date date = new Date(System.currentTimeMillis());
            SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
            String time = sdf.format(date);
           // String[][] list = {{"�׹�������","1000000","20100914001000000000","haha","hahah","http://baigungun.blog.com.cn/index",time},{"����ҹ��","300000","20100914002","haha","hahah","http://yehua.com.cn/index",time}};
            List<Map<String, String>>datas=new ArrayList<>();
    		Map<String,String>map=new HashMap<>();
    		Map<String, String>map2=new HashMap<>();
    		map.put("�˿ں�","�׹�������");
    		map.put("��������", "1000000");
    		map.put("��������", "20100914001000000000");
    		map.put("��������", "haha");
    		map.put("������ַ", "hahah");
    		map.put("��¼ֵ", "http://baigungun.blog.com.cn/index");
    		map.put("��¼ʱ��", time);
    		
//    		map2.put("id","02");
//    		map2.put("name", "kxy");
//    		map2.put("pass", "0111");
    		map2.put("�˿ں�","�׹�������");
    		map2.put("��������", "1000000");
    		map2.put("��������", "20100914001000000000");
    		map2.put("��������", "haha");
    		map2.put("������ַ", "hahah");
    		map2.put("��¼ֵ", "http://baigungun.blog.com.cn/index");
    		map2.put("��¼ʱ��", time);
    		datas.add(map);
    		datas.add(map2);
            transToExcel.exporteExcel("ѧ����",headers,datas,os);
            os.close();

        }catch (FileNotFoundException e){
            System.out.println("�޷��ҵ��ļ�");
        }catch (IOException e){
            System.out.println("д���ļ�ʧ��");
        }
    }

}
