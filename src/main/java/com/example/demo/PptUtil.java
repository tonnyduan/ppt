package com.example.demo;


import org.apache.poi.sl.usermodel.TextParagraph;
import org.apache.poi.xslf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTable;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTableRow;

import javax.imageio.stream.FileImageInputStream;
import java.awt.*;
import java.io.*;
import java.util.List;

/**
 * 操作ppt相关的工具类,支持对文字/表格/图表的修改
 */
public class PptUtil<T>{


    /**
     * 替换指定页码中表格数据
     *
     * @param xmlSlideShow pptx对象
     * @param arr          待插入的数据
     * @param page         ppt页码
     * @param ignoreHeader 表头的行数
     * @param ifMerge 是否合并单元格(默认合并最后一行的第一列和第二列单元格)
     * @return
     * @throws Exception
     */
    public static XMLSlideShow insertTableData(XMLSlideShow xmlSlideShow, Object[][] arr, int page, int ignoreHeader, boolean ifMerge) throws Exception {
        List<XSLFSlide> slides = xmlSlideShow.getSlides();
        XSLFSlide slide = slides.get(page - 1);
        List<XSLFShape> shapes = slide.getShapes();

        Object[][] data = castTwoDimensionArray(arr, ignoreHeader);

        for (XSLFShape shape : shapes) {
            if (shape instanceof XSLFTable) {
                int colSize = ((XSLFTable) shape).getNumberOfColumns();//列总数
                int templateRowSize = ((XSLFTable) shape).getNumberOfRows();//模板的行总数
                int rowSize = data.length;//实际要插入行总数
                System.out.println("templateRowSize=" + templateRowSize);
                System.out.println("rowSize=" + rowSize);
//                XSLFTableRow xslfTableCells = ((XSLFTable) shape).getRows().get(rowSize - 1);
                for (int i = 0; i < rowSize; i++) {
                    if (i <= ignoreHeader - 1) {
                        continue;//忽略表头
                    } else if (i >= templateRowSize) {
                        XSLFTableRow tableCells = ((XSLFTable) shape).addRow();
                        if (ifMerge) {
                            System.out.println("merge`");
                            tableCells.mergeCells(0, 1);
                        }
                        for (int j = 0; j < colSize; j++) {
                            XSLFTableCell cell = tableCells.addCell();
                            XSLFTextParagraph textParagraph = cell.addNewTextParagraph();
                            XSLFTextRun xslfTextRun = textParagraph.addNewTextRun();
                            xslfTextRun.setText(data[i][j].toString());
                            xslfTextRun.setFontSize(12d);
                            textParagraph.setTextAlign(TextParagraph.TextAlign.CENTER);
                            xslfTextRun.setFontFamily("微软雅黑");
                        }
                        continue;
                    } else {
                        for (int j = 0; j < colSize; j++) {
                            XSLFTableCell cell = ((XSLFTable) shape).getCell(i, j);
                            if (cell != null) {
                                XSLFTextParagraph textParagraph = cell.addNewTextParagraph();
                                XSLFTextRun xslfTextRun = textParagraph.addNewTextRun();
                                xslfTextRun.setText(data[i][j].toString());
                                xslfTextRun.setFontSize(12d);
                                textParagraph.setTextAlign(TextParagraph.TextAlign.CENTER);
                                xslfTextRun.setFontFamily("微软雅黑");
                            }

                        }
                        //删掉模板中多余的数据
                        if (templateRowSize > rowSize) {
                            CTTable ctTable = ((XSLFTable) shape).getCTTable();
                            List<CTTableRow> trList = ctTable.getTrList();
                            for (int k = trList.size() - 1; k >= rowSize; k--) {
                                ctTable.removeTr(k);
                            }
                        }
                    }

                }
                XSLFTableRow xslfTableCells = ((XSLFTable) shape).getRows().get(rowSize - 1);
                if (ifMerge) {
                    xslfTableCells.mergeCells(0, 1);//默认合并最后一行的第一列和第二列数据
                }
            }
        }
        return xmlSlideShow;
    }

    /**
     * @param sourcePath         ppt模板存放路径
     * @param destPath           最终生成 ppt存放的路径
     * @param arr                待插入的数据
     * @param page               待插入的ppt页码
     * @param ignoreHeaderNumber 要忽略的表头行数
     * @return
     * @throws Exception
     */
    public static void insertTableData(String sourcePath, String destPath, Object[][] arr, int page, int ignoreHeaderNumber, boolean flag) throws Exception {
        FileInputStream inputStream = null;
        XMLSlideShow xmlSlideShow = null;
        try {
            File file = new File(sourcePath);
            inputStream = new FileInputStream(file);
            xmlSlideShow = new XMLSlideShow(inputStream);
            List<XSLFSlide> slides = xmlSlideShow.getSlides();
            XSLFSlide slide = slides.get(page - 1);
            List<XSLFShape> shapes = slide.getShapes();
            Object[][] data = castTwoDimensionArray(arr, ignoreHeaderNumber);

            for (XSLFShape shape : shapes) {
                if (shape instanceof XSLFTable) {
                    int colSize = ((XSLFTable) shape).getNumberOfColumns();//列总数
                    int templateRowSize = ((XSLFTable) shape).getNumberOfRows();//模板的行总数
                    int rowSize = data.length;//实际要插入行总数
                    System.out.println("templateRowSize=" + templateRowSize);
                    System.out.println("rowSize=" + rowSize);
                    for (int i = 0; i < rowSize; i++) {
                        if (i <= ignoreHeaderNumber - 1) {
                            continue;//忽略表头
                        } else if (i >= templateRowSize) {
                            XSLFTableRow tableCells = ((XSLFTable) shape).addRow();
                            if (flag) {
                                tableCells.mergeCells(0, 1);
                            }
                            for (int j = 0; j < colSize; j++) {
                                XSLFTableCell cell = tableCells.addCell();
                                XSLFTextParagraph textParagraph = cell.addNewTextParagraph();
                                XSLFTextRun xslfTextRun = textParagraph.addNewTextRun();
                                xslfTextRun.setText(data[i][j].toString());
                                xslfTextRun.setFontSize(12d);
                                textParagraph.setTextAlign(TextParagraph.TextAlign.CENTER);
                                xslfTextRun.setFontFamily("微软雅黑");
                                cell.setFillColor(Color.WHITE);
                            }
                            continue;
                        } else {
                            for (int j = 0; j < colSize; j++) {
                                XSLFTableCell cell = ((XSLFTable) shape).getCell(i, j);
                                if (cell != null) {

                                    XSLFTextParagraph textParagraph = cell.addNewTextParagraph();
                                    XSLFTextRun xslfTextRun = textParagraph.addNewTextRun();
                                    xslfTextRun.setText(data[i][j].toString());
                                    xslfTextRun.setFontSize(12d);
                                    textParagraph.setTextAlign(TextParagraph.TextAlign.CENTER);
                                    xslfTextRun.setFontFamily("微软雅黑");
                                    cell.setFillColor(Color.WHITE);
                                }

                            }
                            //删掉模板中多余的数据
                            if (templateRowSize > rowSize) {
                                CTTable ctTable = ((XSLFTable) shape).getCTTable();
                                List<CTTableRow> trList = ctTable.getTrList();
                                for (int k = trList.size() - 1; k >= rowSize; k--) {
                                    ctTable.removeTr(k);
                                }
                            }
                        }
                    }
                }
            }
            xmlSlideShow.write(new FileOutputStream(destPath));
        } catch (IOException e) {
            e.printStackTrace();
            throw new RuntimeException(e.getCause());
        } finally {
            if (inputStream != null) {
                inputStream.close();
            }
            if (xmlSlideShow != null) {
                xmlSlideShow.close();
            }
        }
    }


    /**
     * 替换指定页码中图片
     *
     * @param xmlSlideShow ppt操作对象
     * @param path         待替换图片路径
     * @param page         ppt页码
     * @return
     * @throws Exception
     */
    public static XMLSlideShow insertPictureData(XMLSlideShow xmlSlideShow, String path, int page) throws Exception {
        List<XSLFSlide> slides = xmlSlideShow.getSlides();
        XSLFSlide slide = slides.get(page - 1);
        List<XSLFShape> shapes = slide.getShapes();
        for (XSLFShape shape : shapes) {
            if (shape instanceof XSLFPictureShape) {
                XSLFPictureData pictureData = ((XSLFPictureShape) shape).getPictureData();
                byte[] data = image2byte(path);
                pictureData.setData(data);
            }
        }
        return xmlSlideShow;
    }


    /**
     * 图片到byte数组
     *
     * @param path
     * @return
     */
    public static byte[] image2byte(String path) {
        byte[] data = null;
        FileImageInputStream input = null;
        try {
            input = new FileImageInputStream(new File(path));
            ByteArrayOutputStream output = new ByteArrayOutputStream();
            byte[] buf = new byte[1024];
            int numBytesRead = 0;
            while ((numBytesRead = input.read(buf)) != -1) {
                output.write(buf, 0, numBytesRead);
            }
            data = output.toByteArray();
            output.close();
            input.close();
        } catch (FileNotFoundException ex1) {
            ex1.printStackTrace();
        } catch (IOException ex1) {
            ex1.printStackTrace();
        }
        return data;
    }

    /**
     * 将文件流转pptx
     *
     * @param inputStream
     * @return
     * @throws IOException
     */
    public static XMLSlideShow convertPPtx(InputStream inputStream) throws IOException {
        XMLSlideShow xmlSlideShow = new XMLSlideShow(inputStream);
        return xmlSlideShow;
    }

    public static XMLSlideShow convertPPtx(String path) throws IOException {
        FileInputStream inputStream=new FileInputStream(new File(path));
        XMLSlideShow xmlSlideShow = new XMLSlideShow(inputStream);
        return xmlSlideShow;
    }


    /**
     * 转换成特定格式二维数组便于后面遍历
     *
     * @param arr
     * @param l   表头行数
     * @return
     * @throws IOException
     */
    public static Object[][] castTwoDimensionArray(Object[][] arr, int l) throws Exception {
        if (arr == null || arr.length == 0) {
            throw new IllegalArgumentException("arr cannot be empty!");
        }
        int m = arr.length;
        int n = arr[0].length;
        Object[][] result = new Object[m + l][n];
        int k = result.length;
        for (int i = 0; i < k; i++) {
            if (i < l) {
                for (int j = 0; j < l; j++) {
                    result[i][j] = "";
                }
            } else {
                for (int j = 0; j < n; j++) {
                    result[i][j] = arr[i - l][j];
                }
            }

        }
        return result;
    }

    public static void main(String[] args) throws Exception {
        FileInputStream fileInputStream = new FileInputStream(new File("D:\\workspace_learn\\SpringCloud-Learning\\demo\\20200708.pptx"));
        Object[][] arr1 = {
                {"测试1组", "S-总账系统1", "1", "2", "3", "4", "5", "6", "7", "8", "80%"},
                {"测试2组", "P-票交所直联系统2", "1", "2", "3", "4", "5", "6", "7", "8", "85%"},
                {"测试3组", "统一支付平台3", "1", "2", "3", "4", "5", "6", "7", "8", "90%"},
                {"测试4组", "统一支付平台4", "1", "2", "3", "4", "5", "6", "7", "8", "90%"},
                {"测试5组", "统一支付平台5", "1", "2", "3", "4", "5", "6", "7", "8", "90%"},
                {"测试6组", "统一支付平台6", "1", "2", "3", "4", "5", "6", "7", "8", "90%"},
                {"测试7组", "统一支付平台7", "1", "2", "3", "4", "5", "6", "7", "8", "90%"},
                {"测试8组", "统一支付平台8", "1", "2", "3", "4", "5", "6", "7", "8", "90%"},
                {"测试9组", "统一支付平台9", "1", "2", "3", "4", "5", "6", "7", "8", "90%"},
                {"测试10组", "统一支付平台10", "1", "2", "3", "4", "5", "6", "7", "8", "90%"},

        };

        Object[][] arr2 = {
                {"1", "P-手机银行1", "1", "1", "1", "1", "1", "5"},
                {"2", "S-总账系统2", "2", "1", "1", "1", "1", "6"},
                {"3", "直销银行3", "3", "1", "1", "1", "1", "7"},
                {"4", "P-网上银行4", "4", "1", "1", "1", "1", "8"},
                {"5", "P-票交所直联系统5", "5", "1", "1", "1", "1", "9"},
                {"6", "P-票交所直联系统6", "5", "1", "1", "1", "1", "9"},
                {"7", "P-票交所直联系统7", "5", "1", "1", "1", "1", "9"},
                {"8", "P-票交所直联系统8", "5", "1", "1", "1", "1", "9"},
                {"9", "P-票交所直联系统9", "5", "1", "1", "1", "1", "9"},
                {"10", "P-票交所直联系统10", "5", "1", "1", "1", "1", "9"}
        };

        Object[][] arr3 = {
                {"未分配1", "1", "", "", "", "", "", "", "", "", "", "", "", "", "1"},
                {"ECIF系统2", "2", "", "", "", "", "", "", "", "", "", "", "", "", "2"},
                {"ECIF系统3", "3", "", "", "", "", "", "", "", "", "", "", "", "", "3"},
                {"ECIF系统4", "4", "", "", "", "", "", "", "", "", "", "", "", "", "4"},
                {"ECIF系统5", "5", "", "", "", "", "", "", "", "", "", "", "", "", "5"},
                {"ECIF系统6", "6", "", "", "", "", "", "", "", "", "", "", "", "", "6"},
                {"ECIF系统7", "7", "", "", "", "", "", "", "", "", "", "", "", "", "7"},
                {"ECIF系统8", "8", "", "", "", "", "", "", "", "", "", "", "", "", "8"}
        };
        XMLSlideShow xmlSlideShow = insertTableData(convertPPtx(fileInputStream), arr1, 2, 2, true);
        xmlSlideShow = insertTableData(xmlSlideShow, arr2, 3, 1, false);
        xmlSlideShow = insertTableData(xmlSlideShow, arr3, 4, 3, false);
        xmlSlideShow = insertPictureData(xmlSlideShow, "D:\\temp\\test.jpg", 5);
        xmlSlideShow.write(new FileOutputStream("D:\\ppt\\test2.pptx"));
    }

}

