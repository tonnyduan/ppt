package com.example.demo;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@RestController
public class TestController {

    @Autowired
    private Config config;

    @GetMapping("/config")
    public String getConfig() throws IOException {
        XMLSlideShow xmlSlideShow = null;
        FileInputStream inputStream = null;
        try {
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

            Object[][] arr4 = {
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

            Map<String,Object[][]> dataMap=new HashMap<>();
            dataMap.put("2",arr1);
            dataMap.put("3",arr2);
            dataMap.put("4",arr3);
            dataMap.put("7",arr4);

            HashMap<String, String> pathMap = config.getPath();
            List<Map<String, String>> pagesList = config.getPages();
            String templatePath = pathMap.get("template-path");//ppt模板路径
            String newFilepath = pathMap.get("newfile-path");//生成的ppt路径
            File file=new File(templatePath);
            inputStream=new FileInputStream(file);
            xmlSlideShow = PptUtil.convertPPtx(inputStream);
            for (Map<String, String> pagesMap : pagesList) {
                int index = Integer.parseInt(pagesMap.get("index"));//ppt的页码
//                int data = Integer.parseInt(pagesMap.get("data"));
                String type = pagesMap.get("type");
                if ("table".equals(type)) {
                    int headerNumber = Integer.parseInt(pagesMap.get("header-number"));//表头数
                    boolean ifMerge = Boolean.valueOf(pagesMap.get("if-merge"));//是否合并
                    xmlSlideShow = PptUtil.insertTableData(xmlSlideShow, dataMap.get(index+""), index, headerNumber, ifMerge);
                } else if ("picture".equals(type)) {
                    String picPath = pagesMap.get("pic-path");//图片路径
                    xmlSlideShow = PptUtil.insertPictureData(xmlSlideShow, picPath, index);
                }

            }
            xmlSlideShow.write(new FileOutputStream(newFilepath));
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if(inputStream!=null){
                inputStream.close();
            }
            if(xmlSlideShow!=null){
                xmlSlideShow.close();
            }
        }
        return "ok";

    }
}
