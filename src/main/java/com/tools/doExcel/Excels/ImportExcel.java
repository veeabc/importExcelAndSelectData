package com.tools.doExcel.Excels;

import com.tools.doExcel.Model.CloudServer;
import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.*;

public class ImportExcel {

    public static void importData(String fileName, MultipartFile mfile, Map map) {
        File uploadDir = new File("src/main/java/com/test/files/");
        //创建一个目录 （它的路径名由当前 File 对象指定，包括任一必须的父路径。）
        if (!uploadDir.exists()) uploadDir.mkdirs();
        //新建一个文件
        File tempFile = new File("src/main/java/com/test/files/" + new Date().getTime() + ".xlsx");
        if (!tempFile.exists()) {
            try {
                tempFile.createNewFile();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        //初始化输入流
        InputStream is = null;
        try {
            //将上传的文件写入新建的文件中
            // 在spring boot 内嵌的tomcat中不work
            //mfile.transferTo(tempFile);

            FileUtils.copyInputStreamToFile(mfile.getInputStream(), tempFile);

            //根据新建的文件实例化输入流
            is = new FileInputStream(tempFile);

            //根据版本选择创建Workbook的方式
            Workbook wb = null;
            //根据文件名判断文件是2003版本还是2007版本
            if (ImportExcelUtils.isExcel2007(fileName)) {
                wb = new XSSFWorkbook(is);
            } else {
                wb = new HSSFWorkbook(is);
            }
            //根据excel里面的内容读取知识库信息
            readExcelValue(wb, tempFile, map);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (is != null) {
                try {
                    is.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    private static void getRandomNameList(List<CloudServer> cloudServerList, Map map){
        /*不同组别的人员放入对应组别的List*/
        List<CloudServer> cloudServerListA = new ArrayList<CloudServer>();
        List<CloudServer> cloudServerListB = new ArrayList<CloudServer>();
        List<CloudServer> cloudServerListC = new ArrayList<CloudServer>();

        List<CloudServer> resultCloudServerList = new ArrayList<CloudServer>();


        CloudServer cloudServerA;
        CloudServer cloudServerB;
        CloudServer cloudServerC;

        String departmentA;
        String departmentB;
        String departmentC;

        String employeeNameA;
        String employeeNameB;
        String employeeNameC;

        /*分组*/
        CloudServer cloudServer;
        for(int i = 0; i < cloudServerList.size(); i++){
            cloudServer = cloudServerList.get(i);
//            System.out.println("第" + i + "次分组");
            switch(cloudServer.getGroupId()){
                case "A":
                    cloudServerListA.add(cloudServer);
                    break;
                case "B":
                    cloudServerListB.add(cloudServer);
                    break;
                case "C":
                    cloudServerListC.add(cloudServer);
                    break;
                default:
                    break;
            }
        }

        System.out.println("A组共" + cloudServerListA.size() + "人"
                +"；B组共" + cloudServerListB.size() + "人"
                + "；C组共" + cloudServerListC.size() + "人");

        /*随机从A、B、C三个列表中随机取出一组数据进行比较*/
        int calNum = 10000;/*定义变量，决定每次提交最多运算N次，避免程序死循环*/
        int j = 0;

        while(j < calNum){

            /*先对三个队列的数据进行随机排序*/
            Collections.shuffle(cloudServerListA);
            cloudServerA = cloudServerListA.get(0);

            Collections.shuffle(cloudServerListB);
            cloudServerB = cloudServerListB.get(0);

            Collections.shuffle(cloudServerListC);
            cloudServerC = cloudServerListC.get(0);

            employeeNameA = cloudServerA.getEmployeeName();
            departmentA = cloudServerA.getDepartment();
            employeeNameB = cloudServerB.getEmployeeName();
            departmentB = cloudServerB.getDepartment();
            employeeNameC = cloudServerC.getEmployeeName();
            departmentC = cloudServerC.getDepartment();

            String result = "第1个人员：[" + employeeNameA + "]部门[" + departmentA
                    + "]；第2个人员：["+ employeeNameB + "]部门[" + departmentB
                    + "]；第3个人员：["+ employeeNameC + "]部门[" + departmentC + "]";


            if(departmentA.equals(departmentB) || departmentB.equals(departmentC) || departmentC.equals(departmentA)){
                System.out.println("第"+ (j + 1) +"次选择，" + result + "不满足条件！");
                j += 1;
                continue;
            }else{
                /*找到后退出循环*/
                System.out.println("第"+ (j + 1) +"次选择，" + result + "满足条件！");
//                map.put("Result", result);
                resultCloudServerList.add(cloudServerA);
                resultCloudServerList.add(cloudServerB);
                resultCloudServerList.add(cloudServerC);
                map.put("resultCloudServerList", resultCloudServerList );
                break;
            }

        }

    }


    private static void readExcelValue(Workbook wb, File tempFile, Map map) {
        //错误信息接收器
        String errorMsg = "";
        //得到第一个shell
        Sheet sheet = wb.getSheetAt(0);
        //得到Excel的行数
        int totalRows = sheet.getPhysicalNumberOfRows();
        //总列数
        int totalCells = 0;
        //得到Excel的列数(前提是有行数)，从第二行算起
        if (totalRows >= 2 && sheet.getRow(1) != null) {
            totalCells = sheet.getRow(1).getPhysicalNumberOfCells();
        }
        /*不分组别的数据存储*/
        List<CloudServer> cloudServerList = new ArrayList<CloudServer>();

        CloudServer cloudServer;

        String br = "\n";
        Row rowHeader = sheet.getRow(0);

        //循环Excel行数,从第二行开始。标题不入库
        for (int r = 1; r < totalRows; r++) {
            String rowMessage = "";
            Row row = sheet.getRow(r);
            if (row == null) {
                errorMsg += br + "第" + (r + 1) + "行数据有问题，请仔细检查！";
                continue;
            }
            cloudServer = new CloudServer();
            //循环Excel的列
            //采用反射的方式判断每列数据并读取数据
            Class<? extends CloudServer> aClass = cloudServer.getClass();
            Field[] declaredFields = aClass.getDeclaredFields();
            for (int c = 0; c < totalCells; c++) {
                Cell cell = row.getCell(c);
                if (null != cell) {
                    Object cellValue = ExcelUtils.getCellValue(cell);
                    if (cellValue == null || "".equals(cellValue.toString())) {
                        rowMessage += "第" + (c + 1) + "列数据有问题，请仔细检查；";
                    } else {
                        Cell cellHeader = rowHeader.getCell(c);
                        System.out.println((r + 1) + "行" + (c + 1) + "列 ==cellHeader==" + cellHeader + "==cellValue==" + cellValue);
                        for (Field f : declaredFields) {
                            f.setAccessible(true);
                            if (f.getName().equals(String.valueOf(ExcelUtils.getCellValue(cellHeader)))) {
/*
                                if ("ip".equals(f.getName()) && !(StringUtils.isboolIp(cellValue.toString()))) {
                                    rowMessage += "第" + (c + 1) + "列数据不是有效的IP，请仔细检查；";
*/
                                if ("groupId".equals(f.getName()) && !(StringUtils.isABC(cellValue.toString()))) {
                                    rowMessage += "第" + (c + 1) + "列数据不是有效的组别，请仔细检查；";
                                    break;
                                }
                                try {
                                    switch (f.getGenericType().toString()) {
                                        case "class java.lang.String":
                                            f.set(cloudServer, String.valueOf(cellValue));
                                            break;
                                        case "int":
                                            f.set(cloudServer, Integer.valueOf(cellValue.toString()));
                                            break;
                                        default:
                                            break;
                                    }
                                } catch (IllegalAccessException e) {
                                    e.printStackTrace();
                                }
                                break;
                            }
                        }
                    }
                } else {
                    rowMessage += "第" + (c + 1) + "列数据有问题，请仔细检查；";
                }
            }
            //拼接每行的错误提示
            if (!StringUtils.isEmpty(rowMessage)) {
                errorMsg += br + "第" + (r + 1) + "行，" + rowMessage;
                break;
            } else {
                cloudServerList.add(cloudServer);
/*
                switch(cloudServer.getGroupId()){
                    case "A":
                        cloudServerListA.add(cloudServer);
                        break;
                    case "B":
                        cloudServerListB.add(cloudServer);
                        break;
                    case "C":
                        cloudServerListC.add(cloudServer);
                        break;
                    default:
                        break;
                }
*/
            }
        }

        //删除上传的临时文件
        if (tempFile.exists()) {
            tempFile.delete();
        }

        //全部验证通过才导入到数据库,同时进行数据筛选
        if (StringUtils.isEmpty(errorMsg)) {
            for (CloudServer server : cloudServerList) {
                //TODO 插入到数据库
                System.out.println(server);
            }
            errorMsg = "导入成功，共" + cloudServerList.size() + "条数据！";

            /*找到对应条件的人员*/
            getRandomNameList(cloudServerList, map);
        }


        map.put("cloudServerList", cloudServerList);
        map.put("errorMsg", errorMsg);
    }

}
