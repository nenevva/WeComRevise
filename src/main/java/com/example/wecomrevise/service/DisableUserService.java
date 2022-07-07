package com.example.wecomrevise.service;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONObject;
import lombok.RequiredArgsConstructor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.springframework.web.client.RestTemplate;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;

@Service
@RequiredArgsConstructor
public class DisableUserService {

    private final RestTemplate restTemplate;

    @Value("${columnPos}")
    private int columnPos;

    @Value("${columnUserNamePos}")
    private int columnUserNamePos;

    @Value("${rowStartPos}")
    private int rowStartPos;

    @Value("${filePath}")
    private String filePath;

    @Value("${corpid}")
    private String corpid;

    @Value("${corpsecret}")
    private String corpsecret;

    @Value("${disableUser}")
    private boolean disableUser;

    @Value("${output}")
    private boolean output;

    private String access_token;

    //服务起点
    public void startDisableUserService() throws IOException {

        getToken();
        checkDisabledUser();
    }

    //发送请求
    public JSONObject requestSender(String url) {

        String response = restTemplate.getForObject(url, String.class);
        return JSON.parseObject(response);
    }

    //禁用用户
    public void disableUser(String userid) {

        String request = "{\"userid\": \"" + userid + "\",\"enable\": 0}";
        String response = restTemplate.postForObject("https://qyapi.weixin.qq.com/cgi-bin/user/update?access_token=" + access_token, request, String.class);
        JSONObject jsonResponse = JSONObject.parseObject(response);
        assert jsonResponse != null;
        System.out.println(userid + " " + jsonResponse.getString("errmsg"));
    }

    //获取access_token
    public void getToken() {

        String url = "https://qyapi.weixin.qq.com/cgi-bin/gettoken?corpid=" + corpid + "&corpsecret=" + corpsecret;
        JSONObject response = requestSender(url);
        access_token = response.getString("access_token");

        if (response.getInteger("errcode").equals(0)) {
            System.out.println("Succeed getting token.");
        } else {
            System.out.println("[ERROR] Failed when getting token: " + response.getString("\"errmsg\""));
            System.exit(1);
        }

    }

    //创建新表格
    public void createWorkbook(String path, ArrayList<Row> rowArrayList) throws IOException {

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Result");
        sheet.copyRows(rowArrayList, 0, new CellCopyPolicy());
        FileOutputStream fos = new FileOutputStream(path);
        workbook.write(fos);
        fos.close();
        System.out.println("New workbook created.");
    }

    //读取表格，检查在rtx中被禁用的用户是否存在于企业微信中
    public void checkDisabledUser() throws IOException {

        //设置输入输出文件路径
        Path path = Paths.get(filePath);
        String resultPath = path.getParent().toString() + "\\result.xlsx";
        InputStream is = Files.newInputStream(path);

        System.out.println("Source file path: " + filePath);
        System.out.println("Output file path: " + resultPath);

        //读取文件，获取表格内容
        Workbook workbook = null;

        if (filePath.endsWith("xls")) {
            workbook = new HSSFWorkbook(is);
        } else if (filePath.endsWith("xlsx")) {
            workbook = new XSSFWorkbook(is);
        } else {
            System.out.println("[ERROR] Unsupported source file format.");
            System.exit(1);
        }

        //逐行读取
        Sheet sheet = workbook.getSheetAt(0);
        Row row;
        ArrayList<Row> rowArrayList = new ArrayList<>();
        rowArrayList.add(sheet.getRow(0));
        int rowNum = 1;
        for (int i = rowStartPos; i <= sheet.getLastRowNum(); i++) {
            row = sheet.getRow(i);
            if (row.getCell(columnPos).getNumericCellValue() == 1) {
                //此用户在rtx中被禁用
                //发送请求，检查此用户是否活跃于企业微信中
                String userid = row.getCell(columnUserNamePos).getStringCellValue();
                String url = "https://qyapi.weixin.qq.com/cgi-bin/user/get?access_token=" + access_token + "&userid=" + userid;
                JSONObject response = requestSender(url);

                //如果用户存在于企业微信中，则禁用此用户
                if (response.getInteger("errcode").equals(0) && response.getInteger("enable") == 1) {
                    row.setRowNum(rowNum++);
                    rowArrayList.add(row);
                    if (disableUser) {
                        disableUser(userid);
                    }
                }
            }
        }
        //创建新表格
        if (output) {
            createWorkbook(resultPath, rowArrayList);
        }
    }
}
