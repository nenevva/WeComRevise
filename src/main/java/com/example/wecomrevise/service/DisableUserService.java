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

    public void startDisableUserService() throws IOException {

        getToken();
        checkDisabledUser();
    }

    public JSONObject requestSender(String url) {

        String response = restTemplate.getForObject(url, String.class);
        return JSON.parseObject(response);
    }

    public void disableUser(String userid) {

        String request = "{\"userid\": \"" + userid + "\",\"enable\": 0}";
        String response = restTemplate.postForObject("https://qyapi.weixin.qq.com/cgi-bin/user/update?access_token=" + access_token, request, String.class);
        JSONObject jsonResponse = JSONObject.parseObject(response);
        assert jsonResponse != null;
        System.out.println(userid + " " + jsonResponse.getString("errmsg"));
    }

    public void getToken() {

        String url = "https://qyapi.weixin.qq.com/cgi-bin/gettoken?corpid=" + corpid + "&corpsecret=" + corpsecret;
        JSONObject response = requestSender(url);
        access_token = response.getString("access_token");

        if (response.getInteger("errcode").equals(0)) {
            System.out.println("Succeed getting token.");
        } else {
            System.out.println("[ERROR] Failed when getting token: " + response.getString("\"errmsg\""));
        }

    }

    public void createWorkbook(String path, ArrayList<Row> rowArrayList) throws IOException {

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Result");
        sheet.copyRows(rowArrayList, 0, new CellCopyPolicy());
        FileOutputStream fos = new FileOutputStream(path);
        workbook.write(fos);
        fos.close();
        System.out.println("New workbook created.");
    }

    public void checkDisabledUser() throws IOException {

        System.out.println("checkDisabledUser()");
        System.out.println(filePath);

        Path path = Paths.get(filePath);
        String resultPath = path.getParent().toString() + "\\result.xlsx";
        System.out.println("Output file path: " + resultPath);

        InputStream is = Files.newInputStream(path);

        Workbook workbook = null;

        if (filePath.endsWith("xls")) {
            workbook = new HSSFWorkbook(is);
        } else if (filePath.endsWith("xlsx")) {
            workbook = new XSSFWorkbook(is);
        } else {
            System.out.println("[ERROR] checkDisabledUser(): Unsupported file format.");
            System.exit(1);
        }

        Sheet sheet = workbook.getSheetAt(0);
        Row row;
        ArrayList<Row> rowArrayList = new ArrayList<>();
        rowArrayList.add(sheet.getRow(0));
        int rowNum = 1;
        for (int i = rowStartPos; i <= sheet.getLastRowNum(); i++) {
            row = sheet.getRow(i);
            if (row.getCell(columnPos).getNumericCellValue() == 1) {

                String userid = row.getCell(columnUserNamePos).getStringCellValue();
                String url = "https://qyapi.weixin.qq.com/cgi-bin/user/get?access_token=" + access_token + "&userid=" + userid;
                JSONObject response = requestSender(url);

                if (response.getInteger("errcode").equals(0) && response.getInteger("enable") == 1) {
                    row.setRowNum(rowNum++);
                    rowArrayList.add(row);
                    if (disableUser) {
                        disableUser(userid);
                    }
                }
            }
        }
        if (output) {
            createWorkbook(resultPath, rowArrayList);
        }
    }
}
