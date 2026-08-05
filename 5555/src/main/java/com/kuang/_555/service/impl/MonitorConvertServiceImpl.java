package com.kuang._555.service.impl;

import com.kuang._555.entity.MonitorPoint;
import com.kuang._555.service.MonitorConvertService;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Pattern;

@Service
public class MonitorConvertServiceImpl implements MonitorConvertService {
    private static final Logger log = LoggerFactory.getLogger(MonitorConvertServiceImpl.class);
    // 清洗空白、全角空格、换行、特殊符号正则
    private static final Pattern CLEAN_PATTERN = Pattern.compile("[\\s\\u3000\\u200b\\r\\n\\t\\v＿－]+");

    // 写死目标桌面文件夹，不再读取配置文件
    private final String baseTempDir = "C:\\Users\\89435\\Desktop\\第二份工作";

    @Override
    public String convertTo3dModel(String sourceFilePath, String templateFilePath) throws IOException {
        // 读取云申监测源Excel，封装实体
        List<MonitorPoint> pointList = readSourceExcel(sourceFilePath);
        log.info("读取云申监测点源文件，有效数据共{}条", pointList.size());

        // 无有效数据直接返回
        if (pointList.isEmpty()) {
            log.error("源Excel不存在有效监测点数据");
            return null;
        }

        // 填充三维模板生成新Excel
        String resultFilePath = write3dTargetExcel(pointList, templateFilePath);
        log.info("三维模型文件转换完成，输出路径：{}", resultFilePath);
        return resultFilePath;
    }

    /**
     * 读取云申监测点源Excel（A=code B=dlwz C=name）
     */
    private List<MonitorPoint> readSourceExcel(String filePath) throws IOException {
        List<MonitorPoint> dataList = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = WorkbookFactory.create(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            // 第0行为表头，数据从第1行开始遍历
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                String rawCode = getCellText(row.getCell(0));
                String rawDlwz = getCellText(row.getCell(1));
                String rawName = getCellText(row.getCell(2));

                // 清洗文本
                String cleanCode = cleanStr(rawCode);
                String cleanAddr = cleanStr(rawDlwz);
                String cleanName = cleanStr(rawName);

                // 编号为空直接跳过该行无效数据
                if (cleanCode.isBlank()) continue;

                // 拆分四级行政地址
                MonitorPoint point = splitAddress(cleanAddr, cleanCode, cleanName);
                dataList.add(point);
            }
        }
        return dataList;
    }

    /**
     * 【修复版地址拆分】优先匹配长行政词组，解决街道办事处、镇人民政府截断问题
     */
    private MonitorPoint splitAddress(String addr, String code, String name) {
        MonitorPoint point = new MonitorPoint();
        point.setCode(code);
        point.setName(name);

        // 移除湖北省前缀
        String tempAddr = addr.replace("湖北省", "").trim();
        String city = "十堰市";
        String county = "";
        String town = "";
        String village = "";

        // 1、提取市
        int cityIndex = tempAddr.indexOf("市");
        if (cityIndex != -1) {
            city = tempAddr.substring(0, cityIndex + 1);
            tempAddr = tempAddr.substring(cityIndex + 1).trim();
        }

        // 2、提取区/县
        int xianIdx = tempAddr.indexOf("县");
        int quIdx = tempAddr.indexOf("区");
        int countyIdx = -1;
        if (xianIdx != -1 && quIdx != -1) {
            countyIdx = Math.min(xianIdx, quIdx);
        } else if (xianIdx != -1) {
            countyIdx = xianIdx;
        } else if (quIdx != -1) {
            countyIdx = quIdx;
        }
        if (countyIdx != -1) {
            county = tempAddr.substring(0, countyIdx + 1);
            tempAddr = tempAddr.substring(countyIdx + 1).trim();
        }

        // 3、乡镇关键词：长词组放最前面，优先完整匹配不截断
        String[] townKeys = {
                "街道办事处", "镇人民政府", "保护区管理局",
                "开发区", "街办", "街道", "镇", "乡"
        };
        int minIndex = -1;
        String hitKey = "";
        for (String key : townKeys) {
            int idx = tempAddr.indexOf(key);
            if (idx == -1) continue;
            // 同位置匹配，选择更长词组；不同位置取靠前
            if (minIndex == -1 || idx < minIndex || (idx == minIndex && key.length() > hitKey.length())) {
                minIndex = idx;
                hitKey = key;
            }
        }
        if (minIndex != -1) {
            town = tempAddr.substring(0, minIndex + hitKey.length());
            tempAddr = tempAddr.substring(minIndex + hitKey.length()).trim();
        }

        // 4、区分村/社区截取长度（社区2个字，村1个字）
        int villageIdx = tempAddr.indexOf("村");
        int communityIdx = tempAddr.indexOf("社区");
        int vilCutIndex = -1;
        int cutLength = 1;
        if (villageIdx != -1 && communityIdx != -1) {
            vilCutIndex = Math.min(villageIdx, communityIdx);
            cutLength = vilCutIndex == communityIdx ? 2 : 1;
        } else if (communityIdx != -1) {
            vilCutIndex = communityIdx;
            cutLength = 2;
        } else if (villageIdx != -1) {
            vilCutIndex = villageIdx;
            cutLength = 1;
        }
        if (vilCutIndex != -1) {
            village = tempAddr.substring(0, vilCutIndex + cutLength);
        }

        point.setCity(city);
        point.setCounty(county);
        point.setTown(town);
        point.setVillage(village);
        return point;
    }

    /**
     * 读取三维模板，填充转换后数据生成新Excel
     */
    private String write3dTargetExcel(List<MonitorPoint> pointList, String templatePath) throws IOException {
        // 修改文件名前缀为：新三维模型监测点，带时间戳防止覆盖旧文件
        String outFileName = "新三维模型监测点_" + System.currentTimeMillis() + ".xlsx";
        String outFullPath = baseTempDir + "/" + outFileName;

        try (FileInputStream fis = new FileInputStream(templatePath);
             Workbook wb = WorkbookFactory.create(fis);
             FileOutputStream fos = new FileOutputStream(outFullPath)) {
            Sheet sheet = wb.getSheetAt(0);
            int writeRowNum = 1; // 表头第0行，数据从第二行开始写入
            for (MonitorPoint p : pointList) {
                Row dataRow = sheet.createRow(writeRowNum++);
                // 0 项目固定值
                dataRow.createCell(0).setCellValue("十堰技术支撑平台三维一张图");
                // 1 市
                dataRow.createCell(1).setCellValue(p.getCity());
                // 2 县/区
                dataRow.createCell(2).setCellValue(p.getCounty());
                // 3 乡镇街道
                dataRow.createCell(3).setCellValue(p.getTown());
                // 4 村/社区
                dataRow.createCell(4).setCellValue(p.getVillage());
                // 5 类别固定专业监测
                dataRow.createCell(5).setCellValue("专业监测");
                // 6 编号code
                dataRow.createCell(6).setCellValue(p.getCode());
                // 7 点名称name
                dataRow.createCell(7).setCellValue(p.getName());
                // 8列往后无需赋值，自动空白
            }
            wb.write(fos);
        }
        return outFullPath;
    }

    /**
     * 清洗文本，去除所有空白、特殊符号
     */
    private String cleanStr(String rawText) {
        if (rawText == null) return "";
        return CLEAN_PATTERN.matcher(rawText).replaceAll("");
    }

    /**
     * 通用单元格读取工具，兼容文本、数字、公式
     */
    private String getCellText(Cell cell) {
        if (cell == null) return "";
        CellType cellType = cell.getCellType();
        if (cellType == CellType.FORMULA) {
            cellType = cell.getCachedFormulaResultType();
        }
        return switch (cellType) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> String.valueOf(cell.getNumericCellValue());
            default -> "";
        };
    }
}