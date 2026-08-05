package com.kuang._555.controller;

import com.kuang._555.service.MonitorConvertService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;

@RestController
@RequestMapping("/api/monitor")
public class MonitorConvertController {

    @Autowired
    private MonitorConvertService monitorConvertService;

    // 固定文件保存路径，方法内赋值规避语法报错
    private String baseTempDir;

    public MonitorConvertController() {
        this.baseTempDir = "C:\\Users\\89435\\Desktop\\第二份工作";
    }

    /**
     * 上传接口：源监测点Excel + 三维模板Excel，生成转换文件
     */
    @PostMapping("/convert3d")
    public ResponseEntity<String> convert3dExcel(
            @RequestParam("sourceFile") MultipartFile sourceFile,
            @RequestParam("templateFile") MultipartFile templateFile
    ) throws IOException {
        // 1. 创建临时存储文件夹
        File tempRoot = new File(baseTempDir);
        // 接收mkdirs返回值消除警告
        boolean createSuccess = tempRoot.mkdirs();

        // 2. 生成临时文件路径，避免重名
        String sourceTempPath = baseTempDir + "/" + System.currentTimeMillis() + "_source.xlsx";
        String templateTempPath = baseTempDir + "/" + System.currentTimeMillis() + "_template.xlsx";

        // 3. 将上传的文件写入本地临时文件
        sourceFile.transferTo(new File(sourceTempPath));
        templateFile.transferTo(new File(templateTempPath));

        try {
            // 4. 调用业务层执行转换
            String resultFilePath = monitorConvertService.convertTo3dModel(sourceTempPath, templateTempPath);
            if (resultFilePath == null) {
                return ResponseEntity.badRequest().body("转换失败：源文件无有效监测点数据");
            }
            return ResponseEntity.ok("转换完成，生成文件路径：" + resultFilePath);
        } finally {
            // 5. 删除上传临时文件
            Files.deleteIfExists(new File(sourceTempPath).toPath());
            Files.deleteIfExists(new File(templateTempPath).toPath());
        }
    }
}