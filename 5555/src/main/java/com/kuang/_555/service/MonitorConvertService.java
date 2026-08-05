package com.kuang._555.service;

import java.io.IOException;

public interface MonitorConvertService {
    /**
     * 云申监测Excel转三维模型标准Excel
     * @param sourceFilePath 云申监测源文件路径
     * @param templateFilePath 三维模型模板文件路径
     * @return 生成后的Excel完整路径
     * @throws IOException 文件读写异常
     */
    String convertTo3dModel(String sourceFilePath, String templateFilePath) throws IOException;
}