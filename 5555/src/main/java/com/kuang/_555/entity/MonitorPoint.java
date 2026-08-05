package com.kuang._555.entity;

import lombok.Data;

@Data
public class MonitorPoint {
    // 源表code，对应三维表编号
    private String code;
    // 源表name，对应三维表点名称
    private String name;
    // dlwz拆分行政地址
    private String city;
    private String county;
    private String town;
    private String village;
}