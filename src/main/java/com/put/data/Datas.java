package com.put.data;

import lombok.Data;

/**
 * TODO
 *
 * @author zhujiqian
 * @date 2020/10/27 20:09
 */
@Data
public class Datas {
    //名字
    private String name;
    //保证金
    private String margin;
    //可用资金
    private String avaFunds;

    public Datas(String name, String margin, String avaFunds) {
        this.name = name;
        this.margin = margin;
        this.avaFunds = avaFunds;
    }

}
