package com.put;

import com.put.data.Datas;
import com.put.put.ExportExcleClient;
import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * TODO
 *
 * @author zhujiqian
 * @date 2020/10/27 20:08
 */
public class PutExcel {

    public PutExcel() {
    }


    public static List<String> readTxtFile1(String TxtName, String filePath) {
        ArrayList list = new ArrayList();

        try {

            System.out.println("该组合：" + TxtName);
            File file = new File(filePath);
            if (file.isFile() && file.exists()) {
                InputStreamReader read = new InputStreamReader(new FileInputStream(file), "GBK");
                BufferedReader bufferedReader = new BufferedReader(read);
                String lineTxt = null;

                while(true) {
                    do {
                        if ((lineTxt = bufferedReader.readLine()) == null) {
                            read.close();
                            return list;
                        }
                    } while(!lineTxt.contains("持仓保证金:") && !lineTxt.contains("保证金占用:") && !lineTxt.contains("保证金占用 Margin Occupied：") && !lineTxt.contains("保证金占用 Margin Occupied:"));

                    String ZiJin;
                    int a;
                    String b;
                    int c;
                    String d;
                    int e;
                    int f;
                    String E;
                    if (lineTxt.contains("持仓保证金:")) {
                        ZiJin = lineTxt.replace(" ", "");
                        a = ZiJin.indexOf("持");
                        b = ZiJin.substring(a);
                        c = b.indexOf(".");
                        d = b.substring(0, c + 3);
                        e = d.indexOf(":");
                        f = d.length();
                        E = d.substring(e + 1, f);
                        list.add(E);
                    } else if (lineTxt.contains("保证金占用:")) {
                        ZiJin = lineTxt.replace(" ", "");
                        a = ZiJin.indexOf("保");
                        b = ZiJin.substring(a);
                        c = b.indexOf(".");
                        d = b.substring(0, c + 3);
                        e = d.indexOf(":");
                        f = d.length();
                        E = d.substring(e + 1, f);
                        list.add(E);
                    } else if (lineTxt.contains("保证金占用 Margin Occupied：")) {
                        ZiJin = lineTxt.replace(" ", "");
                        a = ZiJin.indexOf("保");
                        b = ZiJin.substring(a);
                        c = b.indexOf(".");
                        d = b.substring(0, c + 3);
                        e = d.indexOf("：");
                        f = d.length();
                        E = d.substring(e + 1, f);
                        list.add(E);
                    } else if (lineTxt.contains("保证金占用 Margin Occupied:")) {
                        ZiJin = lineTxt.replace(" ", "");
                        a = ZiJin.indexOf("保");
                        b = ZiJin.substring(a);
                        c = b.indexOf(".");
                        d = b.substring(0, c + 3);
                        e = d.indexOf(":");
                        f = d.length();
                        E = d.substring(e + 1, f);
                        list.add(E);
                    }
                }
            } else {
                System.out.println("找不到指定的文件");
            }
        } catch (Exception var16) {
            System.out.println("读取文件内容出错");
            var16.printStackTrace();
        }

        return list;
    }

    public static List<String> readTxtFile2(String TxtName, String filePath) {
        ArrayList list = new ArrayList();

        try {

            System.out.println("该组合：" + TxtName);
            File file = new File(filePath);
            if (file.isFile() && file.exists()) {
                InputStreamReader read = new InputStreamReader(new FileInputStream(file), "GBK");
                BufferedReader bufferedReader = new BufferedReader(read);
                String lineTxt = null;

                while(true) {
                    do {
                        if ((lineTxt = bufferedReader.readLine()) == null) {
                            read.close();
                            return list;
                        }
                    } while(!lineTxt.contains("可用资金:") && !lineTxt.contains("可用资金 Fund Avail.：") && !lineTxt.contains("可用资金 Fund Avail.:"));

                    String ZiJin;
                    int a;
                    String b;
                    int c;
                    String d;
                    int e;
                    int f;
                    String E;
                    if (lineTxt.contains("可用资金:")) {
                        ZiJin = lineTxt.replace(" ", "");
                        ZiJin = lineTxt.replace(" ", "");
                        a = ZiJin.indexOf("可");
                        b = ZiJin.substring(a);
                        c = b.indexOf(".");
                        d = b.substring(0, c + 3);
                        e = d.indexOf(":");
                        f = d.length();
                        E = d.substring(e + 1, f);
                        list.add(E);
                    } else if (lineTxt.contains("可用资金 Fund Avail.：")) {
                        ZiJin = lineTxt.replace(" ", "");
                        ZiJin = lineTxt.replace(" ", "");
                        a = ZiJin.indexOf("可");
                        b = ZiJin.substring(a);
                        c = b.lastIndexOf(".");
                        d = b.substring(0, c + 3);
                        e = d.indexOf("：");
                        f = d.length();
                        E = d.substring(e + 1, f);
                        list.add(E);
                    } else if (lineTxt.contains("可用资金 Fund Avail.:")) {
                        lineTxt.replace(" ", "");
                        ZiJin = lineTxt.replace(" ", "");
                        a = ZiJin.indexOf("可");
                        b = ZiJin.substring(a);
                        c = b.lastIndexOf(".");
                        d = b.substring(0, c + 3);
                        e = d.indexOf(":");
                        f = d.length();
                        E = d.substring(e + 1, f);
                        list.add(E);
                    }
                }
            } else {
                System.out.println("找不到指定的文件");
            }
        } catch (Exception var16) {
            System.out.println("读取文件内容出错");
            var16.printStackTrace();
        }

        return list;
    }

    public static void main(String[] argv) {
        String path = "D:\\批量处理对账单\\对账单\\对账单批量名字集合.txt";
        List<String> nums = writeToDat(path);
        List<Datas> listData = new ArrayList();

        for(int i = 0; i < nums.size(); ++i) {
            if (!(nums.get(i)).equals("对账单批量名字集合.txt")) {
                listData.add(ZuHe(nums.get(i)));
                System.out.println("--------==========" + ZuHe(nums.get(i)));
            }
        }

        System.out.println("-----------" + listData);
        ExportExcleClient client = new ExportExcleClient();
        client.alldata(listData);
        String url = client.exportExcel();
        System.out.println(url);
    }

    public static Datas ZuHe(String TxtName) {
        String address = "D:\\批量处理对账单\\对账单\\" + TxtName;
        List<String> r1 = readTxtFile1(TxtName, address);
        List<String> r2 = readTxtFile2(TxtName, address);
        int c = TxtName.indexOf(".");
        String txt = TxtName.substring(0, c);
        System.out.println(txt);
        System.out.println(r1);
        Datas d = null;
        if (r1.isEmpty() && r2.isEmpty()) {
            if (r1.isEmpty() && r2.isEmpty()) {
                System.out.println(txt + "--" + r1 + "---" + r2);
                d = new Datas(txt, "无", "无");
            }
        } else {
            d = new Datas(txt, r1.get(0), r2.get(0));
        }

        return d;
    }

    public static List<String> writeToDat(String path) {
        File file = new File(path);
        ArrayList list = new ArrayList();

        try {
            String encoding = "GBK";
            InputStreamReader read = new InputStreamReader(new FileInputStream(file), encoding);
            BufferedReader bw = new BufferedReader(read);
            String line = null;

            while((line = bw.readLine()) != null) {
                list.add(line);
            }

            bw.close();
        } catch (IOException var7) {
            var7.printStackTrace();
        }

        return list;
    }


}
