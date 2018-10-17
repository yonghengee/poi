package com.yqh.bbct;

import com.yqh.bbct.configuration.ResourceRealPath;
import com.yqh.bbct.entity.OrdersGoods;
import com.yqh.bbct.excel.ExcelUtils;
import com.yqh.bbct.excel.model.GoodsData;
import com.yqh.bbct.excel.model.OrderData;
import com.yqh.bbct.excel.model.OrdersGoodsData;
import com.yqh.bbct.excel.model.SheetData;
import com.yqh.bbct.utils.MapUtils;
import com.yqh.bbct.utils.UUIDUtil;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by wyh in 2018/9/14 14:51
 **/
public class DoExcel {
    public static void main(String[] args) {

        List<GoodsData> goodsData = new ArrayList<>();
        goodsData.add(new GoodsData(1, "2", "3"));
//        doGoodsData(goodsData);
    }




    public static String doOrderData(OrderData orderData, List<OrdersGoodsData> list) throws Exception {
        //获取模板
        String model = ResourceRealPath.resourceDisk+"/model/orderModel.xls";
        String fileName = UUIDUtil.generateUUID().toString() + ".xls";
        File f = new File(ResourceRealPath.excel + "/order/" + fileName);

        SheetData sds = new SheetData("报关单");

        System.out.println(orderData.getTransportTypeId());
        Map<String, Object> map = MapUtils.convertToMap(orderData);
        String transportTypeId = map.get("transportTypeId").toString();
        System.err.println(transportTypeId);
        sds.setMap(map);
        List<Map<String, Object>> hashMaps = MapUtils.convertToListMap(list);

        sds.setDatas(hashMaps);
        try {
            ExcelUtils.writeData(model, new FileOutputStream(f), sds);
        } catch (FileNotFoundException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        return "/excel/order/" + fileName;
    }

    public static String doInvoiceData(Map<String,Object> o,List<Map<String,Object>> list) {
        //获取模板
        String model = ResourceRealPath.resourceDisk+"/model/invoiceModel.xls";
        String fileName = UUIDUtil.generateUUID().toString() + ".xls";
        File f = new File(ResourceRealPath.excel + "/invoice/" + fileName);


        SheetData sds = new SheetData("发票");

        sds.setMap(o);
        sds.setDatas(list);
        try {
            ExcelUtils.writeData(model, new FileOutputStream(f), sds);
        } catch (FileNotFoundException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        return "/excel/invoice/" + fileName;
    }

    public static String doPackingListData(Map<String,Object> o,List<Map<String,Object>> list) {
        //获取模板
        String model = ResourceRealPath.resourceDisk+"/model/packingListModel.xls";
        String fileName = UUIDUtil.generateUUID().toString() + ".xls";
        File f = new File(ResourceRealPath.excel + "/packingList/" + fileName);

        SheetData sds = new SheetData("装箱单");

        sds.setMap(o);
        sds.setDatas(list);

        try {
            ExcelUtils.writeData(model, new FileOutputStream(f), sds);
        } catch (FileNotFoundException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        return "/excel/packingList/" + fileName;
    }

    public static String doSalesContractData(Map<String,Object> o,List<Map<String,Object>> list) {
        //获取模板
        String model = ResourceRealPath.resourceDisk+"/model/salesContractModel.xls";
        String fileName = UUIDUtil.generateUUID().toString() + ".xls";
        File f = new File(ResourceRealPath.excel + "/saleContract/" + fileName);

        SheetData sds = new SheetData("货物申报要素表");

        sds.setMap(o);
        sds.setDatas(list);
        try {
            ExcelUtils.writeData(model, new FileOutputStream(f), sds);
        } catch (FileNotFoundException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        return "/excel/saleContract/" + fileName;
    }

    public static String doGoodsData(List<GoodsData> goodsDataList)throws Exception {
        //获取模板
        String model = ResourceRealPath.resourceDisk+"/model/goodsModel.xls";
        String fileName = UUIDUtil.generateUUID().toString() + ".xls";
        File f = new File(ResourceRealPath.excel + "/goods/" + fileName);


        SheetData sds = new SheetData("货物申报要素表");

        List<Map<String, Object>> list = MapUtils.convertToListMap(goodsDataList);
        sds.addDatas(list);
        try {
            ExcelUtils.writeData(model, new FileOutputStream(f), sds);
        } catch (FileNotFoundException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        return "/excel/goods/" + fileName;
    }

    public static String doWithDrawData(List<Map<String, Object>> goodsDataList) {
        //获取模板
        String model = ResourceRealPath.resourceDisk+"/model/withdrawModel.xlsx";
        String fileName = UUIDUtil.generateUUID().toString() + ".xls";
        File f = new File(ResourceRealPath.excel + "withdraw/" + fileName);

        SheetData sds = new SheetData("提现记录");

        sds.addDatas(goodsDataList);

        try {
            ExcelUtils.writeData(model, new FileOutputStream(f), sds);
        } catch (FileNotFoundException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        return "/excel/withdraw/" + fileName;
    }

}
