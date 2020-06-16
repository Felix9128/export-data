package com.cm.excel.controller;

import com.cm.excel.utils.ExcelCellAttr;
import com.cm.excel.utils.ExcelExportParms;
import com.cm.excel.utils.ExportExcel;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.*;

/**
 * ClassName: ExcelExportTest
 * Function:  TODO  功能说明.
 * <p>
 * date: 2020年06月15日  22:16
 *
 * @author Felix
 * @since JDK 1.8
 * <p>
 * Modified By： <修改人>
 * Modified Date: <修改日期，格式:YYYY-MM-DD>
 * Why & What is modified: <修改描述>
 */
@RestController
public class ExcelExportController {

    @RequestMapping("/excelExport")
    public void excelExport(HttpServletResponse response) throws IOException {
        List<Map<String,Object>> listMap = new ArrayList<>();
        Map<String, Object> map = new HashMap<>();
        map.put("编号",1);
        map.put("姓名","Felix");
        map.put("性别","男");
        map.put("生日","1991-01-01");
        map.put("身份证号","510603********1111");
        listMap.add(map);

        List<ExcelCellAttr> cellAttrs = Arrays.asList(
                new ExcelCellAttr("编号","编号"),
                new ExcelCellAttr("姓名","姓名"),
                new ExcelCellAttr("性别","性别"),
                new ExcelCellAttr("生日","生日"),
                new ExcelCellAttr("身份证号","身份证号")
        );

        ExcelExportParms<Map<String,Object>> cellAttr = new ExcelExportParms<>();
        cellAttr.setDatas(listMap).setCellAttrs(cellAttrs).setSheetTitle("身份信息表").setFileName("身份信息表");
        ExportExcel.exportExcel(cellAttr,response);
    }

}