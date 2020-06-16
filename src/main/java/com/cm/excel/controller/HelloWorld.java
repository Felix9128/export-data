package com.cm.excel.controller;

import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

/**
 * ClassName: HelloWorld
 * Function:  TODO  功能说明.
 * <p>
 * date: 2020年06月15日  23:19
 *
 * @author Felix
 * @since JDK 1.8
 * <p>
 * Modified By： <修改人>
 * Modified Date: <修改日期，格式:YYYY-MM-DD>
 * Why & What is modified: <修改描述>
 */
@RestController
public class HelloWorld {

    @RequestMapping("/index")
    public ModelAndView helloWorld(){
        System.out.println("hello");
        return new ModelAndView("hello");
    }

    @RequestMapping("/export")
    public ModelAndView export(){
        return new ModelAndView("export");
    }
}