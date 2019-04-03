package caojx.learn.excelutildemo.controller;

import caojx.learn.excelutildemo.pojo.search.StudentSearch;
import caojx.learn.excelutildemo.service.interfaces.IStudentService;
import caojx.learn.excelutildemo.util.ExcelUtils;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.List;
import java.util.Map;


/**
 * 类注释，描述
 *
 * @author caojx
 * @version $Id: ExcelTestController.java,v 1.0 2019-04-03 17:24 caojx
 * @date 2019-04-03 17:24
 */
@Slf4j
@RestController
@RequestMapping("/excel")
public class ExcelTestController {

    /**
     * 文件存放路径
     */
    private static final String UPLOAD_PATH = "/Users/caojx/Desktop/";

    @Autowired
    private IStudentService studentService;

    @GetMapping("/export")
    public void export(HttpServletRequest request, HttpServletResponse response, StudentSearch studentSearch) {
        try {
            ExcelUtils.exportToExcel(request, response, "学生信息.xlsx", studentService.exportExt(studentSearch));
        } catch (Exception e) {
            log.error("export excel error:{}", e);
        }
    }

    @PostMapping("/upload")
    public void export(@RequestParam("file") MultipartFile file) {
        try {
            Map<Integer, List<List<? super Object>>> map = ExcelUtils.uploadExcel(file, UPLOAD_PATH);
            System.out.println(map);
        } catch (Exception e) {
            log.error("export excel error:{}", e);
        }
    }
}