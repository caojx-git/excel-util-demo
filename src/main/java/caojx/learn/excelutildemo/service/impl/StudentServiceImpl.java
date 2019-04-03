package caojx.learn.excelutildemo.service.impl;

import caojx.learn.excelutildemo.pojo.search.StudentSearch;
import caojx.learn.excelutildemo.pojo.vo.StudentExcelVo;
import caojx.learn.excelutildemo.service.interfaces.IStudentService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.ArrayList;
import java.util.List;

/**
 * 类注释，描述
 *
 * @author caojx
 * @version $Id: StudentService.java,v 1.0 2019-04-03 17:27 caojx
 * @date 2019-04-03 17:27
 */
@Service
public class StudentServiceImpl implements IStudentService {
    /**
     * 线索信息导出excel MovClueExtendExcelVo
     *
     * @param studentSearch @return
     * @return
     */
    @Override
    public List<StudentExcelVo> exportExt(StudentSearch studentSearch) {
        //1.数据库查询学生信息
        //List<Student> studentList = ;
        //2.转成List<StudentExcelVo>

        List<StudentExcelVo> studentExcelVoList = new ArrayList<>();
        StudentExcelVo studentExcelVo1 = new StudentExcelVo();
        studentExcelVo1.setId(1L);
        studentExcelVo1.setName("a");
        studentExcelVo1.setAge(20);
        StudentExcelVo studentExcelVo2 = new StudentExcelVo();
        studentExcelVo2.setId(2L);
        studentExcelVo2.setName("b");
        studentExcelVo2.setAge(30);

        studentExcelVoList.add(studentExcelVo1);
        studentExcelVoList.add(studentExcelVo2);
        return studentExcelVoList;
    }
}