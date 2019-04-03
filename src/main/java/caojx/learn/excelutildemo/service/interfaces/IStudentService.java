package caojx.learn.excelutildemo.service.interfaces;

import caojx.learn.excelutildemo.pojo.search.StudentSearch;
import caojx.learn.excelutildemo.pojo.vo.StudentExcelVo;
import java.util.List;

/**
 * 类注释，描述
 *
 * @author caojx
 * @version $Id: IStudentService.java,v 1.0 2019-04-03 17:29 caojx
 * @date 2019-04-03 17:29
 */
public interface IStudentService {
    /**
     * 线索信息导出excel MovClueExtendExcelVo
     *
     * @param studentSearch
     * @return
     */
    List<StudentExcelVo> exportExt(StudentSearch studentSearch);
}