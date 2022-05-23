package cc.mrbird.febs.system.service;

import java.text.SimpleDateFormat;

import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import cc.mrbird.febs.system.entity.Invoice;
import cc.mrbird.febs.system.repository.ResponseUtil;
import cc.mrbird.febs.system.repository.UserDetailRepository;



@Service
public class UserServiceDL {
    @Autowired
    private UserDetailRepository userDetailRepository;
//    @Autowired
//    private UserRepository userRepository;
    /**
     * ユーザー情報 主キー検索
     * @return 検索結果
     */


    public Invoice userfindById(Long id) {
      return userDetailRepository.findById(id).get();
    }


    public void downLoadUserInfoByTemplate( Long id,HttpServletResponse response)throws Exception{

    	final String templateFile = "E:\\javawork\\FEBS-Shiro-2.0\\src\\main\\resources\\templates\\spexcel.xlsx";
    	//File templateFile = new File("spexcel.xlsx"); //サンプル取得
    	Workbook workbook = new XSSFWorkbook(templateFile);
    	Sheet sheet =workbook.getSheetAt(0);

		CellStyle style = workbook.createCellStyle();
		CellStyle style2 = workbook.createCellStyle();
		Font font = workbook.createFont();

		font.setFontHeightInPoints((short) 14);
		style.setFillBackgroundColor(IndexedColors.WHITE.getIndex());
		style2.setFont(font);

    	/**
    	 *
    	 */
		Invoice user= userfindById(id);


    	sheet.getRow(4).getCell(0).setCellStyle(style);
    	sheet.getRow(4).getCell(0).setCellStyle(style2);
    	sheet.getRow(4).getCell(0).setCellValue(user.getCustomerName());
    	sheet.getRow(2).getCell(6).setCellStyle(style);
    	sheet.getRow(2).getCell(6).setCellValue(new SimpleDateFormat("yyyy年MM月dd日").format(user.getInvoiceDate()));
    	sheet.getRow(17).getCell(1).setCellStyle(style);
    	sheet.getRow(17).getCell(1).setCellValue(user.getEmployeeName());
    	sheet.getRow(20).getCell(0).setCellStyle(style);
    	sheet.getRow(20).getCell(0).setCellValue(user.getProjectName());
    	sheet.getRow(20).getCell(5).setCellStyle(style);
		sheet.getRow(20).getCell(5).setCellValue(user.getCount());
		sheet.getRow(20).getCell(6).setCellStyle(style);
		sheet.getRow(20).getCell(6).setCellValue(user.getPrice());


		workbook.setForceFormulaRecalculation(true);


		ResponseUtil.export(response, workbook, "請求書.xlsx");


    }

}
