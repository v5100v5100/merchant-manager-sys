package com.mms.util;

import java.util.Date;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
//import org.springframework.web.servlet.view.document.AbstractExcelView;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.web.servlet.view.document.AbstractXlsView;

/**
* 导入到EXCEL
* 类名称：ObjectExcelView.java
* @author 
* @version 1.0
 */
public class ObjectExcelView extends /*AbstractExcelView*/AbstractXlsView {

//	@Override
//	protected void buildExcelDocument(Map<String, Object> model,
//			HSSFWorkbook workbook, HttpServletRequest request,
//			HttpServletResponse response) throws Exception {
//		// TODO Auto-generated method stub
//		Date date = new Date();
//		String filename = Tools.date2Str(date, "yyyyMMddHHmmss");
//		HSSFSheet sheet;
//		HSSFCell cell;
//		response.setContentType("application/octet-stream");
//		response.setHeader("Content-Disposition", "attachment;filename="+filename+".xls");
//		sheet = workbook.createSheet("sheet1");
//
//		List<String> titles = (List<String>) model.get("titles");
//		int len = titles.size();
//		HSSFCellStyle headerStyle = workbook.createCellStyle(); //标题样式
//		headerStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
//		headerStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
//		HSSFFont headerFont = workbook.createFont();	//标题字体
//		headerFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
//		headerFont.setFontHeightInPoints((short)11);
//		headerStyle.setFont(headerFont);
//		short width = 20,height=25*20;
//		sheet.setDefaultColumnWidth(width);
//		for(int i=0; i<len; i++){ //设置标题
//			String title = titles.get(i);
//			cell = getCell(sheet, 0, i);
//			cell.setCellStyle(headerStyle);
//			setText(cell,title);
//		}
//		sheet.getRow(0).setHeight(height);
//
//		HSSFCellStyle contentStyle = workbook.createCellStyle(); //内容样式
//		contentStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
//		List<PageData> varList = (List<PageData>) model.get("varList");
//		int varCount = varList.size();
//		for(int i=0; i<varCount; i++){
//			PageData vpd = varList.get(i);
//			for(int j=0;j<len;j++){
//				String varstr = vpd.getString("var"+(j+1)) != null ? vpd.getString("var"+(j+1)) : "";
//				cell = getCell(sheet, i+1, j);
//				cell.setCellStyle(contentStyle);
//				setText(cell,varstr);
//			}
//		}
//	}


	////////////////////////////////////////////////////////////////////////////////////////////////
	////////////////////////////////////////////////////////////////////////////////////////////////

//	private String fileName=null;
//	private ExcelExportService excelExportService=null;
//	public ExcelView(ExcelExportService excelExportService){
//		this.excelExportService=excelExportService;
//	}
//	public ExcelView(String fileName,ExcelExportService excelExportService){
//		this.fileName=fileName;
//		this.excelExportService=excelExportService;
//	}
//
//	public String getFileName() {
//		return fileName;
//	}
//
//	public void setFileName(String fileName) {
//		this.fileName = fileName;
//	}
//
//	public ExcelExportService getExcelExportService() {
//		return excelExportService;
//	}
//
//	public void setExcelExportService(ExcelExportService excelExportService) {
//		this.excelExportService = excelExportService;
//	}

	/**
	 * Application-provided subclasses must implement this method to populate
	 * the Excel workbook document, given the model.
	 *
	 * @param model    the model Map
	 * @param workbook the Excel workbook to populate
	 * @param request  in case we need locale etc. Shouldn't look at attributes.
	 * @param response in case we need to set cookies. Shouldn't write to it.
	 */
	@Override
	protected void buildExcelDocument(Map<String, Object> model, Workbook workbook, HttpServletRequest request, HttpServletResponse response) throws Exception {
//		if(excelExportService==null){
//			throw new RuntimeException("导出接口不能为空");
//		}
//		if(!StringUtils.isEmpty(fileName)){
//			String reqCharset=request.getCharacterEncoding();
//			reqCharset=reqCharset==null?"UTF-8":reqCharset;
//			fileName=new String(fileName.getBytes(reqCharset),"ISO8859-1");
//			response.setHeader("Content-disposition","attachment;fileName="+fileName);
//		}
		//excelExportService.makeWorkBook(model,workbook);

	}
}
