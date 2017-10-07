package application.model.service;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFHeader;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import application.model.vo.DamageAndPicture;

public class ExcelImage {
	
	public void make_1(File pictureDir,XSSFWorkbook workbook ,XSSFSheet sheet ,Object sheets ,int pictureNoColumn_) throws Exception {
		int data_st_pic = 0;
		//폰트스타일
		Font fontBody = workbook.createFont();
		fontBody.setColor(HSSFColor.BLACK.index);
		fontBody.setFontHeight((short)220);
		fontBody.setFontName("굴림체");
		//셀스타일
		CellStyle textheader_style = workbook.createCellStyle();
		textheader_style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		textheader_style.setAlignment(CellStyle.ALIGN_CENTER);
		textheader_style.setFont(fontBody);
		textheader_style.setBorderTop(CellStyle.BORDER_MEDIUM);                      
		textheader_style.setBorderLeft(CellStyle.BORDER_MEDIUM);
		textheader_style.setBorderRight(CellStyle.BORDER_MEDIUM);
		textheader_style.setBorderBottom(CellStyle.BORDER_MEDIUM);
		
		CellStyle text_style = workbook.createCellStyle();
		text_style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		text_style.setFont(fontBody);
		text_style.setIndention((short)1);
		text_style.setBorderTop(CellStyle.BORDER_MEDIUM);                      
		text_style.setBorderLeft(CellStyle.BORDER_MEDIUM);
		text_style.setBorderRight(CellStyle.BORDER_MEDIUM);
		text_style.setBorderBottom(CellStyle.BORDER_MEDIUM);

		CellStyle picTop_style = workbook.createCellStyle();
		picTop_style.setBorderLeft(CellStyle.BORDER_MEDIUM);
		picTop_style.setBorderTop(CellStyle.BORDER_MEDIUM);  
		picTop_style.setBorderRight(CellStyle.BORDER_MEDIUM);
		
		CellStyle picL_style = workbook.createCellStyle();
		picL_style.setBorderLeft(CellStyle.BORDER_MEDIUM);

		CellStyle picR_style = workbook.createCellStyle();
		picR_style.setBorderRight(CellStyle.BORDER_MEDIUM);
		
		//하이퍼링크관련
		CellStyle hlink_style = workbook.createCellStyle();
	    Font hlink_font = workbook.createFont();
	    hlink_font.setUnderline(Font.U_SINGLE);
	    hlink_font.setColor(IndexedColors.BLUE.getIndex());
	    hlink_style.setFont(hlink_font);
	    hlink_style.setAlignment(CellStyle.ALIGN_CENTER);
	    hlink_style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
	    
		InputStream pictureFIS;
		Row rowTemp;
		XSSFCell cellTemp;
		
		//사진대지
		sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
		
		Header pageHeader = sheet.getHeader();	//머릿말
		pageHeader.setCenter(HSSFHeader.font("휴먼옛체", "Normal") +HSSFHeader.fontSize((short) 26) + "사 진 대 지");
						
        //출력 row 생성
		rowTemp = sheet.createRow(0);		
			
		List<DamageAndPicture> DamageAndPictureSheet = (List<DamageAndPicture>) sheets;
		int rowcount = 9;
		data_st_pic = 0;
		
		CreationHelper createHelper = workbook.getCreationHelper();//하이퍼링크용
		
		String sheet_hyper_name = sheet.getSheetName();
		sheet_hyper_name = sheet_hyper_name.replaceAll("_사진", ""); //원본시트이름
		
		XSSFSheet sheet_hyper = workbook.getSheet(sheet_hyper_name); // 원본시트를 찾아감.?
		
		for (int i = 0; i < DamageAndPictureSheet.size(); i++) {				
		//로우 페이지 설정
			sheet.setRowBreak(rowcount);
			rowcount =  rowcount + 10;
			DamageAndPicture dmgStatePic = (DamageAndPicture)DamageAndPictureSheet.get(i);
			String picFileNameInExcel = dmgStatePic.getPictureFileNameInExcel();		
			String position = dmgStatePic.getPosition();
			String content = dmgStatePic.getContent();
			String supply = dmgStatePic.getSupply();
			String unit = dmgStatePic.getUnit();
			String ea = dmgStatePic.getEa();
			
			String basePath = pictureDir.getPath();
			File pictureFile = new File(basePath+"\\"+picFileNameInExcel+".jpg");			// 파일정보

			pictureFIS = new FileInputStream(pictureFile);				// InputStream에 파일 set		*FileNotFoundException -> try catch 있어야됨
			byte[] bytes = IOUtils.toByteArray(pictureFIS);				// 이미지 binary를 byte 배열에 담음

			CreationHelper helper = workbook.getCreationHelper();
			XSSFDrawing drawing = sheet.createDrawingPatriarch();	//그림 컨테이너, 그림을 실제로 insert하는 놈
			ClientAnchor anchor = helper.createClientAnchor();	// 그림을 넣을 좌표를 지정하기 위한 객체
			
			rowTemp = sheet.createRow(data_st_pic);
			rowTemp.setHeight((short)500);
			
			for (int k = 0;  k < 10; k++) {
				Cell cells1 = rowTemp.createCell(k);
				cells1.setCellStyle(picTop_style);
				cells1.setCellStyle(picTop_style);
			}
			sheet.addMergedRegion(new CellRangeAddress(
					data_st_pic, //first row (0-based)
					data_st_pic, //last row  (0-based)
			        0, //first column (0-based)
			        9  //last column  (0-based)
			));	
			data_st_pic = data_st_pic +1;
			
			rowTemp = sheet.createRow(data_st_pic);
			rowTemp.setHeight((short)4700);
			Cell cell_ar = rowTemp.createCell(0);//하이퍼링크때문에 변경
			cell_ar.setCellStyle(picL_style);
			rowTemp.createCell(9).setCellStyle(picR_style);
			
			//하이퍼링크
			String cell_name = getCellName(cell_ar);
			Hyperlink link2 = createHelper.createHyperlink(Hyperlink.LINK_DOCUMENT);
			link2.setAddress("'"+sheet.getSheetName()+"'!"+cell_name);
			for (Row row : sheet_hyper) {
				if(row.getCell(pictureNoColumn_)!=null){
					Cell cell = row.getCell(pictureNoColumn_);
					if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
						if (cell.getNumericCellValue() == Integer.parseInt(picFileNameInExcel)) {
							cell.setHyperlink(link2);
							cell.setCellStyle(hlink_style);
						}
					}
				}
			}					
			
			anchor.setCol1(1);	
			anchor.setRow1(data_st_pic);	// => A의 우측, 1행의 아래가 그림의 좌측상단 = 시작점이 된다.
			anchor.setCol2(9);					
			data_st_pic = data_st_pic+1;
			anchor.setRow2(data_st_pic);	// => D의 우측, 5행의 아래가 그림의 우측하단 = 끝점이 된다
							
			int pictureIndex = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);	// workbook에 추가하고, 추가한 그림의 id 같은걸 받아옴
			
			XSSFPicture picture = drawing.createPicture(anchor, pictureIndex);			// 좌표(anchor)와 그림id로 그림을 insert함	
			//( 참고 ) picture.resize();	// 원래 이미지 크기로 resize하는것
			
			rowTemp = sheet.createRow(data_st_pic);
			rowTemp.setHeight((short)500);
			for (int k = 0; k < 10; k++) {
				Cell cells1 = rowTemp.createCell(k);
				if(k==0){
					cells1.setCellStyle(picL_style);
				}else if(k==9){
					cells1.setCellStyle(picR_style);
				}
			}
			
			data_st_pic = data_st_pic +1;
			
			//출력Cell 생성
			rowTemp = sheet.createRow(data_st_pic);
			rowTemp.setHeight((short)500);
			
			for (int k = 0; k < 10; k++) {
				Cell cells1 = rowTemp.createCell(k);
				if(k == 0){
					cells1.setCellValue("위  치");
					cells1.setCellStyle(textheader_style);
				}else if(k == 2){
					cells1.setCellValue(position);
					cells1.setCellStyle(text_style);
				}else{
					cells1.setCellStyle(text_style);
				}
			}
			
			sheet.addMergedRegion(new CellRangeAddress(
					data_st_pic, //first row (0-based)
					data_st_pic, //last row  (0-based)
			        0, //first column (0-based)
			        1  //last column  (0-based)
			));	
			sheet.addMergedRegion(new CellRangeAddress(
					data_st_pic, //first row (0-based)
					data_st_pic, //last row  (0-based)
			        2, //first column (0-based)
			        9  //last column  (0-based)
			));
			
			data_st_pic = data_st_pic+1;	
			rowTemp = sheet.createRow(data_st_pic);
			rowTemp.setHeight((short)500);
			
			for (int k = 0; k < 10; k++) {
				Cell cells1 = rowTemp.createCell(k);
				if(k == 0){
					cells1.setCellValue("내  용");
					cells1.setCellStyle(textheader_style);
				}else if(k == 2){
					cells1.setCellValue(content);
					cells1.setCellStyle(text_style);
				}else if(k == 6){
					cells1.setCellValue(supply+" / "+unit+" / "+ea+"EA");
					cells1.setCellStyle(text_style);
				}else{
					cells1.setCellStyle(text_style);
				}					
			}
			
			sheet.addMergedRegion(new CellRangeAddress(
					data_st_pic, //first row (0-based)
					data_st_pic, //last row  (0-based)
			        0, //first column (0-based)
			        1  //last column  (0-based)
			));	
			sheet.addMergedRegion(new CellRangeAddress(
					data_st_pic, //first row (0-based)
					data_st_pic, //last row  (0-based)
			        2, //first column (0-based)
			        5  //last column  (0-based)
			));
			sheet.addMergedRegion(new CellRangeAddress(
					data_st_pic, //first row (0-based)
					data_st_pic, //last row  (0-based)
			        6, //first column (0-based)
			        9  //last column  (0-based)
			));
			data_st_pic = data_st_pic+1;
		}
	}
	
	public void make_2(File pictureDir,XSSFWorkbook workbook ,XSSFSheet sheet ,Object sheets ,int pictureNoColumn_) throws Exception {
	
		//폰트스타일
		Font fontBody = workbook.createFont();
		fontBody.setColor(HSSFColor.BLACK.index);
		fontBody.setFontHeight((short)220);
		fontBody.setFontName("굴림체");

		//셀스타일
		CellStyle textheader_style = workbook.createCellStyle();
		textheader_style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		textheader_style.setAlignment(CellStyle.ALIGN_CENTER);
		textheader_style.setFont(fontBody);
		textheader_style.setBorderTop(CellStyle.BORDER_MEDIUM);                      
		textheader_style.setBorderLeft(CellStyle.BORDER_MEDIUM);
		textheader_style.setBorderRight(CellStyle.BORDER_MEDIUM);
		textheader_style.setBorderBottom(CellStyle.BORDER_MEDIUM);
		
		CellStyle text_style = workbook.createCellStyle();
		text_style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		text_style.setFont(fontBody);
		text_style.setIndention((short)1);
		text_style.setBorderTop(CellStyle.BORDER_MEDIUM);                      
		text_style.setBorderLeft(CellStyle.BORDER_MEDIUM);
		text_style.setBorderRight(CellStyle.BORDER_MEDIUM);
		text_style.setBorderBottom(CellStyle.BORDER_MEDIUM);

		CellStyle picTop_style = workbook.createCellStyle();
		picTop_style.setBorderLeft(CellStyle.BORDER_MEDIUM);
		picTop_style.setBorderTop(CellStyle.BORDER_MEDIUM);  
		picTop_style.setBorderRight(CellStyle.BORDER_MEDIUM);
		
		CellStyle picL_style = workbook.createCellStyle();
		picL_style.setBorderLeft(CellStyle.BORDER_MEDIUM);

		CellStyle picR_style = workbook.createCellStyle();
		picR_style.setBorderRight(CellStyle.BORDER_MEDIUM);
		
		//하이퍼링크관련
		CellStyle hlink_style = workbook.createCellStyle();
		Font hlink_font = workbook.createFont();
		hlink_font.setUnderline(Font.U_SINGLE);
		hlink_font.setColor(IndexedColors.BLUE.getIndex());
		hlink_style.setFont(hlink_font);
		hlink_style.setAlignment(CellStyle.ALIGN_CENTER);
		hlink_style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		    
		
		InputStream pictureFIS;
		Row rowTemp_pic1,rowTemp_pic2,rowTemp_pic3,rowTemp_pos,rowTemp_data;
		XSSFCell cellTemp;
		
		//사진대지
		sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
		
		Header pageHeader = sheet.getHeader();	//머릿말
		pageHeader.setCenter(HSSFHeader.font("휴먼옛체", "Normal") +HSSFHeader.fontSize((short) 26) + "사 진 대 지");
						
        //출력 row 생성
		rowTemp_pic1 = sheet.createRow(0);
		rowTemp_pic2 = sheet.createRow(0);
		rowTemp_pic3 = sheet.createRow(0);
		
		rowTemp_pos = sheet.createRow(0);
		rowTemp_data = sheet.createRow(0);
		
		List<DamageAndPicture> DamageAndPictureSheet = (List<DamageAndPicture>) sheets;

		CreationHelper createHelper = workbook.getCreationHelper();//하이퍼링크용
		
		String sheet_hyper_name = sheet.getSheetName();
		sheet_hyper_name = sheet_hyper_name.replaceAll("_사진", ""); //원본시트이름
		
		XSSFSheet sheet_hyper = workbook.getSheet(sheet_hyper_name); // 원본시트를 찾아감.?
		
		int odd_rowcount = 9;
		
		int odd_st_pic = 0;
		int even_st_pic = 0;
		
		for (int i = 0; i < DamageAndPictureSheet.size(); i++) {		
			if (i % 2 == 0){
				odd_rowcount =  odd_rowcount + 10;
				DamageAndPicture dmgStatePic = (DamageAndPicture)DamageAndPictureSheet.get(i);
				String picFileNameInExcel = dmgStatePic.getPictureFileNameInExcel();		
				String position = dmgStatePic.getPosition();
				String content = dmgStatePic.getContent();
				String supply = dmgStatePic.getSupply();
				String unit = dmgStatePic.getUnit();
				String ea = dmgStatePic.getEa();
				
				String basePath = pictureDir.getPath();
				File pictureFile = new File(basePath+"\\"+picFileNameInExcel+".jpg");			// 파일정보
				
				pictureFIS = new FileInputStream(pictureFile);				// InputStream에 파일 set		*FileNotFoundException -> try catch 있어야됨
				byte[] bytes = IOUtils.toByteArray(pictureFIS);				// 이미지 binary를 byte 배열에 담음
	
				CreationHelper helper = workbook.getCreationHelper();
				XSSFDrawing drawing = sheet.createDrawingPatriarch();	//그림 컨테이너, 그림을 실제로 insert하는 놈
				ClientAnchor anchor = helper.createClientAnchor();	// 그림을 넣을 좌표를 지정하기 위한 객체
				
				rowTemp_pic1 = sheet.createRow(odd_st_pic);
				rowTemp_pic1.setHeight((short)500);
				
				for (int k = 0;  k < 10; k++) {
					Cell cells1 = rowTemp_pic1.createCell(k);
					cells1.setCellStyle(picTop_style);
					cells1.setCellStyle(picTop_style);
				}
				sheet.addMergedRegion(new CellRangeAddress(
						odd_st_pic, //first row (0-based)
						odd_st_pic, //last row  (0-based)
				        0, //first column (0-based)
				        9  //last column  (0-based)
				));	
				odd_st_pic = odd_st_pic +1;
				
				rowTemp_pic2 = sheet.createRow(odd_st_pic);
				rowTemp_pic2.setHeight((short)4700);
				Cell cell_ar = rowTemp_pic2.createCell(0);
				cell_ar.setCellStyle(picL_style);
				rowTemp_pic2.createCell(9).setCellStyle(picR_style);
				
				//하이퍼링크
				String cell_name = getCellName(cell_ar);
				Hyperlink link2 = createHelper.createHyperlink(Hyperlink.LINK_DOCUMENT);
				link2.setAddress("'"+sheet.getSheetName()+"'!"+cell_name);
				for (Row row : sheet_hyper) {
					if(row.getCell(pictureNoColumn_)!=null){
						Cell cell = row.getCell(pictureNoColumn_);
						if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
							if (cell.getNumericCellValue() == Integer.parseInt(picFileNameInExcel)) {
								cell.setHyperlink(link2);
								cell.setCellStyle(hlink_style);
							}
						}
					}
				}							
				
				anchor.setCol1(1);	
				anchor.setRow1(odd_st_pic);	// => A의 우측, 1행의 아래가 그림의 좌측상단 = 시작점이 된다.
				anchor.setCol2(9);					
				odd_st_pic = odd_st_pic+1;
				anchor.setRow2(odd_st_pic);	// => D의 우측, 5행의 아래가 그림의 우측하단 = 끝점이 된다
								
				int pictureIndex = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);	// workbook에 추가하고, 추가한 그림의 id 같은걸 받아옴
				
				XSSFPicture picture = drawing.createPicture(anchor, pictureIndex);			// 좌표(anchor)와 그림id로 그림을 insert함	
				//( 참고 ) picture.resize();	// 원래 이미지 크기로 resize하는것
				
				rowTemp_pic3 = sheet.createRow(odd_st_pic);
				rowTemp_pic3.setHeight((short)500);
	
				for (int k = 0; k < 10; k++) {
					Cell cells1 = rowTemp_pic3.createCell(k);
					if(k==0){
						cells1.setCellStyle(picL_style);
					}else if(k==9){
						cells1.setCellStyle(picR_style);
					}
				}
				
				odd_st_pic = odd_st_pic +1;
				
				//출력Cell 생성
				rowTemp_pos = sheet.createRow(odd_st_pic);
				rowTemp_pos.setHeight((short)500);
				
				for (int k = 0; k < 10; k++) {
					Cell cells1 = rowTemp_pos.createCell(k);
					if(k == 0){
						cells1.setCellValue("위  치");
						cells1.setCellStyle(textheader_style);
					}else if(k == 2){
						cells1.setCellValue(position);
						cells1.setCellStyle(text_style);
					}else{
						cells1.setCellStyle(text_style);
					}
				}
				
				sheet.addMergedRegion(new CellRangeAddress(
						odd_st_pic, //first row (0-based)
						odd_st_pic, //last row  (0-based)
				        0, //first column (0-based)
				        1  //last column  (0-based)
				));	
				sheet.addMergedRegion(new CellRangeAddress(
						odd_st_pic, //first row (0-based)
						odd_st_pic, //last row  (0-based)
				        2, //first column (0-based)
				        9  //last column  (0-based)
				));
				
				odd_st_pic = odd_st_pic+1;	
				rowTemp_data = sheet.createRow(odd_st_pic);
				rowTemp_data.setHeight((short)500);
				
				for (int k = 0; k < 10; k++) {
					Cell cells1 = rowTemp_data.createCell(k);
					if(k == 0){
						cells1.setCellValue("내  용");
						cells1.setCellStyle(textheader_style);
					}else if(k == 2){
						cells1.setCellValue(content);
						cells1.setCellStyle(text_style);
					}else if(k == 6){
						cells1.setCellValue(supply+" / "+unit+" / "+ea+"EA");
						cells1.setCellStyle(text_style);
					}else{
						cells1.setCellStyle(text_style);
					}					
				}
				
				sheet.addMergedRegion(new CellRangeAddress(
						odd_st_pic, //first row (0-based)
						odd_st_pic, //last row  (0-based)
				        0, //first column (0-based)
				        1  //last column  (0-based)
				));	
				sheet.addMergedRegion(new CellRangeAddress(
						odd_st_pic, //first row (0-based)
						odd_st_pic, //last row  (0-based)
				        2, //first column (0-based)
				        5  //last column  (0-based)
				));
				sheet.addMergedRegion(new CellRangeAddress(
						odd_st_pic, //first row (0-based)
						odd_st_pic, //last row  (0-based)
				        6, //first column (0-based)
				        9  //last column  (0-based)
				));
				odd_st_pic = odd_st_pic+1;
			}else{
				DamageAndPicture dmgStatePic = (DamageAndPicture)DamageAndPictureSheet.get(i);
				String picFileNameInExcel = dmgStatePic.getPictureFileNameInExcel();		
				String position = dmgStatePic.getPosition();
				String content = dmgStatePic.getContent();
				String supply = dmgStatePic.getSupply();
				String unit = dmgStatePic.getUnit();
				String ea = dmgStatePic.getEa();
				
				String basePath = pictureDir.getPath();
				File pictureFile = new File(basePath+"\\"+picFileNameInExcel+".jpg");			// 파일정보
				
				pictureFIS = new FileInputStream(pictureFile);				// InputStream에 파일 set		*FileNotFoundException -> try catch 있어야됨
				byte[] bytes = IOUtils.toByteArray(pictureFIS);				// 이미지 binary를 byte 배열에 담음
	
				CreationHelper helper = workbook.getCreationHelper();
				XSSFDrawing drawing = sheet.createDrawingPatriarch();	//그림 컨테이너, 그림을 실제로 insert하는 놈
				ClientAnchor anchor = helper.createClientAnchor();	// 그림을 넣을 좌표를 지정하기 위한 객체
				
								
				for (int k = 10;  k < 20; k++) {
					Cell cells1 = rowTemp_pic1.createCell(k);
					cells1.setCellStyle(picTop_style);
					cells1.setCellStyle(picTop_style);
				}
				sheet.addMergedRegion(new CellRangeAddress(
						even_st_pic, //first row (0-based)
						even_st_pic, //last row  (0-based)
				        10, //first column (0-based)
				        19  //last column  (0-based)
				));	
				even_st_pic = even_st_pic +1;
				
				Cell cell_ar = rowTemp_pic2.createCell(10);
				cell_ar.setCellStyle(picL_style);
				rowTemp_pic2.createCell(19).setCellStyle(picR_style);
				
				//하이퍼링크
				String cell_name = getCellName(cell_ar);
				Hyperlink link2 = createHelper.createHyperlink(Hyperlink.LINK_DOCUMENT);
				link2.setAddress("'"+sheet.getSheetName()+"'!"+cell_name);
				for (Row row : sheet_hyper) {
					if(row.getCell(pictureNoColumn_)!=null){
						Cell cell = row.getCell(pictureNoColumn_);
						if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
							if (cell.getNumericCellValue() == Integer.parseInt(picFileNameInExcel)) {
								cell.setHyperlink(link2);
								cell.setCellStyle(hlink_style);
							}
						}
					}
				}
				
				anchor.setCol1(11);	
				anchor.setRow1(even_st_pic);	// => A의 우측, 1행의 아래가 그림의 좌측상단 = 시작점이 된다.
				anchor.setCol2(19);					
				even_st_pic = even_st_pic+1;
				anchor.setRow2(even_st_pic);	// => D의 우측, 5행의 아래가 그림의 우측하단 = 끝점이 된다
								
				int pictureIndex = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);	// workbook에 추가하고, 추가한 그림의 id 같은걸 받아옴
				
				XSSFPicture picture = drawing.createPicture(anchor, pictureIndex);			// 좌표(anchor)와 그림id로 그림을 insert함	
				//( 참고 ) picture.resize();	// 원래 이미지 크기로 resize하는것
			
				for (int k = 10; k < 20; k++) {
					Cell cells1 = rowTemp_pic3.createCell(k);
					if(k==10){
						cells1.setCellStyle(picL_style);
					}else if(k==19){
						cells1.setCellStyle(picR_style);
					}
				}
				
				even_st_pic = even_st_pic +1;
				
				//출력Cell 생성
				
				for (int k = 10; k < 20; k++) {
					Cell cells1 = rowTemp_pos.createCell(k);
					if(k == 10){
						cells1.setCellValue("위  치");
						cells1.setCellStyle(textheader_style);
					}else if(k == 12){
						cells1.setCellValue(position);
						cells1.setCellStyle(text_style);
					}else{
						cells1.setCellStyle(text_style);
					}
				}
				
				sheet.addMergedRegion(new CellRangeAddress(
						even_st_pic, //first row (0-based)
						even_st_pic, //last row  (0-based)
				        10, //first column (0-based)
				        11  //last column  (0-based)
				));	
				sheet.addMergedRegion(new CellRangeAddress(
						even_st_pic, //first row (0-based)
						even_st_pic, //last row  (0-based)
				        12, //first column (0-based)
				        19  //last column  (0-based)
				));
				
				even_st_pic = even_st_pic+1;
				
				for (int k = 10; k < 20; k++) {
					Cell cells1 = rowTemp_data.createCell(k);
					if(k == 10){
						cells1.setCellValue("내  용");
						cells1.setCellStyle(textheader_style);
					}else if(k == 12){
						cells1.setCellValue(content);
						cells1.setCellStyle(text_style);
					}else if(k == 16){
						cells1.setCellValue(supply+" / "+unit+" / "+ea+"EA");
						cells1.setCellStyle(text_style);
					}else{
						cells1.setCellStyle(text_style);
					}					
				}
				
				sheet.addMergedRegion(new CellRangeAddress(
						even_st_pic, //first row (0-based)
						even_st_pic, //last row  (0-based)
				        10, //first column (0-based)
				        11  //last column  (0-based)
				));	
				sheet.addMergedRegion(new CellRangeAddress(
						even_st_pic, //first row (0-based)
						even_st_pic, //last row  (0-based)
				        12, //first column (0-based)
				        15  //last column  (0-based)
				));
				sheet.addMergedRegion(new CellRangeAddress(
						even_st_pic, //first row (0-based)
						even_st_pic, //last row  (0-based)
				        16, //first column (0-based)
				        19  //last column  (0-based)
				));
				even_st_pic = even_st_pic+1;

				//로우 페이지 설정
				sheet.setRowBreak(odd_rowcount);
			}
		}
	}
	
	public void make_3(File pictureDir,XSSFWorkbook workbook ,XSSFSheet sheet ,Object sheets ,int pictureNoColumn_) throws Exception {
		//폰트스타일
		Font fontBody = workbook.createFont();
		fontBody.setColor(HSSFColor.BLACK.index);
		fontBody.setFontHeight((short)220);
		fontBody.setFontName("굴림체");

		//셀스타일
		CellStyle textheader_style = workbook.createCellStyle();
		textheader_style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		textheader_style.setAlignment(CellStyle.ALIGN_CENTER);
		textheader_style.setFont(fontBody);
		textheader_style.setBorderTop(CellStyle.BORDER_MEDIUM);                      
		textheader_style.setBorderLeft(CellStyle.BORDER_MEDIUM);
		textheader_style.setBorderRight(CellStyle.BORDER_MEDIUM);
		textheader_style.setBorderBottom(CellStyle.BORDER_MEDIUM);
				
		CellStyle text_style = workbook.createCellStyle();
		text_style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		text_style.setFont(fontBody);
		text_style.setIndention((short)1);
		text_style.setBorderTop(CellStyle.BORDER_MEDIUM);                      
		text_style.setBorderLeft(CellStyle.BORDER_MEDIUM);
		text_style.setBorderRight(CellStyle.BORDER_MEDIUM);
		text_style.setBorderBottom(CellStyle.BORDER_MEDIUM);

		CellStyle picTop_style = workbook.createCellStyle();
		picTop_style.setBorderLeft(CellStyle.BORDER_MEDIUM);
		picTop_style.setBorderTop(CellStyle.BORDER_MEDIUM);  
		picTop_style.setBorderRight(CellStyle.BORDER_MEDIUM);
		
		CellStyle picL_style = workbook.createCellStyle();
		picL_style.setBorderLeft(CellStyle.BORDER_MEDIUM);

		CellStyle picR_style = workbook.createCellStyle();
		picR_style.setBorderRight(CellStyle.BORDER_MEDIUM);
		
		//하이퍼링크관련
		CellStyle hlink_style = workbook.createCellStyle();
		Font hlink_font = workbook.createFont();
		hlink_font.setUnderline(Font.U_SINGLE);
		hlink_font.setColor(IndexedColors.BLUE.getIndex());
		hlink_style.setFont(hlink_font);
		hlink_style.setAlignment(CellStyle.ALIGN_CENTER);
		hlink_style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);

				
		InputStream pictureFIS;
		Row rowTemp_pic1,rowTemp_pic2,rowTemp_pic3,rowTemp_pos,rowTemp_data;
		XSSFCell cellTemp;
			
		//사진대지
		sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
				
		Header pageHeader = sheet.getHeader();	//머릿말
		pageHeader.setCenter(HSSFHeader.font("휴먼옛체", "Normal") +HSSFHeader.fontSize((short) 26) + "사 진 대 지");
								
		//출력 row 생성
		rowTemp_pic1 = sheet.createRow(0);
		rowTemp_pic2 = sheet.createRow(0);
		rowTemp_pic3 = sheet.createRow(0);
				
		rowTemp_pos = sheet.createRow(0);
		rowTemp_data = sheet.createRow(0);
				
		List<DamageAndPicture> DamageAndPictureSheet = (List<DamageAndPicture>) sheets;
		
		CreationHelper createHelper = workbook.getCreationHelper();//하이퍼링크용
		
		String sheet_hyper_name = sheet.getSheetName();
		sheet_hyper_name = sheet_hyper_name.replaceAll("_사진", ""); //원본시트이름
		
		XSSFSheet sheet_hyper = workbook.getSheet(sheet_hyper_name); // 원본시트를 찾아감.?

		int odd_rowcount = 9;
				
		int odd_st_pic = 0;
		int even_st_pic = 0;
		int third_st_pic = 0;
		
		for (int i = 0; i < DamageAndPictureSheet.size(); i++) {		
			if (i % 3 == 0){
				odd_rowcount =  odd_rowcount + 10;
				DamageAndPicture dmgStatePic = (DamageAndPicture)DamageAndPictureSheet.get(i);
				String picFileNameInExcel = dmgStatePic.getPictureFileNameInExcel();		
				String position = dmgStatePic.getPosition();
				String content = dmgStatePic.getContent();
				String supply = dmgStatePic.getSupply();
				String unit = dmgStatePic.getUnit();
				String ea = dmgStatePic.getEa();
			
				String basePath = pictureDir.getPath();
				File pictureFile = new File(basePath+"\\"+picFileNameInExcel+".jpg");			// 파일정보
				
				pictureFIS = new FileInputStream(pictureFile);				// InputStream에 파일 set		*FileNotFoundException -> try catch 있어야됨
				byte[] bytes = IOUtils.toByteArray(pictureFIS);				// 이미지 binary를 byte 배열에 담음
		
				CreationHelper helper = workbook.getCreationHelper();
				XSSFDrawing drawing = sheet.createDrawingPatriarch();	//그림 컨테이너, 그림을 실제로 insert하는 놈
				ClientAnchor anchor = helper.createClientAnchor();	// 그림을 넣을 좌표를 지정하기 위한 객체
						
				rowTemp_pic1 = sheet.createRow(odd_st_pic);
				rowTemp_pic1.setHeight((short)500);
						
				for (int k = 0;  k < 10; k++) {
					Cell cells1 = rowTemp_pic1.createCell(k);
					cells1.setCellStyle(picTop_style);
					cells1.setCellStyle(picTop_style);
				}
				sheet.addMergedRegion(new CellRangeAddress(
						odd_st_pic, //first row (0-based)
						odd_st_pic, //last row  (0-based)
				        0, //first column (0-based)
				        9  //last column  (0-based)
				));	
				
				odd_st_pic = odd_st_pic +1;
				rowTemp_pic2 = sheet.createRow(odd_st_pic);
				rowTemp_pic2.setHeight((short)4700);
				Cell cell_ar = rowTemp_pic2.createCell(0);
				cell_ar.setCellStyle(picL_style);
				rowTemp_pic2.createCell(9).setCellStyle(picR_style);
						
				//하이퍼링크
				String cell_name = getCellName(cell_ar);
				Hyperlink link2 = createHelper.createHyperlink(Hyperlink.LINK_DOCUMENT);
				link2.setAddress("'"+sheet.getSheetName()+"'!"+cell_name);
				for (Row row : sheet_hyper) {
					if(row.getCell(pictureNoColumn_)!=null){
						Cell cell = row.getCell(pictureNoColumn_);
						if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
							if (cell.getNumericCellValue() == Integer.parseInt(picFileNameInExcel)) {
								cell.setHyperlink(link2);
								cell.setCellStyle(hlink_style);
							}
						}
					}
				}
				
				anchor.setCol1(1);	
				anchor.setRow1(odd_st_pic);	// => A의 우측, 1행의 아래가 그림의 좌측상단 = 시작점이 된다.
				anchor.setCol2(9);					
				odd_st_pic = odd_st_pic+1;
				anchor.setRow2(odd_st_pic);	// => D의 우측, 5행의 아래가 그림의 우측하단 = 끝점이 된다
										
				int pictureIndex = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);	// workbook에 추가하고, 추가한 그림의 id 같은걸 받아옴
						
				XSSFPicture picture = drawing.createPicture(anchor, pictureIndex);			// 좌표(anchor)와 그림id로 그림을 insert함	
				//( 참고 ) picture.resize();	// 원래 이미지 크기로 resize하는것
						
				rowTemp_pic3 = sheet.createRow(odd_st_pic);
				rowTemp_pic3.setHeight((short)500);
			
				for (int k = 0; k < 10; k++) {
					Cell cells1 = rowTemp_pic3.createCell(k);
					if(k==0){
						cells1.setCellStyle(picL_style);
					}else if(k==9){
						cells1.setCellStyle(picR_style);
					}
				}
						
				odd_st_pic = odd_st_pic +1;
						
				//출력Cell 생성
				rowTemp_pos = sheet.createRow(odd_st_pic);
				rowTemp_pos.setHeight((short)500);
						
				for (int k = 0; k < 10; k++) {
					Cell cells1 = rowTemp_pos.createCell(k);
					if(k == 0){
						cells1.setCellValue("위  치");
						cells1.setCellStyle(textheader_style);
					}else if(k == 2){
						cells1.setCellValue(position);
						cells1.setCellStyle(text_style);
					}else{
						cells1.setCellStyle(text_style);
					}
				}
						
				sheet.addMergedRegion(new CellRangeAddress(
						odd_st_pic, //first row (0-based)
						odd_st_pic, //last row  (0-based)
				        0, //first column (0-based)
				        1  //last column  (0-based)
				));	
				sheet.addMergedRegion(new CellRangeAddress(
						odd_st_pic, //first row (0-based)
						odd_st_pic, //last row  (0-based)
				        2, //first column (0-based)
				        9  //last column  (0-based)
				));
				
				odd_st_pic = odd_st_pic+1;	
				rowTemp_data = sheet.createRow(odd_st_pic);
				rowTemp_data.setHeight((short)500);
				
				for (int k = 0; k < 10; k++) {
					Cell cells1 = rowTemp_data.createCell(k);
					if(k == 0){
						cells1.setCellValue("내  용");
						cells1.setCellStyle(textheader_style);
					}else if(k == 2){
						cells1.setCellValue(content);
						cells1.setCellStyle(text_style);
					}else if(k == 6){
						cells1.setCellValue(supply+" / "+unit+" / "+ea+"EA");
						cells1.setCellStyle(text_style);
					}else{
						cells1.setCellStyle(text_style);
					}					
				}
						
				sheet.addMergedRegion(new CellRangeAddress(
						odd_st_pic, //first row (0-based)
						odd_st_pic, //last row  (0-based)
				        0, //first column (0-based)
				        1  //last column  (0-based)
				));	
				sheet.addMergedRegion(new CellRangeAddress(
						odd_st_pic, //first row (0-based)
						odd_st_pic, //last row  (0-based)
				        2, //first column (0-based)
				        5  //last column  (0-based)
				));
				sheet.addMergedRegion(new CellRangeAddress(
						odd_st_pic, //first row (0-based)
						odd_st_pic, //last row  (0-based)
				        6, //first column (0-based)
				        9  //last column  (0-based)
				));
				odd_st_pic = odd_st_pic+1;
			}else if (i % 3 == 1){
				DamageAndPicture dmgStatePic = (DamageAndPicture)DamageAndPictureSheet.get(i);
				String picFileNameInExcel = dmgStatePic.getPictureFileNameInExcel();		
				String position = dmgStatePic.getPosition();
				String content = dmgStatePic.getContent();
				String supply = dmgStatePic.getSupply();
				String unit = dmgStatePic.getUnit();
				String ea = dmgStatePic.getEa();
					
				String basePath = pictureDir.getPath();
				File pictureFile = new File(basePath+"\\"+picFileNameInExcel+".jpg");			// 파일정보
				
				pictureFIS = new FileInputStream(pictureFile);				// InputStream에 파일 set		*FileNotFoundException -> try catch 있어야됨
				byte[] bytes = IOUtils.toByteArray(pictureFIS);				// 이미지 binary를 byte 배열에 담음
			
				CreationHelper helper = workbook.getCreationHelper();
				XSSFDrawing drawing = sheet.createDrawingPatriarch();	//그림 컨테이너, 그림을 실제로 insert하는 놈
				ClientAnchor anchor = helper.createClientAnchor();	// 그림을 넣을 좌표를 지정하기 위한 객체
				
				for (int k = 10;  k < 20; k++) {
					Cell cells1 = rowTemp_pic1.createCell(k);
					cells1.setCellStyle(picTop_style);
					cells1.setCellStyle(picTop_style);
				}
			
				sheet.addMergedRegion(new CellRangeAddress(
						even_st_pic, //first row (0-based)
						even_st_pic, //last row  (0-based)
				        10, //first column (0-based)
				        19  //last column  (0-based)
				));	
				even_st_pic = even_st_pic +1;
						
				Cell cell_ar = rowTemp_pic2.createCell(10);
				cell_ar.setCellStyle(picL_style);
				rowTemp_pic2.createCell(19).setCellStyle(picR_style);
				
				//하이퍼링크
				String cell_name = getCellName(cell_ar);
				Hyperlink link2 = createHelper.createHyperlink(Hyperlink.LINK_DOCUMENT);
				link2.setAddress("'"+sheet.getSheetName()+"'!"+cell_name);
				for (Row row : sheet_hyper) {
					if(row.getCell(pictureNoColumn_)!=null){
						Cell cell = row.getCell(pictureNoColumn_);
						if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
							if (cell.getNumericCellValue() == Integer.parseInt(picFileNameInExcel)) {
								cell.setHyperlink(link2);
								cell.setCellStyle(hlink_style);
							}
						}
					}
				}
				anchor.setCol1(11);	
				anchor.setRow1(even_st_pic);	// => A의 우측, 1행의 아래가 그림의 좌측상단 = 시작점이 된다.
				anchor.setCol2(19);					
				even_st_pic = even_st_pic+1;
				anchor.setRow2(even_st_pic);	// => D의 우측, 5행의 아래가 그림의 우측하단 = 끝점이 된다
								
				int pictureIndex = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);	// workbook에 추가하고, 추가한 그림의 id 같은걸 받아옴
						
				XSSFPicture picture = drawing.createPicture(anchor, pictureIndex);			// 좌표(anchor)와 그림id로 그림을 insert함	
				//( 참고 ) picture.resize();	// 원래 이미지 크기로 resize하는것
					
				for (int k = 10; k < 20; k++) {
					Cell cells1 = rowTemp_pic3.createCell(k);
					if(k==10){
						cells1.setCellStyle(picL_style);
					}else if(k==19){
						cells1.setCellStyle(picR_style);
					}
				}
						
				even_st_pic = even_st_pic +1;
				
				//출력Cell 생성		
				for (int k = 10; k < 20; k++) {
					Cell cells1 = rowTemp_pos.createCell(k);
					if(k == 10){
						cells1.setCellValue("위  치");
						cells1.setCellStyle(textheader_style);
					}else if(k == 12){
						cells1.setCellValue(position);
						cells1.setCellStyle(text_style);
					}else{
						cells1.setCellStyle(text_style);
					}
				}
					
				sheet.addMergedRegion(new CellRangeAddress(
						even_st_pic, //first row (0-based)
						even_st_pic, //last row  (0-based)
				        10, //first column (0-based)
				        11  //last column  (0-based)
				));	
				sheet.addMergedRegion(new CellRangeAddress(
						even_st_pic, //first row (0-based)
						even_st_pic, //last row  (0-based)
				        12, //first column (0-based)
				        19  //last column  (0-based)
				));
						
				even_st_pic = even_st_pic+1;
						
				for (int k = 10; k < 20; k++) {
					Cell cells1 = rowTemp_data.createCell(k);
					if(k == 10){
						cells1.setCellValue("내  용");
						cells1.setCellStyle(textheader_style);
					}else if(k == 12){
						cells1.setCellValue(content);
						cells1.setCellStyle(text_style);
					}else if(k == 16){
						cells1.setCellValue(supply+" / "+unit+" / "+ea+"EA");
						cells1.setCellStyle(text_style);
					}else{
						cells1.setCellStyle(text_style);
					}					
				}
						
				sheet.addMergedRegion(new CellRangeAddress(
						even_st_pic, //first row (0-based)
						even_st_pic, //last row  (0-based)
				        10, //first column (0-based)
				        11  //last column  (0-based)
				));	
				sheet.addMergedRegion(new CellRangeAddress(
						even_st_pic, //first row (0-based)
						even_st_pic, //last row  (0-based)
				        12, //first column (0-based)
				        15  //last column  (0-based)
				));
				sheet.addMergedRegion(new CellRangeAddress(
						even_st_pic, //first row (0-based)
						even_st_pic, //last row  (0-based)
				        16, //first column (0-based)
				        19  //last column  (0-based)
				));
				even_st_pic = even_st_pic+1;
			}else{
				
				DamageAndPicture dmgStatePic = (DamageAndPicture)DamageAndPictureSheet.get(i);
				String picFileNameInExcel = dmgStatePic.getPictureFileNameInExcel();		
				String position = dmgStatePic.getPosition();
				String content = dmgStatePic.getContent();
				String supply = dmgStatePic.getSupply();
				String unit = dmgStatePic.getUnit();
				String ea = dmgStatePic.getEa();
					
				String basePath = pictureDir.getPath();
				File pictureFile = new File(basePath+"\\"+picFileNameInExcel+".jpg");			// 파일정보
				
				pictureFIS = new FileInputStream(pictureFile);				// InputStream에 파일 set		*FileNotFoundException -> try catch 있어야됨
				byte[] bytes = IOUtils.toByteArray(pictureFIS);				// 이미지 binary를 byte 배열에 담음
			
				CreationHelper helper = workbook.getCreationHelper();
				XSSFDrawing drawing = sheet.createDrawingPatriarch();	//그림 컨테이너, 그림을 실제로 insert하는 놈
				ClientAnchor anchor = helper.createClientAnchor();	// 그림을 넣을 좌표를 지정하기 위한 객체
				
				for (int k = 20;  k < 30; k++) {
					Cell cells1 = rowTemp_pic1.createCell(k);
					cells1.setCellStyle(picTop_style);
					cells1.setCellStyle(picTop_style);
				}
			
				sheet.addMergedRegion(new CellRangeAddress(
						third_st_pic, //first row (0-based)
						third_st_pic, //last row  (0-based)
				        20, //first column (0-based)
				        29  //last column  (0-based)
				));	
				third_st_pic = third_st_pic +1;
						
				Cell cell_ar = rowTemp_pic2.createCell(20);
				cell_ar.setCellStyle(picL_style);
				rowTemp_pic2.createCell(29).setCellStyle(picR_style);
				
				//하이퍼링크
				String cell_name = getCellName(cell_ar);
				Hyperlink link2 = createHelper.createHyperlink(Hyperlink.LINK_DOCUMENT);
				link2.setAddress("'"+sheet.getSheetName()+"'!"+cell_name);
				for (Row row : sheet_hyper) {
					if(row.getCell(pictureNoColumn_)!=null){
						Cell cell = row.getCell(pictureNoColumn_);
						if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
							if (cell.getNumericCellValue() == Integer.parseInt(picFileNameInExcel)) {
								cell.setHyperlink(link2);
								cell.setCellStyle(hlink_style);
							}
						}
					}
				}
				
				anchor.setCol1(21);	
				anchor.setRow1(third_st_pic);	// => A의 우측, 1행의 아래가 그림의 좌측상단 = 시작점이 된다.
				anchor.setCol2(29);					
				third_st_pic = third_st_pic+1;
				anchor.setRow2(third_st_pic);	// => D의 우측, 5행의 아래가 그림의 우측하단 = 끝점이 된다
								
				int pictureIndex = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);	// workbook에 추가하고, 추가한 그림의 id 같은걸 받아옴
						
				XSSFPicture picture = drawing.createPicture(anchor, pictureIndex);			// 좌표(anchor)와 그림id로 그림을 insert함	
				//( 참고 ) picture.resize();	// 원래 이미지 크기로 resize하는것
					
				for (int k = 20; k < 30; k++) {
					Cell cells1 = rowTemp_pic3.createCell(k);
					if(k==20){
						cells1.setCellStyle(picL_style);
					}else if(k==29){
						cells1.setCellStyle(picR_style);
					}
				}
						
				third_st_pic = third_st_pic +1;
				
				//출력Cell 생성		
				for (int k = 20; k < 30; k++) {
					Cell cells1 = rowTemp_pos.createCell(k);
					if(k == 20){
						cells1.setCellValue("위  치");
						cells1.setCellStyle(textheader_style);
					}else if(k == 22){
						cells1.setCellValue(position);
						cells1.setCellStyle(text_style);
					}else{
						cells1.setCellStyle(text_style);
					}
				}
					
				sheet.addMergedRegion(new CellRangeAddress(
						third_st_pic, //first row (0-based)
						third_st_pic, //last row  (0-based)
				        20, //first column (0-based)
				        21  //last column  (0-based)
				));	
				sheet.addMergedRegion(new CellRangeAddress(
						third_st_pic, //first row (0-based)
						third_st_pic, //last row  (0-based)
				        22, //first column (0-based)
				        29  //last column  (0-based)
				));
						
				third_st_pic = third_st_pic+1;
						
				for (int k = 20; k < 30; k++) {
					Cell cells1 = rowTemp_data.createCell(k);
					if(k == 20){
						cells1.setCellValue("내  용");
						cells1.setCellStyle(textheader_style);
					}else if(k == 22){
						cells1.setCellValue(content);
						cells1.setCellStyle(text_style);
					}else if(k == 26){
						cells1.setCellValue(supply+" / "+unit+" / "+ea+"EA");
						cells1.setCellStyle(text_style);
					}else{
						cells1.setCellStyle(text_style);
					}					
				}
						
				sheet.addMergedRegion(new CellRangeAddress(
						third_st_pic, //first row (0-based)
						third_st_pic, //last row  (0-based)
				        20, //first column (0-based)
				        21  //last column  (0-based)
				));	
				sheet.addMergedRegion(new CellRangeAddress(
						third_st_pic, //first row (0-based)
						third_st_pic, //last row  (0-based)
				        22, //first column (0-based)
				        25  //last column  (0-based)
				));
				sheet.addMergedRegion(new CellRangeAddress(
						third_st_pic, //first row (0-based)
						third_st_pic, //last row  (0-based)
				        26, //first column (0-based)
				        29  //last column  (0-based)
				));
				third_st_pic = third_st_pic+1;
			}
		}				
	}
	
	public void make_4(File pictureDir,XSSFWorkbook workbook ,XSSFSheet sheet ,Object sheets ,int pictureNoColumn_) throws Exception {
		//폰트스타일
		Font fontBody = workbook.createFont();
		fontBody.setColor(HSSFColor.BLACK.index);
		fontBody.setFontHeight((short)220);
		fontBody.setFontName("굴림체");

		//셀스타일
		CellStyle textheader_style = workbook.createCellStyle();
		textheader_style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		textheader_style.setAlignment(CellStyle.ALIGN_CENTER);
		textheader_style.setFont(fontBody);
		textheader_style.setBorderTop(CellStyle.BORDER_MEDIUM);                      
		textheader_style.setBorderLeft(CellStyle.BORDER_MEDIUM);
		textheader_style.setBorderRight(CellStyle.BORDER_MEDIUM);
		textheader_style.setBorderBottom(CellStyle.BORDER_MEDIUM);
				
		CellStyle text_style = workbook.createCellStyle();
		text_style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		text_style.setFont(fontBody);
		text_style.setIndention((short)1);
		text_style.setBorderTop(CellStyle.BORDER_MEDIUM);                      
		text_style.setBorderLeft(CellStyle.BORDER_MEDIUM);
		text_style.setBorderRight(CellStyle.BORDER_MEDIUM);
		text_style.setBorderBottom(CellStyle.BORDER_MEDIUM);

		CellStyle picTop_style = workbook.createCellStyle();
		picTop_style.setBorderLeft(CellStyle.BORDER_MEDIUM);
		picTop_style.setBorderTop(CellStyle.BORDER_MEDIUM);  
		picTop_style.setBorderRight(CellStyle.BORDER_MEDIUM);
		
		CellStyle picL_style = workbook.createCellStyle();
		picL_style.setBorderLeft(CellStyle.BORDER_MEDIUM);

		CellStyle picR_style = workbook.createCellStyle();
		picR_style.setBorderRight(CellStyle.BORDER_MEDIUM);

		//하이퍼링크관련
		CellStyle hlink_style = workbook.createCellStyle();
		Font hlink_font = workbook.createFont();
		hlink_font.setUnderline(Font.U_SINGLE);
		hlink_font.setColor(IndexedColors.BLUE.getIndex());
		hlink_style.setFont(hlink_font);
		hlink_style.setAlignment(CellStyle.ALIGN_CENTER);
		hlink_style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
				
		InputStream pictureFIS;
		Row rowTemp_pic1,rowTemp_pic2,rowTemp_pic3,rowTemp_pos,rowTemp_data;
		XSSFCell cellTemp;
			
		//사진대지
		sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
				
		Header pageHeader = sheet.getHeader();	//머릿말
		pageHeader.setCenter(HSSFHeader.font("휴먼옛체", "Normal") +HSSFHeader.fontSize((short) 26) + "사 진 대 지");
								
		//출력 row 생성
		rowTemp_pic1 = sheet.createRow(0);
		rowTemp_pic2 = sheet.createRow(0);
		rowTemp_pic3 = sheet.createRow(0);
				
		rowTemp_pos = sheet.createRow(0);
		rowTemp_data = sheet.createRow(0);
				
		List<DamageAndPicture> DamageAndPictureSheet = (List<DamageAndPicture>) sheets;

		CreationHelper createHelper = workbook.getCreationHelper();//하이퍼링크용
		
		String sheet_hyper_name = sheet.getSheetName();
		sheet_hyper_name = sheet_hyper_name.replaceAll("_사진", ""); //원본시트이름
		
		XSSFSheet sheet_hyper = workbook.getSheet(sheet_hyper_name); // 원본시트를 찾아감.?

		int odd_rowcount = 9;
				
		int odd_st_pic = 0;
		int even_st_pic = 0;
		int third_st_pic = 0;
		int fourth_st_pic = 0;
		
		for (int i = 0; i < DamageAndPictureSheet.size(); i++) {		
			if (i % 4 == 0){
				odd_rowcount =  odd_rowcount + 10;
				DamageAndPicture dmgStatePic = (DamageAndPicture)DamageAndPictureSheet.get(i);
				String picFileNameInExcel = dmgStatePic.getPictureFileNameInExcel();		
				String position = dmgStatePic.getPosition();
				String content = dmgStatePic.getContent();
				String supply = dmgStatePic.getSupply();
				String unit = dmgStatePic.getUnit();
				String ea = dmgStatePic.getEa();
			
				String basePath = pictureDir.getPath();
				File pictureFile = new File(basePath+"\\"+picFileNameInExcel+".jpg");			// 파일정보
				
				pictureFIS = new FileInputStream(pictureFile);				// InputStream에 파일 set		*FileNotFoundException -> try catch 있어야됨
				byte[] bytes = IOUtils.toByteArray(pictureFIS);				// 이미지 binary를 byte 배열에 담음
		
				CreationHelper helper = workbook.getCreationHelper();
				XSSFDrawing drawing = sheet.createDrawingPatriarch();	//그림 컨테이너, 그림을 실제로 insert하는 놈
				ClientAnchor anchor = helper.createClientAnchor();	// 그림을 넣을 좌표를 지정하기 위한 객체
						
				rowTemp_pic1 = sheet.createRow(odd_st_pic);
				rowTemp_pic1.setHeight((short)500);
						
				for (int k = 0;  k < 10; k++) {
					Cell cells1 = rowTemp_pic1.createCell(k);
					cells1.setCellStyle(picTop_style);
					cells1.setCellStyle(picTop_style);
				}
				sheet.addMergedRegion(new CellRangeAddress(
						odd_st_pic, //first row (0-based)
						odd_st_pic, //last row  (0-based)
				        0, //first column (0-based)
				        9  //last column  (0-based)
				));	
				
				odd_st_pic = odd_st_pic +1;
				rowTemp_pic2 = sheet.createRow(odd_st_pic);
				rowTemp_pic2.setHeight((short)4700);
				Cell cell_ar = rowTemp_pic2.createCell(0);
				cell_ar.setCellStyle(picL_style);
				rowTemp_pic2.createCell(9).setCellStyle(picR_style);
				
				//하이퍼링크
				String cell_name = getCellName(cell_ar);
				Hyperlink link2 = createHelper.createHyperlink(Hyperlink.LINK_DOCUMENT);
				link2.setAddress("'"+sheet.getSheetName()+"'!"+cell_name);
				for (Row row : sheet_hyper) {
					if(row.getCell(pictureNoColumn_)!=null){
						Cell cell = row.getCell(pictureNoColumn_);
						if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
							if (cell.getNumericCellValue() == Integer.parseInt(picFileNameInExcel)) {
								cell.setHyperlink(link2);
								cell.setCellStyle(hlink_style);
							}
						}
					}
				}
				
				anchor.setCol1(1);	
				anchor.setRow1(odd_st_pic);	// => A의 우측, 1행의 아래가 그림의 좌측상단 = 시작점이 된다.
				anchor.setCol2(9);					
				odd_st_pic = odd_st_pic+1;
				anchor.setRow2(odd_st_pic);	// => D의 우측, 5행의 아래가 그림의 우측하단 = 끝점이 된다
										
				int pictureIndex = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);	// workbook에 추가하고, 추가한 그림의 id 같은걸 받아옴
						
				XSSFPicture picture = drawing.createPicture(anchor, pictureIndex);			// 좌표(anchor)와 그림id로 그림을 insert함	
				//( 참고 ) picture.resize();	// 원래 이미지 크기로 resize하는것
						
				rowTemp_pic3 = sheet.createRow(odd_st_pic);
				rowTemp_pic3.setHeight((short)500);
			
				for (int k = 0; k < 10; k++) {
					Cell cells1 = rowTemp_pic3.createCell(k);
					if(k==0){
						cells1.setCellStyle(picL_style);
					}else if(k==9){
						cells1.setCellStyle(picR_style);
					}
				}
						
				odd_st_pic = odd_st_pic +1;
						
				//출력Cell 생성
				rowTemp_pos = sheet.createRow(odd_st_pic);
				rowTemp_pos.setHeight((short)500);
						
				for (int k = 0; k < 10; k++) {
					Cell cells1 = rowTemp_pos.createCell(k);
					if(k == 0){
						cells1.setCellValue("위  치");
						cells1.setCellStyle(textheader_style);
					}else if(k == 2){
						cells1.setCellValue(position);
						cells1.setCellStyle(text_style);
					}else{
						cells1.setCellStyle(text_style);
					}
				}
						
				sheet.addMergedRegion(new CellRangeAddress(
						odd_st_pic, //first row (0-based)
						odd_st_pic, //last row  (0-based)
				        0, //first column (0-based)
				        1  //last column  (0-based)
				));	
				sheet.addMergedRegion(new CellRangeAddress(
						odd_st_pic, //first row (0-based)
						odd_st_pic, //last row  (0-based)
				        2, //first column (0-based)
				        9  //last column  (0-based)
				));
				
				odd_st_pic = odd_st_pic+1;	
				rowTemp_data = sheet.createRow(odd_st_pic);
				rowTemp_data.setHeight((short)500);
				
				for (int k = 0; k < 10; k++) {
					Cell cells1 = rowTemp_data.createCell(k);
					if(k == 0){
						cells1.setCellValue("내  용");
						cells1.setCellStyle(textheader_style);
					}else if(k == 2){
						cells1.setCellValue(content);
						cells1.setCellStyle(text_style);
					}else if(k == 6){
						cells1.setCellValue(supply+" / "+unit+" / "+ea+"EA");
						cells1.setCellStyle(text_style);
					}else{
						cells1.setCellStyle(text_style);
					}					
				}
						
				sheet.addMergedRegion(new CellRangeAddress(
						odd_st_pic, //first row (0-based)
						odd_st_pic, //last row  (0-based)
				        0, //first column (0-based)
				        1  //last column  (0-based)
				));	
				sheet.addMergedRegion(new CellRangeAddress(
						odd_st_pic, //first row (0-based)
						odd_st_pic, //last row  (0-based)
				        2, //first column (0-based)
				        5  //last column  (0-based)
				));
				sheet.addMergedRegion(new CellRangeAddress(
						odd_st_pic, //first row (0-based)
						odd_st_pic, //last row  (0-based)
				        6, //first column (0-based)
				        9  //last column  (0-based)
				));
				odd_st_pic = odd_st_pic+1;
			}else if (i % 4 == 1){
				DamageAndPicture dmgStatePic = (DamageAndPicture)DamageAndPictureSheet.get(i);
				String picFileNameInExcel = dmgStatePic.getPictureFileNameInExcel();		
				String position = dmgStatePic.getPosition();
				String content = dmgStatePic.getContent();
				String supply = dmgStatePic.getSupply();
				String unit = dmgStatePic.getUnit();
				String ea = dmgStatePic.getEa();
					
				String basePath = pictureDir.getPath();
				File pictureFile = new File(basePath+"\\"+picFileNameInExcel+".jpg");			// 파일정보
				
				pictureFIS = new FileInputStream(pictureFile);				// InputStream에 파일 set		*FileNotFoundException -> try catch 있어야됨
				byte[] bytes = IOUtils.toByteArray(pictureFIS);				// 이미지 binary를 byte 배열에 담음
			
				CreationHelper helper = workbook.getCreationHelper();
				XSSFDrawing drawing = sheet.createDrawingPatriarch();	//그림 컨테이너, 그림을 실제로 insert하는 놈
				ClientAnchor anchor = helper.createClientAnchor();	// 그림을 넣을 좌표를 지정하기 위한 객체
				
				for (int k = 10;  k < 20; k++) {
					Cell cells1 = rowTemp_pic1.createCell(k);
					cells1.setCellStyle(picTop_style);
					cells1.setCellStyle(picTop_style);
				}
			
				sheet.addMergedRegion(new CellRangeAddress(
						even_st_pic, //first row (0-based)
						even_st_pic, //last row  (0-based)
				        10, //first column (0-based)
				        19  //last column  (0-based)
				));	
				even_st_pic = even_st_pic +1;
						
				Cell cell_ar = rowTemp_pic2.createCell(10);
				cell_ar.setCellStyle(picL_style);
				rowTemp_pic2.createCell(19).setCellStyle(picR_style);
				
				//하이퍼링크
				String cell_name = getCellName(cell_ar);
				Hyperlink link2 = createHelper.createHyperlink(Hyperlink.LINK_DOCUMENT);
				link2.setAddress("'"+sheet.getSheetName()+"'!"+cell_name);
				for (Row row : sheet_hyper) {
					if(row.getCell(pictureNoColumn_)!=null){
						Cell cell = row.getCell(pictureNoColumn_);
						if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
							if (cell.getNumericCellValue() == Integer.parseInt(picFileNameInExcel)) {
								cell.setHyperlink(link2);
								cell.setCellStyle(hlink_style);
							}
						}
					}
				}
				
				anchor.setCol1(11);	
				anchor.setRow1(even_st_pic);	// => A의 우측, 1행의 아래가 그림의 좌측상단 = 시작점이 된다.
				anchor.setCol2(19);					
				even_st_pic = even_st_pic+1;
				anchor.setRow2(even_st_pic);	// => D의 우측, 5행의 아래가 그림의 우측하단 = 끝점이 된다
								
				int pictureIndex = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);	// workbook에 추가하고, 추가한 그림의 id 같은걸 받아옴
						
				XSSFPicture picture = drawing.createPicture(anchor, pictureIndex);			// 좌표(anchor)와 그림id로 그림을 insert함	
				//( 참고 ) picture.resize();	// 원래 이미지 크기로 resize하는것
					
				for (int k = 10; k < 20; k++) {
					Cell cells1 = rowTemp_pic3.createCell(k);
					if(k==10){
						cells1.setCellStyle(picL_style);
					}else if(k==19){
						cells1.setCellStyle(picR_style);
					}
				}
						
				even_st_pic = even_st_pic +1;
				
				//출력Cell 생성		
				for (int k = 10; k < 20; k++) {
					Cell cells1 = rowTemp_pos.createCell(k);
					if(k == 10){
						cells1.setCellValue("위  치");
						cells1.setCellStyle(textheader_style);
					}else if(k == 12){
						cells1.setCellValue(position);
						cells1.setCellStyle(text_style);
					}else{
						cells1.setCellStyle(text_style);
					}
				}
					
				sheet.addMergedRegion(new CellRangeAddress(
						even_st_pic, //first row (0-based)
						even_st_pic, //last row  (0-based)
				        10, //first column (0-based)
				        11  //last column  (0-based)
				));	
				sheet.addMergedRegion(new CellRangeAddress(
						even_st_pic, //first row (0-based)
						even_st_pic, //last row  (0-based)
				        12, //first column (0-based)
				        19  //last column  (0-based)
				));
						
				even_st_pic = even_st_pic+1;
						
				for (int k = 10; k < 20; k++) {
					Cell cells1 = rowTemp_data.createCell(k);
					if(k == 10){
						cells1.setCellValue("내  용");
						cells1.setCellStyle(textheader_style);
					}else if(k == 12){
						cells1.setCellValue(content);
						cells1.setCellStyle(text_style);
					}else if(k == 16){
						cells1.setCellValue(supply+" / "+unit+" / "+ea+"EA");
						cells1.setCellStyle(text_style);
					}else{
						cells1.setCellStyle(text_style);
					}					
				}
						
				sheet.addMergedRegion(new CellRangeAddress(
						even_st_pic, //first row (0-based)
						even_st_pic, //last row  (0-based)
				        10, //first column (0-based)
				        11  //last column  (0-based)
				));	
				sheet.addMergedRegion(new CellRangeAddress(
						even_st_pic, //first row (0-based)
						even_st_pic, //last row  (0-based)
				        12, //first column (0-based)
				        15  //last column  (0-based)
				));
				sheet.addMergedRegion(new CellRangeAddress(
						even_st_pic, //first row (0-based)
						even_st_pic, //last row  (0-based)
				        16, //first column (0-based)
				        19  //last column  (0-based)
				));
				even_st_pic = even_st_pic+1;
			}else if (i % 4 == 2){
				
				DamageAndPicture dmgStatePic = (DamageAndPicture)DamageAndPictureSheet.get(i);
				String picFileNameInExcel = dmgStatePic.getPictureFileNameInExcel();		
				String position = dmgStatePic.getPosition();
				String content = dmgStatePic.getContent();
				String supply = dmgStatePic.getSupply();
				String unit = dmgStatePic.getUnit();
				String ea = dmgStatePic.getEa();
					
				String basePath = pictureDir.getPath();
				File pictureFile = new File(basePath+"\\"+picFileNameInExcel+".jpg");			// 파일정보
				
				pictureFIS = new FileInputStream(pictureFile);				// InputStream에 파일 set		*FileNotFoundException -> try catch 있어야됨
				byte[] bytes = IOUtils.toByteArray(pictureFIS);				// 이미지 binary를 byte 배열에 담음
			
				CreationHelper helper = workbook.getCreationHelper();
				XSSFDrawing drawing = sheet.createDrawingPatriarch();	//그림 컨테이너, 그림을 실제로 insert하는 놈
				ClientAnchor anchor = helper.createClientAnchor();	// 그림을 넣을 좌표를 지정하기 위한 객체
				
				for (int k = 20;  k < 30; k++) {
					Cell cells1 = rowTemp_pic1.createCell(k);
					cells1.setCellStyle(picTop_style);
					cells1.setCellStyle(picTop_style);
				}
			
				sheet.addMergedRegion(new CellRangeAddress(
						third_st_pic, //first row (0-based)
						third_st_pic, //last row  (0-based)
				        20, //first column (0-based)
				        29  //last column  (0-based)
				));	
				third_st_pic = third_st_pic +1;
						
				Cell cell_ar = rowTemp_pic2.createCell(20);
				cell_ar.setCellStyle(picL_style);
				rowTemp_pic2.createCell(29).setCellStyle(picR_style);
				
				//하이퍼링크
				String cell_name = getCellName(cell_ar);
				Hyperlink link2 = createHelper.createHyperlink(Hyperlink.LINK_DOCUMENT);
				link2.setAddress("'"+sheet.getSheetName()+"'!"+cell_name);
				for (Row row : sheet_hyper) {
					if(row.getCell(pictureNoColumn_)!=null){
						Cell cell = row.getCell(pictureNoColumn_);
						if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
							if (cell.getNumericCellValue() == Integer.parseInt(picFileNameInExcel)) {
								cell.setHyperlink(link2);
								cell.setCellStyle(hlink_style);
							}
						}
					}
				}
				
				anchor.setCol1(21);	
				anchor.setRow1(third_st_pic);	// => A의 우측, 1행의 아래가 그림의 좌측상단 = 시작점이 된다.
				anchor.setCol2(29);					
				third_st_pic = third_st_pic+1;
				anchor.setRow2(third_st_pic);	// => D의 우측, 5행의 아래가 그림의 우측하단 = 끝점이 된다
								
				int pictureIndex = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);	// workbook에 추가하고, 추가한 그림의 id 같은걸 받아옴
						
				XSSFPicture picture = drawing.createPicture(anchor, pictureIndex);			// 좌표(anchor)와 그림id로 그림을 insert함	
				//( 참고 ) picture.resize();	// 원래 이미지 크기로 resize하는것
					
				for (int k = 20; k < 30; k++) {
					Cell cells1 = rowTemp_pic3.createCell(k);
					if(k==20){
						cells1.setCellStyle(picL_style);
					}else if(k==29){
						cells1.setCellStyle(picR_style);
					}
				}
						
				third_st_pic = third_st_pic +1;
				
				//출력Cell 생성		
				for (int k = 20; k < 30; k++) {
					Cell cells1 = rowTemp_pos.createCell(k);
					if(k == 20){
						cells1.setCellValue("위  치");
						cells1.setCellStyle(textheader_style);
					}else if(k == 22){
						cells1.setCellValue(position);
						cells1.setCellStyle(text_style);
					}else{
						cells1.setCellStyle(text_style);
					}
				}
					
				sheet.addMergedRegion(new CellRangeAddress(
						third_st_pic, //first row (0-based)
						third_st_pic, //last row  (0-based)
				        20, //first column (0-based)
				        21  //last column  (0-based)
				));	
				sheet.addMergedRegion(new CellRangeAddress(
						third_st_pic, //first row (0-based)
						third_st_pic, //last row  (0-based)
				        22, //first column (0-based)
				        29  //last column  (0-based)
				));
						
				third_st_pic = third_st_pic+1;
						
				for (int k = 20; k < 30; k++) {
					Cell cells1 = rowTemp_data.createCell(k);
					if(k == 20){
						cells1.setCellValue("내  용");
						cells1.setCellStyle(textheader_style);
					}else if(k == 22){
						cells1.setCellValue(content);
						cells1.setCellStyle(text_style);
					}else if(k == 26){
						cells1.setCellValue(supply+" / "+unit+" / "+ea+"EA");
						cells1.setCellStyle(text_style);
					}else{
						cells1.setCellStyle(text_style);
					}					
				}
						
				sheet.addMergedRegion(new CellRangeAddress(
						third_st_pic, //first row (0-based)
						third_st_pic, //last row  (0-based)
				        20, //first column (0-based)
				        21  //last column  (0-based)
				));	
				sheet.addMergedRegion(new CellRangeAddress(
						third_st_pic, //first row (0-based)
						third_st_pic, //last row  (0-based)
				        22, //first column (0-based)
				        25  //last column  (0-based)
				));
				sheet.addMergedRegion(new CellRangeAddress(
						third_st_pic, //first row (0-based)
						third_st_pic, //last row  (0-based)
				        26, //first column (0-based)
				        29  //last column  (0-based)
				));
				third_st_pic = third_st_pic+1;
			}else{
				
				DamageAndPicture dmgStatePic = (DamageAndPicture)DamageAndPictureSheet.get(i);
				String picFileNameInExcel = dmgStatePic.getPictureFileNameInExcel();		
				String position = dmgStatePic.getPosition();
				String content = dmgStatePic.getContent();
				String supply = dmgStatePic.getSupply();
				String unit = dmgStatePic.getUnit();
				String ea = dmgStatePic.getEa();
					
				String basePath = pictureDir.getPath();
				File pictureFile = new File(basePath+"\\"+picFileNameInExcel+".jpg");			// 파일정보
				
				pictureFIS = new FileInputStream(pictureFile);				// InputStream에 파일 set		*FileNotFoundException -> try catch 있어야됨
				byte[] bytes = IOUtils.toByteArray(pictureFIS);				// 이미지 binary를 byte 배열에 담음
			
				CreationHelper helper = workbook.getCreationHelper();
				XSSFDrawing drawing = sheet.createDrawingPatriarch();	//그림 컨테이너, 그림을 실제로 insert하는 놈
				ClientAnchor anchor = helper.createClientAnchor();	// 그림을 넣을 좌표를 지정하기 위한 객체
				
				for (int k = 30;  k < 40; k++) {
					Cell cells1 = rowTemp_pic1.createCell(k);
					cells1.setCellStyle(picTop_style);
					cells1.setCellStyle(picTop_style);
				}
			
				sheet.addMergedRegion(new CellRangeAddress(
						fourth_st_pic, //first row (0-based)
						fourth_st_pic, //last row  (0-based)
				        30, //first column (0-based)
				        39  //last column  (0-based)
				));	
				fourth_st_pic = fourth_st_pic +1;
						
				Cell cell_ar = rowTemp_pic2.createCell(30);
				cell_ar.setCellStyle(picL_style);
				rowTemp_pic2.createCell(39).setCellStyle(picR_style);
				
				//하이퍼링크
				String cell_name = getCellName(cell_ar);
				Hyperlink link2 = createHelper.createHyperlink(Hyperlink.LINK_DOCUMENT);
				link2.setAddress("'"+sheet.getSheetName()+"'!"+cell_name);
				for (Row row : sheet_hyper) {
					if(row.getCell(pictureNoColumn_)!=null){
						Cell cell = row.getCell(pictureNoColumn_);
						if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
							if (cell.getNumericCellValue() == Integer.parseInt(picFileNameInExcel)) {
								cell.setHyperlink(link2);
								cell.setCellStyle(hlink_style);
							}
						}
					}
				}
				
				anchor.setCol1(31);	
				anchor.setRow1(fourth_st_pic);	// => A의 우측, 1행의 아래가 그림의 좌측상단 = 시작점이 된다.
				anchor.setCol2(39);					
				fourth_st_pic = fourth_st_pic+1;
				anchor.setRow2(fourth_st_pic);	// => D의 우측, 5행의 아래가 그림의 우측하단 = 끝점이 된다
								
				int pictureIndex = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);	// workbook에 추가하고, 추가한 그림의 id 같은걸 받아옴
						
				XSSFPicture picture = drawing.createPicture(anchor, pictureIndex);			// 좌표(anchor)와 그림id로 그림을 insert함	
				//( 참고 ) picture.resize();	// 원래 이미지 크기로 resize하는것
					
				for (int k = 30; k < 40; k++) {
					Cell cells1 = rowTemp_pic3.createCell(k);
					if(k==30){
						cells1.setCellStyle(picL_style);
					}else if(k==39){
						cells1.setCellStyle(picR_style);
					}
				}
						
				fourth_st_pic = fourth_st_pic +1;
				
				//출력Cell 생성		
				for (int k = 30; k < 40; k++) {
					Cell cells1 = rowTemp_pos.createCell(k);
					if(k == 30){
						cells1.setCellValue("위  치");
						cells1.setCellStyle(textheader_style);
					}else if(k == 32){
						cells1.setCellValue(position);
						cells1.setCellStyle(text_style);
					}else{
						cells1.setCellStyle(text_style);
					}
				}
					
				sheet.addMergedRegion(new CellRangeAddress(
						fourth_st_pic, //first row (0-based)
						fourth_st_pic, //last row  (0-based)
				        30, //first column (0-based)
				        31  //last column  (0-based)
				));	
				sheet.addMergedRegion(new CellRangeAddress(
						fourth_st_pic, //first row (0-based)
						fourth_st_pic, //last row  (0-based)
				        32, //first column (0-based)
				        39  //last column  (0-based)
				));
						
				fourth_st_pic = fourth_st_pic+1;
						
				for (int k = 30; k < 40; k++) {
					Cell cells1 = rowTemp_data.createCell(k);
					if(k == 30){
						cells1.setCellValue("내  용");
						cells1.setCellStyle(textheader_style);
					}else if(k == 32){
						cells1.setCellValue(content);
						cells1.setCellStyle(text_style);
					}else if(k == 36){
						cells1.setCellValue(supply+" / "+unit+" / "+ea+"EA");
						cells1.setCellStyle(text_style);
					}else{
						cells1.setCellStyle(text_style);
					}					
				}
						
				sheet.addMergedRegion(new CellRangeAddress(
						fourth_st_pic, //first row (0-based)
						fourth_st_pic, //last row  (0-based)
				        30, //first column (0-based)
				        31  //last column  (0-based)
				));	
				sheet.addMergedRegion(new CellRangeAddress(
						fourth_st_pic, //first row (0-based)
						fourth_st_pic, //last row  (0-based)
				        32, //first column (0-based)
				        35  //last column  (0-based)
				));
				sheet.addMergedRegion(new CellRangeAddress(
						fourth_st_pic, //first row (0-based)
						fourth_st_pic, //last row  (0-based)
				        36, //first column (0-based)
				        39  //last column  (0-based)
				));
				fourth_st_pic = fourth_st_pic+1;
			}
		}		
	}
	
	private static String getCellName(Cell cell)
	{
	    return CellReference.convertNumToColString(cell.getColumnIndex()) + (cell.getRowIndex() + 1);
	}
}
