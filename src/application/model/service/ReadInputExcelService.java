package application.model.service;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import application.model.async.CallbackChain;
import application.model.async.GenericEvent;
import application.model.async.ServiceEventListener;
import application.model.vo.DamageAndPicture;
import application.model.vo.Work;


/***
 * �񵿱�� ������ �о���� ��ü
 * @author
 *
 */
public class ReadInputExcelService implements Runnable, CallbackChain{
	
	Work currentWork;
	CallbackChain callbackInstance;
	private ServiceEventListener servicEventListener;
	
	public ReadInputExcelService(Work currentWork, CallbackChain callbackInstance) {
		this.currentWork = currentWork;
		this.callbackInstance = callbackInstance;
		
	}

	public void readExcel() {
		if(servicEventListener != null ){	servicEventListener.happen(new GenericEvent(GenericEvent.ET_READ_EXCEL_START, "�Էµ� ���������� �д� ���Դϴ�."));	}
		
        File inputExcel 		= currentWork.getInputExcel();
        
        String positionColumn 	= currentWork.getPositionColumn();						//��ġ�÷� ��ġ
    	String contentColumn 	= currentWork.getContentColumn();						//�����÷� ��ġ
    	String pictureNoColumn 	= currentWork.getPictureNoColumn();						//������ȣ�÷� ��ġ
    	int positionColNo = decodeColumnStringToDecimalIndex(positionColumn);
    	int contentColNo = decodeColumnStringToDecimalIndex(contentColumn);
    	int pictureNoColNo = decodeColumnStringToDecimalIndex(pictureNoColumn);
    	
    	Workbook wb = null;
    	
        try {
        	wb = new XSSFWorkbook(new FileInputStream(inputExcel));
        	List<List<DamageAndPicture>> allReadData = new ArrayList<List<DamageAndPicture>>();
        	
			int sheet_number = wb.getNumberOfSheets();
		
			Cell cell = null;
			String cellValue_3_tmp = "";   

			for (int i = 0; i < sheet_number; i++) {// sheet ���� �ݺ�    
				List<DamageAndPicture> readDataOnCurrentSheet = new ArrayList<DamageAndPicture>();
				for (Row row : wb.getSheetAt(i)) {
					//�� �б� 
					 String cellValue = readCellAsString(row.getCell(contentColNo)); //��������
					 String cellValue2 = readCellAsString(row.getCell(pictureNoColNo)); //������ȣ
					 String cellValue3 = readCellAsString(row.getCell(positionColNo)); //�氣
					 String cellValue3_1 = readCellAsString(row.getCell((positionColNo)+1)); //����
					 String cellValue4 = readCellAsString(row.getCell((pictureNoColNo)-3));; //����
					 String cellValue5 = readCellAsString(row.getCell((pictureNoColNo)-2));; //����
					 String cellValue6 = readCellAsString(row.getCell((pictureNoColNo)-4));; //����
					 String celldata1 = cellValue.replaceAll("\\p{Z}", "");
					 String celldata2 = cellValue2.replaceAll("\\p{Z}", "");
					 String celldata_tmp = cellValue3_1.replaceAll("\\p{Z}", "");
					 cellValue_3_tmp = celldata_tmp;
					 celldata_tmp = cellValue3.replaceAll("\\p{Z}", "");
					 cellValue_3_tmp = cellValue_3_tmp+ "("+celldata_tmp+")";
					 
					 String celldata_sup = cellValue4.replaceAll("\\p{Z}", "");
					 String celldata_unit = cellValue5.replaceAll("\\p{Z}", "");
					 String celldata_ea = cellValue6.replaceAll("\\p{Z}", "");
					 
					 
					 if(!celldata1.equalsIgnoreCase("null") && !celldata1.equalsIgnoreCase("") && !celldata1.startsWith("����") &&
						!celldata2.equalsIgnoreCase("null") && !celldata2.equalsIgnoreCase("") && !celldata2.startsWith("����") &&
						!celldata_sup.equalsIgnoreCase("null") && !celldata_sup.equalsIgnoreCase("") && !celldata_sup.startsWith("����") &&
						!cellValue_3_tmp.equalsIgnoreCase("null") && !cellValue_3_tmp.equalsIgnoreCase("") && !cellValue_3_tmp.startsWith("����") &&
						!celldata_unit.equalsIgnoreCase("null") && !celldata_unit.equalsIgnoreCase("") && !celldata_unit.startsWith("����")&&
						!celldata_ea.equalsIgnoreCase("null") && !celldata_ea.equalsIgnoreCase("") && !celldata_ea.startsWith("����")
						){
						 readDataOnCurrentSheet.add(new DamageAndPicture(cellValue_3_tmp, cellValue, cellValue2, celldata_sup, celldata_unit,celldata_ea,i+1));	
						 //cellValue = ����, cellValue2 = picNO?, cellValue_3_tmp = ��ġ, celldata_sup =���� , celldata_unit = ���� ,celldata_ea = ����
						 //(String position, String content, String pictureFileNameInExcel)
					 }

				}
				allReadData.add(readDataOnCurrentSheet);
				
			}
			
			currentWork.setDamageAndPicturesOnMultiSheets(allReadData);
			
			if(servicEventListener != null ){	servicEventListener.happen(new GenericEvent(GenericEvent.ET_READ_EXCEL_END, "�Էµ� ������ ��� �о����ϴ�."));	}
			
		}catch (Exception e) {
			if(servicEventListener != null ){	servicEventListener.happen(new GenericEvent(GenericEvent.ET_READ_EXCEL_ERR, "������ �߻��߽��ϴ�."));	}
			e.printStackTrace();
			StringWriter sw = new StringWriter();
			e.printStackTrace(new PrintWriter(sw));
			String exceptionAsString = sw.toString();

			ExceptionCheck exx = new ExceptionCheck();
			try {
				exx.ExceptionCall(exceptionAsString);
			} catch (Exception e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
				
			}
		}finally{
			if (wb != null) {
		        try {
		            wb.close();
		        } catch (IOException ex) {
		            // ignore ... any significant errors should already have been
		            // reported via an IOException from the final flush.
		        }
		    };
		}
      
		return;       
	}
	
	
	/**
	 * �� ������ ������ Ȯ���ϰ� ������ string ������ ��ȯ��.
	 * @param cell
	 * @return
	 */
	private String readCellAsString(Cell cell) {
		 String valueStr = "";
		 
		 if(cell != null){
			 switch(cell.getCellType()){
				case Cell.CELL_TYPE_STRING :
					valueStr = cell.getStringCellValue();
					break;
				case Cell.CELL_TYPE_NUMERIC : // ��¥ �����̵� ���� �����̵� �� CELL_TYPE_NUMERIC���� �ν���.
					if(DateUtil.isCellDateFormatted(cell)){ // ��¥ ������ �������� ���,
						SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd", Locale.KOREA);
						String formattedStr = dateFormat.format(cell.getDateCellValue());
						valueStr = formattedStr;
						break;
					}else{ // �����ϰ� ���� �������� ���,
						Double numericCellValue = cell.getNumericCellValue();
						if(Math.floor(numericCellValue) == numericCellValue){ // �Ҽ��� ���ϸ� ���� ���� ������ ���� ���ٸ�,,
							valueStr = numericCellValue.intValue() + ""; // int������ �Ҽ��� ���� ������ String���� ������ ��´�.
						}else{
							valueStr = numericCellValue + "";
						}
						break;
					}
				case Cell.CELL_TYPE_BOOLEAN :
					valueStr = cell.getBooleanCellValue() + "";
					break;
				case Cell.CELL_TYPE_ERROR :
					valueStr = cell.getBooleanCellValue() + "";
					break;
				case Cell.CELL_TYPE_FORMULA :
					switch(cell.getCachedFormulaResultType()) {
		            case Cell.CELL_TYPE_NUMERIC:
		            	valueStr = String.format("%.2f",cell.getNumericCellValue()); 
		                break;
		            case Cell.CELL_TYPE_STRING:
		            	RichTextString data = cell.getRichStringCellValue();
		            	valueStr = data.toString();
		                break;
					}
					break;
				default:
					break;
			}
		}
		return valueStr;		
	}

	
	/**
	 * �÷� ���ĺ��� ���� �ε����� �����ϴ� �޼ҵ�
	 * */
	public static int decodeColumnStringToDecimalIndex(String columnStr) {
		columnStr.toUpperCase();
		char[] eachChar = new char[columnStr.length()];
		columnStr.getChars(0, columnStr.length(), eachChar, 0);
		
		int columnNo = 0;
		for(int i=0; i < eachChar.length; i++){
			int charInt = eachChar[i];
			charInt -=65;
			columnNo += charInt*Math.pow(26, i);
		}

		return columnNo;
	}

	@Override
	public void run() {
		readExcel();
		
		if(callbackInstance != null){
			callbackInstance.callback();
		}
		
		return;
	}

	@Override
	public void callback() {
		run();
		
	}

	public ServiceEventListener getServicEventListener() {
		return servicEventListener;
	}

	public void setServicEventListener(ServiceEventListener servicEventListener) {
		this.servicEventListener = servicEventListener;
	}


}
