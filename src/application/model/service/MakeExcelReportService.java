package application.model.service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFHeader;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import application.model.async.CallbackChain;
import application.model.async.GenericEvent;
import application.model.async.ServiceEventListener;
import application.model.vo.DamageAndPicture;
import application.model.vo.Work;

/**
 * �������� ����
 * 
 */
public class MakeExcelReportService implements CallbackChain, Runnable{

	Work currentWork;
	CallbackChain callbackInstance;
	private ServiceEventListener servicEventListener;
	
	public MakeExcelReportService(Work work, CallbackChain callbackInstance){
		this.currentWork = work;
		this.callbackInstance = callbackInstance;
		
	}
	
	/**
	 * �������� ����, �Ǻ����̺��� �����Ѵ�.
	 */
	public void makeExcelReport(){

		if(servicEventListener != null ){	servicEventListener.happen(new GenericEvent(GenericEvent.ET_MAKE_EXCEL_START, "��� ������ ���� ���Դϴ�."));	}
		
		List<List<DamageAndPicture>> multSheet = this.currentWork.getDamageAndPicturesOnMultiSheets();
		FileInputStream fs_inputExcel;

		ExcelImage excelImage = new ExcelImage();
		ExcelPivot excelPivot = new ExcelPivot();

		try {

			fs_inputExcel = new FileInputStream(this.currentWork.getInputExcel());

			XSSFWorkbook workbook = new XSSFWorkbook(fs_inputExcel); 	

			if(servicEventListener != null ){	servicEventListener.happen(new GenericEvent(GenericEvent.ET_MAKE_EXCEL_ON, "������ �������Դϴ�."));	}
			
			for (int j = 0; j < multSheet.size(); j++) {
				
				Object sheets = multSheet.get(j);
				String sheet_name = workbook.getSheetName(j);
				XSSFSheet sheet = workbook.createSheet(sheet_name+"_����"); //��������
				
				//��������
				Header pageHeader = sheet.getHeader();	//�Ӹ���
				pageHeader.setCenter(HSSFHeader.font("�޸տ�ü", "Normal") +HSSFHeader.fontSize((short) 26) + "�� �� �� ��");
				
				switch (this.currentWork.getSelectedPrintType()) {//��½� ���������κ��� �뿭�� ������� ���ϴºκ�
				case "1": //1��
					sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
					excelImage.make_1(this.currentWork.getInputPictureDir(), workbook, sheet, sheets
							, ReadInputExcelService.decodeColumnStringToDecimalIndex (this.currentWork.getPictureNoColumn()) );
					
					int data_st_pic1 = sheet.getLastRowNum();
					
					int dats1 = workbook.getSheetIndex(sheet_name+"_����");
					
					workbook.setPrintArea(
							dats1, //sheet index
							0, //start column
							9, //end column
							0, //start row
							data_st_pic1 //end row
					);
					sheet.setDisplayGridlines(true);
				    sheet.setPrintGridlines(true);
					break;
				case "2": //2��
					sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
					excelImage.make_2(this.currentWork.getInputPictureDir(), workbook, sheet, sheets
							, ReadInputExcelService.decodeColumnStringToDecimalIndex (this.currentWork.getPictureNoColumn()) );
					
					int data_st_pic2 = sheet.getLastRowNum();	
					int dats2 = workbook.getSheetIndex(sheet_name+"_����");
					
					workbook.setPrintArea(
							dats2, //sheet index
							0, //start column
							19, //end column
							0, //start row
							data_st_pic2 //end row
					);			
					sheet.setDisplayGridlines(true);
				    sheet.setPrintGridlines(true);
					
					break;
				case "3": //3��
					sheet.getPrintSetup().setPaperSize(PrintSetup.A3_PAPERSIZE);
					excelImage.make_3(this.currentWork.getInputPictureDir(), workbook, sheet, sheets
							, ReadInputExcelService.decodeColumnStringToDecimalIndex (this.currentWork.getPictureNoColumn()) );
					
					int data_st_pic3 = sheet.getLastRowNum();
					int dats3 = workbook.getSheetIndex(sheet_name+"_����");
					
					workbook.setPrintArea(
							dats3, //sheet index
							0, //start column
							29, //end column
							0, //start row
							data_st_pic3 //end row
					);	
					sheet.setDisplayGridlines(true);
				    sheet.setPrintGridlines(true);
					
					break;
				case "4": //4��
					sheet.getPrintSetup().setPaperSize(PrintSetup.A3_PAPERSIZE);
					excelImage.make_4(this.currentWork.getInputPictureDir(), workbook, sheet, sheets
							, ReadInputExcelService.decodeColumnStringToDecimalIndex (this.currentWork.getPictureNoColumn()) );	
					
					int data_st_pic4 = sheet.getLastRowNum();
					int dats4 = workbook.getSheetIndex(sheet_name+"_����");
					
					workbook.setPrintArea(
							dats4, //sheet index
							0, //start column
							39, //end column
							0, //start row
							data_st_pic4 //end row
					);	
					sheet.setDisplayGridlines(true);
				    sheet.setPrintGridlines(true);
					
					break;
				default:
					break;
				}
				//�������� END
				
				//�ǹ�
				if(servicEventListener != null ){	servicEventListener.happen(new GenericEvent(GenericEvent.ET_MAKE_EXCEL_ON, "�ǹ����̺��� �������Դϴ�."));	}
				excelPivot.make_pivot(workbook, sheet_name, this.currentWork.getPivot1NoColumn(), this.currentWork.getPivot2NoColumn(), j);
				//�ǹ� END
			}
            
			File outputExcelFile = new File(this.currentWork.getOutputTargetDir().getPath() + "\\" + "��������"+this.currentWork.getSelectedPrintType()+"��.xlsx");
			FileOutputStream out = new FileOutputStream(outputExcelFile);
			workbook.write(out);
            out.close();
            
            if(servicEventListener != null ){	servicEventListener.happen(new GenericEvent(GenericEvent.ET_MAKE_EXCEL_ON, "�������� ������ �Ϸ��߽��ϴ�."));	}

		} catch (Exception e) {
			if(servicEventListener != null ){	servicEventListener.happen(new GenericEvent(GenericEvent.ET_MAKE_EXCEL_ERR, "������ �߻��߽��ϴ�."));	}
			e.printStackTrace();
			StringWriter sw = new StringWriter();
			e.printStackTrace(new PrintWriter(sw));
			String exceptionAsString = sw.toString();

			ExceptionCheck exx = new ExceptionCheck();
			try {
				exx.ExceptionCall(exceptionAsString);
				
			} catch (Exception e1) {
				e1.printStackTrace();
				
			}
		}		
	}
	
	
	@Override
	public void run() {
		makeExcelReport();
		
		if(callbackInstance != null){
			callbackInstance.callback();
			
		}
		
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
