package application.model.service;

import java.io.File;

import application.model.async.CallbackChain;
import application.model.async.ServiceEventListener;
import application.model.vo.Work;

/**
 * ���� ���� �۾��� ���� ��� �޼ҵ尡 �ִ� ��ü
 * 
 * */
public class MainExcelReportService {
	
	private Work currentWork;
	private ServiceEventListener serviceEventListener; 
	
	public MainExcelReportService(Work work) {
		this.currentWork = work;

	}
	
	
	/**
	 * �Է� ���������� currentWork�� �����Ѵ�.
	 * @param inputExcel
	 * @return
	 */
	public void setInputExcel (File inputExcel) {
		checkInputExcelException(inputExcel);
		currentWork.setInputExcel(inputExcel);
		
	}
	
	/**
	 * �Էµ� ���������� �������� �������� üũ�Ѵ�.
	 * @param inputExcel
	 */
	private boolean checkInputExcelException(File inputExcel) {
		String file_root = inputExcel.getAbsolutePath();
		String[] directoryName = file_root.split("\\\\");
		String fileName = directoryName[directoryName.length -1];
		if(inputExcel != null){
			if(!fileName.contains(".xlsx")){
				ExceptionCheck exx = new ExceptionCheck();
				try {
					exx.ExceptionCall("xlsxȮ���� ���ϸ� ��ȯ�� �����մϴ�.\n �̿��� ������ ��ȯ�ؼ� �־��ֽñ� �ٶ��ϴ�.");
					inputExcel = null;
					return false;
				} catch (Exception e1) {
					e1.printStackTrace();
					
				}
				
			}
			
		}
		return true;
	}
	

	/**
	 * �Է� �׸������� currentWork�� �����Ѵ�.
	 * @param inputPictureDir
	 */
	public void setInputPictureDir(File inputPictureDir) {
		currentWork.setInputPictureDir(inputPictureDir);
		
	}
	
	/**
	 * �Էµ� �÷���ġ�� currentWork�� �����Ѵ�.
	 * @param positionColumn_
	 * @param contentColumn_
	 * @param pictureNoColumn_
	 */
	public void setColumnPositions(String positionColumn, String contentColumn, String pictureNoColumn) {
		currentWork.setPositionColumn(positionColumn);
		currentWork.setContentColumn(contentColumn);
		currentWork.setPictureNoColumn(pictureNoColumn);
		
	}
	
	/**
	 * �Էµ� �Ǻ� ��ġ�� currentWork�� �����Ѵ�.
	 * @param a
	 * @param b
	 */
	public void setPivotPositions(String pivot1NoColumn, String pivot2NoColumn) {
		currentWork.setPivot1NoColumn(pivot1NoColumn);
		currentWork.setPivot2NoColumn(pivot2NoColumn);
		
	}

	/**
	 * �̸����� ��ư ����
	 * */
	public void previewReport(CallbackChain callbackInstance) {
		/*
		 * �޼ҵ� ���� ����
		 * 1. readInputExcelService.run()
		 * 2. 1�� ������ callbackInstance -> { showPreview(); ... }
		 * */
		ReadInputExcelService readInputExcelService = new ReadInputExcelService(this.currentWork, callbackInstance);
		readInputExcelService.setServicEventListener(this.serviceEventListener);
		
		Thread readInputExcelThread = new Thread(readInputExcelService);
		readInputExcelThread.start();
		
	}
	
	/**
	 * ������ư ����
	 * */
	public void makeExcelReport(CallbackChain callbackInstance) {
		/*
		 * �޼ҵ� ���� ����
		 * 1. readInputExcelService.run()
		 * 2. 1�� ������ makeExcelReportService.run()
		 * 3. 2�� ������ callbackInstance  -> { showPreview(); ... }
		 * */
		MakeExcelReportService makeExcelReportService = new MakeExcelReportService(this.currentWork, callbackInstance);
		makeExcelReportService.setServicEventListener(this.serviceEventListener);	//this.serviceEventObserver�� Controller ���� MainExcelReportService.setService...���� ����
		
		ReadInputExcelService readInputExcelService = new ReadInputExcelService(this.currentWork, makeExcelReportService);
		readInputExcelService.setServicEventListener(this.serviceEventListener);

		Thread readInputExcelThread = new Thread(readInputExcelService);
		readInputExcelThread.start();

	}

	/**
	 * �Է¿����� �б����� ��ȿ�� �˻縦 �Ѵ�.
	 * 
	 */
	public void checkBeforeReadExcel() {
		//TODO
	}

	/**
	 * ��¿����� ����� �� ��ȿ�� �˻縦 �Ѵ�.
	 */
	public void checkBeforeMakeExcel() {
		//TODO
	}

	
	public ServiceEventListener getServiceEventListener() {
		return serviceEventListener;
		
	}

	public void setServiceEventListener(ServiceEventListener serviceEventListener) {
		this.serviceEventListener = serviceEventListener;
		
	}



}
