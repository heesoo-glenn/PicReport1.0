package application.model.service;

import java.io.File;

import application.model.async.CallbackChain;
import application.model.async.ServiceEventListener;
import application.model.vo.Work;

/**
 * 엑셀 보고서 작업에 대한 모든 메소드가 있는 객체
 * 
 * */
public class MainExcelReportService {
	
	private Work currentWork;
	private ServiceEventListener serviceEventListener; 
	
	public MainExcelReportService(Work work) {
		this.currentWork = work;

	}
	
	
	/**
	 * 입력 엑셀파일을 currentWork에 세팅한다.
	 * @param inputExcel
	 * @return
	 */
	public void setInputExcel (File inputExcel) {
		checkInputExcelException(inputExcel);
		currentWork.setInputExcel(inputExcel);
		
	}
	
	/**
	 * 입력된 엑셀파일이 정상적인 형식인지 체크한다.
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
					exx.ExceptionCall("xlsx확장자 파일만 변환이 가능합니다.\n 이외의 형식은 변환해서 넣어주시기 바랍니다.");
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
	 * 입력 그림폴더를 currentWork에 세팅한다.
	 * @param inputPictureDir
	 */
	public void setInputPictureDir(File inputPictureDir) {
		currentWork.setInputPictureDir(inputPictureDir);
		
	}
	
	/**
	 * 입력된 컬럼위치를 currentWork에 세팅한다.
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
	 * 입력된 피봇 위치를 currentWork에 세팅한다.
	 * @param a
	 * @param b
	 */
	public void setPivotPositions(String pivot1NoColumn, String pivot2NoColumn) {
		currentWork.setPivot1NoColumn(pivot1NoColumn);
		currentWork.setPivot2NoColumn(pivot2NoColumn);
		
	}

	/**
	 * 미리보기 버튼 동작
	 * */
	public void previewReport(CallbackChain callbackInstance) {
		/*
		 * 메소드 동작 순서
		 * 1. readInputExcelService.run()
		 * 2. 1이 끝나면 callbackInstance -> { showPreview(); ... }
		 * */
		ReadInputExcelService readInputExcelService = new ReadInputExcelService(this.currentWork, callbackInstance);
		readInputExcelService.setServicEventListener(this.serviceEventListener);
		
		Thread readInputExcelThread = new Thread(readInputExcelService);
		readInputExcelThread.start();
		
	}
	
	/**
	 * 생성버튼 동작
	 * */
	public void makeExcelReport(CallbackChain callbackInstance) {
		/*
		 * 메소드 동작 순서
		 * 1. readInputExcelService.run()
		 * 2. 1이 끝나면 makeExcelReportService.run()
		 * 3. 2가 끝나면 callbackInstance  -> { showPreview(); ... }
		 * */
		MakeExcelReportService makeExcelReportService = new MakeExcelReportService(this.currentWork, callbackInstance);
		makeExcelReportService.setServicEventListener(this.serviceEventListener);	//this.serviceEventObserver는 Controller 에서 MainExcelReportService.setService...으로 주입
		
		ReadInputExcelService readInputExcelService = new ReadInputExcelService(this.currentWork, makeExcelReportService);
		readInputExcelService.setServicEventListener(this.serviceEventListener);

		Thread readInputExcelThread = new Thread(readInputExcelService);
		readInputExcelThread.start();

	}

	/**
	 * 입력엑셀을 읽기전에 유효성 검사를 한다.
	 * 
	 */
	public void checkBeforeReadExcel() {
		//TODO
	}

	/**
	 * 출력엑셀을 만들기 전 유효성 검사를 한다.
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
