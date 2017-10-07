package application.controllers;

import java.io.File;
import java.io.IOException;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;
import java.util.ResourceBundle;

import application.model.async.ServiceEventListener;
import application.model.service.MainExcelReportService;
import application.model.vo.DamageAndPicture;
import application.model.vo.Work;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.fxml.Initializable;
import javafx.scene.Scene;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.RadioButton;
import javafx.scene.control.TableCell;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.TextField;
import javafx.scene.control.ToggleGroup;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.layout.AnchorPane;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import javafx.util.Callback;
/**
 * main.fxml�� ���� ��Ʈ�ѷ�
 * 
 * */
public class MainController implements Initializable{
	
	@FXML private AnchorPane rootElement;
	@FXML private Button setInputExcelButton;
	@FXML private Label excelPathLabel;
	@FXML private TextField positionColumnTextField;
	@FXML private TextField contentColumnTextField;
	@FXML private TextField pictureNoColumnTextField;
	@FXML private TextField pivot1NoColumnTextField;
	@FXML private TextField pivot2NoColumnTextField;
	@FXML private Button setPicDirButton;
	@FXML private Label pictureDirPathLabel;
	@FXML private Button previewButton;
	@FXML private TableView previewTableView;
	@FXML private Button executeButton;
	@FXML private ToggleGroup outputTypeToggleGroup;

	public static Work CURRENT_WORK;
	private MainExcelReportService mainExcelReportService;

	@Override
	public void initialize(URL arg0, ResourceBundle arg1){
		CURRENT_WORK = new Work(); //���ø����̼� ���� �� �� �۾��� ����
		mainExcelReportService = new MainExcelReportService(CURRENT_WORK);//������ �а�, ���������� �����ϴ� ���� ���μ����� �����ϴ� ���� ��ü

		setInputExcelButton.setOnMouseClicked(event -> {
			FileChooser fileChooser = new FileChooser();
			Stage currentStage = (Stage) setInputExcelButton.getScene().getWindow();
			File inputExcel = (fileChooser.showOpenDialog(currentStage));
			
			mainExcelReportService.setInputExcel(inputExcel);
			
			excelPathLabel.setText(inputExcel.getAbsolutePath());
			
		});
		
		
		setPicDirButton.setOnMouseClicked(event -> {
			DirectoryChooser dirChooser = new DirectoryChooser();
			Stage currentStage = (Stage) setPicDirButton.getScene().getWindow();
			File inputPictureDir = (dirChooser.showDialog(currentStage));
			
			mainExcelReportService.setInputPictureDir(inputPictureDir);
			
			pictureDirPathLabel.setText(inputPictureDir.getAbsolutePath());
			
		});
		
		
		previewButton.setOnMouseClicked(event -> {
			String positionColumn 	=  positionColumnTextField.getText();
			String contentColumn 	=  contentColumnTextField.getText();
			String pictureNoColumn 	=  pictureNoColumnTextField.getText();
			mainExcelReportService.setColumnPositions(positionColumn, contentColumn, pictureNoColumn);
			
			String pivot1NoColumn = pivot1NoColumnTextField.getText();
			String pivot2NoColumn = pivot2NoColumnTextField.getText();
			mainExcelReportService.setPivotPositions(pivot1NoColumn, pivot2NoColumn);

			mainExcelReportService.checkBeforeReadExcel();
			
			ProgressEventController progressEventController = popupProgressEvent();
			ServiceEventListener serviceEventListener = new ServiceEventListener(progressEventController);
			Thread listenerThread = new Thread(serviceEventListener);	//eventListner
			listenerThread.start();
			
			mainExcelReportService.previewReport(() ->{
				showPreview();
				serviceEventListener.setRunning(false);
				progressEventController.closeWindow();

			});

		});
		
		executeButton.setOnMouseClicked(event -> {
			Alert alert = new Alert(AlertType.INFORMATION);
			alert.setTitle("����");
			alert.setHeaderText(null);
			alert.setContentText("��� ������ ������ ������ ������ �ּ���.");
			alert.showAndWait();
			
			//�ƿ�ǲ ������ ������ ���丮�� �����Ѵ�.
			DirectoryChooser dirChooser = new DirectoryChooser();
			Stage currentStage = (Stage) executeButton.getScene().getWindow();
			File outputTargetDir = dirChooser.showDialog(currentStage);
			if(outputTargetDir == null ) {return;}
			CURRENT_WORK.setOutputTargetDir(outputTargetDir);
			
			//��������� �о�� �����Ѵ�.
			RadioButton selectedRB = (RadioButton) outputTypeToggleGroup.getSelectedToggle();
			String selectedPrintType = selectedRB.getUserData().toString();
			CURRENT_WORK.setSelectedPrintType(selectedPrintType);

			mainExcelReportService.checkBeforeMakeExcel();

			ProgressEventController progressEventController = popupProgressEvent();
			ServiceEventListener serviceEventListener = new ServiceEventListener(progressEventController);
			Thread listenerThread = new Thread(serviceEventListener);
			listenerThread.start();
			
			mainExcelReportService.setServiceEventListener(serviceEventListener);
			mainExcelReportService.makeExcelReport(() ->{
				serviceEventListener.setRunning(false);
				progressEventController.endProgress();
			});

		});

	}
	
	/**
	 * �ε�â�� ����. �ε�â�� Controller�� �����Ѵ�.
	 * @return
	 */
	public ProgressEventController popupProgressEvent(){
		Stage ProgressEventstage = new Stage();
		FXMLLoader loader = new FXMLLoader();
		
		loader.setLocation(getClass().getResource("/application/resources/fxml/progressEvent.fxml"));
		try {
			Scene scene = new Scene(loader.load());
			ProgressEventstage.setScene(scene);
			ProgressEventstage.show();
			return loader.getController();
			
		} catch (IOException e) {
			System.out.println("progressEvent.fxml�� �о���µ� �����Ͽ����ϴ�.");
			e.printStackTrace();
			return null;
		}

	}
	
	/**
	 * �̸����� ����� �������Ѵ�.
	 */
	private void showPreview(){
		List<List<DamageAndPicture>> multSheets =  CURRENT_WORK.getDamageAndPicturesOnMultiSheets();
		ObservableList<DamageAndPicture> dataList = FXCollections.observableArrayList();
		
		ObservableList<TableColumn> colLi = previewTableView.getColumns();
		TableColumn sheetCol = colLi.get(0);
		TableColumn positionCol = colLi.get(1);	// ��ġ : 0
		TableColumn contentCol = colLi.get(2);	//������ȣ : 1
		TableColumn pictureNoCol = colLi.get(3);
		TableColumn pictureFile = colLi.get(4);
		
		for(int listCnt = 0 ; listCnt < multSheets.size(); listCnt++){
			Object sheets = multSheets.get(listCnt);
			List<DamageAndPicture> dmgStateAndPictureSheet = (List<DamageAndPicture>) sheets;
			
			List check_pic_num = new ArrayList<>();
			//boolean check_img = true; // ���� �ߺ�üũ�� ���� ���� �ּ�
			
			/*checkPictureFileIsExists(dmgStateAndPictureSheet); // ������ �׸����� ������ �ش��ϴ� ���ϸ��� �׸������� �ִ��� Ȯ���Ѵ�. �ش� ������ fullname�� ����´�.
			HashMap<String, List<DamageAndPicture>> dupObjs = getDSPsDuplicatedOnPictureNumber(dmgStateAndPictureSheet); */
			
			sheetCol.setCellValueFactory(new PropertyValueFactory<DamageAndPicture,String>("sheetnum"));
			positionCol.setCellValueFactory(new PropertyValueFactory<DamageAndPicture,String>("position"));
			pictureNoCol.setCellValueFactory(new PropertyValueFactory<DamageAndPicture,String>("pictureFileNameInExcel"));
			contentCol.setCellValueFactory(new PropertyValueFactory<DamageAndPicture,String>("content"));
			pictureFile.setCellValueFactory(new PropertyValueFactory<DamageAndPicture,String>("pictureFile"));
			
			pictureNoCol.setCellFactory(new Callback<TableColumn<String, String>, TableCell<String, String>>() {
	            @Override
	            public TableCell call(TableColumn p) {
	                return new TableCell<String, String>() {
	                    @Override
	                    public void updateItem(final String item, final boolean empty) {
	                        super.updateItem(item, empty);//*don't forget!
	                        if (item != null) {
	                            setText(item);
	                            if (item.startsWith("�ߺ�")) {
	                                setStyle("-fx-background-color: red; -fx-text-fill: white;");
	                            }else{
	                            	setStyle("");
	                            }
	                        } else {
	                            setText(null);
	                        }
	                    }
	                };
	            }
	        });
			
			for(DamageAndPicture dmgStatPic  : dmgStateAndPictureSheet){
				dataList.add(dmgStatPic);
					
				//���� �ߺ� üũ
				String check_number = dmgStatPic.getPictureFileNameInExcel().toString()+Integer.toString(dmgStatPic.getSheetnum());
				
				if(check_pic_num.contains(check_number)){
					dmgStatPic.setPictureFileNameInExcel("�ߺ�/"+dmgStatPic.getPictureFileNameInExcel().toString());
					//check_img = false; // ���� �ߺ�üũ��
				}else{
					check_pic_num.add(check_number);
				}
				
			}
		
			previewTableView.setItems(dataList);	

		}// END for listCnt
		
		return;
		
	}

}
