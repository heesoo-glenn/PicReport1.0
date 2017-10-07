package application.controllers;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.net.URL;
import java.util.List;
import java.util.ResourceBundle;

import application.Main;
import application.model.async.CallbackChain;
import application.model.async.GenericEvent;
import application.model.async.ServiceEventListener;
import javafx.application.Platform;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.fxml.Initializable;
import javafx.scene.Scene;
import javafx.scene.control.Label;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.layout.AnchorPane;
import javafx.stage.Stage;

/**
 * progressEvent.fxml�� ��Ʈ�ѷ�
 * 
 * @author leehs
 *
 */
public class ProgressEventController implements Initializable{

	@FXML ImageView progressImageView;
	@FXML Label progressText;
	@FXML AnchorPane rootElement;
	
	@Override
	public void initialize(URL location, ResourceBundle resources) {
		setProgressImage();
		progressText.setText("�۾��� ���۵Ǿ����ϴ�.");
	}
	
	public void setProgressImage(){
		Image loadingImage = getLoadingImage();
		progressImageView.setImage(loadingImage);
		
	}

	public void setProgressText(String text){
		Platform.runLater(() -> {
			progressText.setText(text);
			
		});
	}

	public void endProgress(){
		Platform.runLater(() -> {
			progressText.setText("�۾��� �Ϸ�Ǿ����ϴ�.");
			progressImageView.setImage(null);
			
		});
	}
	
	public  void closeWindow(){
		Platform.runLater(() -> {
			( (Stage)rootElement.getScene().getWindow() ).close();
		});
	}
	
	private Image getLoadingImage() {
		String currentWorkingDir = Main.initialWorkingDir;
		File imageFile = new File(currentWorkingDir+"\\appdata\\loading.gif");
		Image gifImage = null;
		try{
			gifImage = new Image(new FileInputStream(imageFile));
			
		}catch(Exception e){
			System.out.println("[Error]�ε��̹��� ������ ã�� �� �����ϴ�." + currentWorkingDir.toString()+"\\appdata\\loading.gif");
			System.out.println(e.getMessage());
			
		}
		return gifImage;

	}



	
}
