package application.model.service;

import java.io.File;
import java.io.FileInputStream;

import application.Main;
import javafx.application.Platform;
import javafx.scene.Scene;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.TextArea;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.layout.VBox;
import javafx.scene.text.Text;
import javafx.stage.Stage;

public class ExceptionCheck {

	 public void ExceptionCall(String e) throws Exception{
		Platform.runLater(()->{
			Stage primaryStage = new Stage();
	        VBox box = new VBox();

		    Scene scene = new Scene(box, 900, 500);
		    
		    String currentWorkingDir = Main.initialWorkingDir;
		    
		    File imageFile = new File(currentWorkingDir+"\\appdata\\mainimage.jpg");
			Image main_image = null;
			try{
				main_image = new Image(new FileInputStream(imageFile));
			}catch(Exception ee){
				System.out.println("[Error]�ε��̹��� ������ ã�� �� �����ϴ�." + currentWorkingDir.toString()+"\\appdata\\mainimage.jpg");
				System.out.println(ee.getMessage());
			}
		    
		    ImageView view_image = new ImageView(main_image);
		    
	        box.getChildren().add(0,view_image);
		 
	        Text text_main = new Text("\n������ �߻��� ��� �Ʒ��� ������ �̸��Ϸ� ÷�����ֽñ�ٶ��ϴ�.\n");
		    text_main.setStyle("-fx-font-size: 20;");
	        
		    box.getChildren().add(1,text_main);
		    	    
		    
		    ScrollPane root = new ScrollPane();
		    
		    TextArea textArea = new TextArea();
		    box.getChildren().add(2, textArea );
	        textArea.setStyle("-fx-font-size: 15;");
	        textArea.setText(e);
	        textArea.deselect();
		    primaryStage.setTitle("Error");
	        primaryStage.setScene(scene);
	        primaryStage.show();	  
		});
		  
	 }
}
