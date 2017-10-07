package application.model.async;

import application.controllers.ProgressEventController;
import javafx.application.Platform;


/**
 * ������ ���� Ȱ��
 * ���� �ٸ� Service���� �̺�Ʈ ��ü�� ��� ������ happen�� ȣ���Ѵ�.
 * EventObsesrver�� EventController�� �����Ͽ� �ε�â�� �����Ѵ�.
 */
public class ServiceEventListener implements Runnable{
	
	ProgressEventController progressEventController;
	
	private boolean isRunning = true;
	
	public ServiceEventListener(ProgressEventController progressEventController) {
		this.progressEventController = progressEventController;
	}
	
	public void happen(GenericEvent event){
		handle(event);
		
	}

	public void handle(GenericEvent event) {
		Platform.runLater(()->{
			progressEventController.setProgressText(event.getMessage());
		});
		
	}

	@Override
	public void run() {
		running();
	}
	
	private void running() {
		while(true){
			if(!isRunning()){
				return;
			}
		}
	}
	
	
	public boolean isRunning() {
		return isRunning;
	}
	
	/***
	 * false�� set�ϸ� �����带 �����.
	 * @return
	 */
	public void setRunning(boolean isRunning) {
		this.isRunning = isRunning;
	}
	
}
