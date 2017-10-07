package application.model.async;

import application.controllers.ProgressEventController;
import javafx.application.Platform;


/**
 * 옵저버 패턴 활용
 * 각각 다른 Service에서 이벤트 객체를 담아 여기의 happen을 호출한다.
 * EventObsesrver는 EventController를 제어하여 로딩창을 변경한다.
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
	 * false로 set하면 쓰레드를 멈춘다.
	 * @return
	 */
	public void setRunning(boolean isRunning) {
		this.isRunning = isRunning;
	}
	
}
