package application.model.async;

/**
 * PicReport에 쓰이는 Event
 * 
 * */
public class GenericEvent {
	
	public final static String ET_READ_EXCEL_START = "ET_READ_EXCEL_START";
	public final static String ET_READ_EXCEL_ON = "ET_READ_EXCEL_ON";
	public final static String ET_READ_EXCEL_ERR = "ET_READ_EXCEL_ERR";
	public final static String ET_READ_EXCEL_END = "ET_READ_EXCEL_END";
	
	public final static String ET_MAKE_EXCEL_START = "ET_READ_EXCEL_END";
	public final static String ET_MAKE_EXCEL_ON = "ET_READ_EXCEL_END";
	public final static String ET_MAKE_EXCEL_ERR = "ET_READ_EXCEL_ERR";
	public final static String ET_MAKE_EXCEL_END = "ET_READ_EXCEL_END";
	
	
	
	
	private String eventType;
	private String message;
	
	/**
	 * eventType 이벤트 유형
	 * message 기타 메세지
	 * */
	public GenericEvent(String eventType, String message){
		this.setEventType(eventType);
		this.setMessage(message);
	}

	public String getMessage() {
		return message;
	}

	public void setMessage(String message) {
		this.message = message;
	}

	public String getEventType() {
		return eventType;
	}

	public void setEventType(String eventType) {
		this.eventType = eventType;
	}
	
	
}
