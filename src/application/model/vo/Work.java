package application.model.vo;

import java.io.File;
import java.util.HashMap;
import java.util.List;

/**
 * 
 * 이클립스로 치면 하나의 프로젝트에 해당하는 객체
 * "작업"
 * 사진대지를 생성하기위한 손상현황 정보, 입력엑셀 파일에 대한 정보 등이 저장된다.
 * 
 * inputExcel : 손상현황 엑셀파일
 * List  damageAndPictures : 엑셀파일에서 읽어드린 손상현황 java object 데이터
 * 
 * pictureDir : 그림폴더
 * List pictures : 그림파일
 * 
 * 
 * */
public class Work {
	private File inputExcel;							//입력 엑셀
	//private List<DamageAndPicture> damageAndPictures;	//손상현황 VO
	private List<List<DamageAndPicture>> damageAndPicturesOnMultiSheets;

	private File inputPictureDir;						//그림폴더
	private List<File> inputPictures;					//그림파일
	

	private String positionColumn;						//위치컬럼 위치
	private String contentColumn;						//내용컬럼 위치
	private String pictureNoColumn;						//사진번호컬럼 위치
	
	private String pivot1NoColumn;						//피벗 참조 셀 위치 1
	private String pivot2NoColumn;						//피벗 참조 셀 위치 2
	
	private String selectedPrintType;					//출력타입

	private File outputTargetDir;						//출력파일 떨굴 폴더
	
	
	
	public File getInputExcel() {
		return inputExcel;
	}
	public void setInputExcel(File inputExcel) {
		this.inputExcel = inputExcel;
	}
	public List<File> getInputPictures() {
		return inputPictures;
	}
	public void setInputPictures(List<File> inputPictures) {
		this.inputPictures = inputPictures;
	}

/*	public List<DamageAndPicture> getDamageAndPictures() {
		return damageAndPictures;
	}
	public void setDamageAndPictures(List<DamageAndPicture> damageAndPictures) {
		this.damageAndPictures = damageAndPictures;
	}*/
	public File getInputPictureDir() {
		return inputPictureDir;
	}
	public void setInputPictureDir(File inputPictureDir) {
		this.inputPictureDir = inputPictureDir;
	}
	public String getContentColumn() {
		return contentColumn;
	}
	public void setContentColumn(String contentColumn) {
		this.contentColumn = contentColumn;
	}
	public String getPositionColumn() {
		return positionColumn;
	}
	public void setPositionColumn(String positionColumn) {
		this.positionColumn = positionColumn;
	}
	public String getPictureNoColumn() {
		return pictureNoColumn;
	}
	public void setPictureNoColumn(String pictureNoColumn) {
		this.pictureNoColumn = pictureNoColumn;
	}
	public String getPivot1NoColumn() {
		return pivot1NoColumn;
	}
	public void setPivot1NoColumn(String pivot1NoColumn) {
		this.pivot1NoColumn = pivot1NoColumn;
	}
	public String getPivot2NoColumn() {
		return pivot2NoColumn;
	}
	public void setPivot2NoColumn(String pivot2NoColumn) {
		this.pivot2NoColumn = pivot2NoColumn;
	}
	public List<List<DamageAndPicture>> getDamageAndPicturesOnMultiSheets() {
		return damageAndPicturesOnMultiSheets;
	}
	public void setDamageAndPicturesOnMultiSheets(List<List<DamageAndPicture>> damageAndPicturesOnMultiSheets) {
		this.damageAndPicturesOnMultiSheets = damageAndPicturesOnMultiSheets;
	}
	public String getSelectedPrintType() {
		return selectedPrintType;
	}
	public void setSelectedPrintType(String selectedPrintType) {
		this.selectedPrintType = selectedPrintType;
	}
	public File getOutputTargetDir() {
		return outputTargetDir;
	}
	public void setOutputTargetDir(File outputTargetDir) {
		this.outputTargetDir = outputTargetDir;
	}

	
	
	
	
}
