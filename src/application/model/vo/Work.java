package application.model.vo;

import java.io.File;
import java.util.HashMap;
import java.util.List;

/**
 * 
 * ��Ŭ������ ġ�� �ϳ��� ������Ʈ�� �ش��ϴ� ��ü
 * "�۾�"
 * ���������� �����ϱ����� �ջ���Ȳ ����, �Է¿��� ���Ͽ� ���� ���� ���� ����ȴ�.
 * 
 * inputExcel : �ջ���Ȳ ��������
 * List  damageAndPictures : �������Ͽ��� �о�帰 �ջ���Ȳ java object ������
 * 
 * pictureDir : �׸�����
 * List pictures : �׸�����
 * 
 * 
 * */
public class Work {
	private File inputExcel;							//�Է� ����
	//private List<DamageAndPicture> damageAndPictures;	//�ջ���Ȳ VO
	private List<List<DamageAndPicture>> damageAndPicturesOnMultiSheets;

	private File inputPictureDir;						//�׸�����
	private List<File> inputPictures;					//�׸�����
	

	private String positionColumn;						//��ġ�÷� ��ġ
	private String contentColumn;						//�����÷� ��ġ
	private String pictureNoColumn;						//������ȣ�÷� ��ġ
	
	private String pivot1NoColumn;						//�ǹ� ���� �� ��ġ 1
	private String pivot2NoColumn;						//�ǹ� ���� �� ��ġ 2
	
	private String selectedPrintType;					//���Ÿ��

	private File outputTargetDir;						//������� ���� ����
	
	
	
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
