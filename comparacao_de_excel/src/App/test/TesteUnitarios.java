package App.test;

import java.io.File;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Test;

import App.ComparacaoDeExcel;

public class TesteUnitarios {

	@Test
	public void extraindoImagensDasPlaninhas() throws EncryptedDocumentException, IOException {
		String userDir = System.getProperty("user.dir");
		Workbook wb1 = WorkbookFactory.create(new File(userDir + "arquivo2.xlsx"));
		Workbook wb2 = WorkbookFactory.create(new File(userDir + "arquivo1.xlsx"));
		ComparacaoDeExcel comparacaoDeExcel = new ComparacaoDeExcel();
		comparacaoDeExcel.extraindoImagensDaPlaninha(wb1, wb2);
	}
}
