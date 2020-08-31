package App;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFPictureData;
import org.junit.Assert;


public class ComparacaoDeExcel {

	public void verificacaoSeOsDadosSaoIguaisEmAmbasPlaninhas(Workbook workbook1, Workbook workbook2) {
		System.out.println("Verificando se ambos os Excel têm os mesmos tipos de dados:");
		
		// Conta abas
		int contaAbas = workbook1.getNumberOfSheets();
		
		// Contando as abas e extraindo o nome dela
		for(int i = 0; i < contaAbas; i++) {
			// Obtendo o indice de cada aba de ambas as planilhas
			Sheet linhasPlaninha1 = workbook1.getSheetAt(i);
			Sheet linhasPlaninha2 = workbook2.getSheetAt(i);
			System.out.println("----------- Aba: " + linhasPlaninha1.getSheetName() + " -----------");
			
			// Extraindo a quantidade de linhas da planilha
			int contagemDeLinhas = linhasPlaninha1.getPhysicalNumberOfRows();
			// Contando celulas
			for(int j = 0; j < contagemDeLinhas; j++) {
				
				int contagemDeCelulas = 0;
				// Tratando linhas nulas
				if(linhasPlaninha1.getRow(j) == null) {
					// Imprimindo a a linha nula
					System.out.println("** Número da linha nula: " + j + " **");
				}else {
					contagemDeCelulas = linhasPlaninha1.getRow(j).getPhysicalNumberOfCells();
				}
				// Extraindo cada celula por linha
				for(int k = 0; k < contagemDeCelulas; k++) {
					// Pegando celulas individualmente e tratando celulas nulas
					Cell celulasDaPlaninha1 = linhasPlaninha1.getRow(j).getCell(k, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					Cell celulasDaPlaninha2 = linhasPlaninha2.getRow(j).getCell(k, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					
					// Capturando os tipos de celulas
					if(celulasDaPlaninha1.getCellType().equals(celulasDaPlaninha2.getCellType())) {
						if(celulasDaPlaninha1.getCellType() == CellType.STRING) {
							String v1 = celulasDaPlaninha1.getStringCellValue();
							String v2 = celulasDaPlaninha2.getStringCellValue();
							Assert.assertEquals(v1, v2);
							System.out.println("Comparação - Planinha(Modelo) " + v1 + " === " + v2 + " Planinha(Teste)");
						}
						if(celulasDaPlaninha1.getCellType() == CellType.NUMERIC) {
							// Verificando se os dados também não são do tipo data
							// Se forem
							if(DateUtil.isCellDateFormatted(celulasDaPlaninha1) | DateUtil.isCellDateFormatted(celulasDaPlaninha2)) {
								// Formatando o conteudo
								DataFormatter df = new DataFormatter();
								String v1 = df.formatCellValue(celulasDaPlaninha1);
								String v2 = df.formatCellValue(celulasDaPlaninha2);
								Assert.assertEquals(v1, v2);
								System.out.println("Comparação - Planinha(Modelo) " + v1 + " === " + v2 + " Planinha(Teste)");
							} else {
								double v1 = celulasDaPlaninha1.getNumericCellValue();
								double v2 = celulasDaPlaninha2.getNumericCellValue();
								if(v1 == v2) System.out.println("Comparação - Planinha(Modelo) " + v1 + " === " + v2 + "Planinha(Teste)");;
							}
						}
						if(celulasDaPlaninha1.getCellType() == CellType.BOOLEAN) {
							boolean v1 = celulasDaPlaninha1.getBooleanCellValue();
							boolean v2 = celulasDaPlaninha2.getBooleanCellValue();
							Assert.assertEquals(v1, v2);
							System.out.println("Comparação - Planinha(Modelo) " + v1 + " === " + v2 + " Planinha(Teste)");
						}
					}
				}
			}
		}
		System.out.println("TESTE FINALIZADO COM SUCESSO!!");
	}
}
