import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.*;
import java.io.*;

//https://www.youtube.com/watch?v=RsrF2Ku7ad4
//https://poi.apache.org/spreadsheet/quick-guide.html
public class reader {

	public static void main(String args[]) throws InvalidFormatException, IOException{
		System.out.println("JEDEM");
	
			//Workbook workbook = jxl.Workbook.getWorkbook(new File("Book1.xls")); 
			File objekt = new File("src/Book1.xlsx");
			System.out.println(objekt.exists());
			HSSFWorkbook novy = new HSSFWorkbook(); //
			try {
				Workbook stary = WorkbookFactory.create(new File("src/Book1.xlsx"));
				
				Sheet sheet1 = stary.getSheetAt(0);
				
				System.out.println("**************");
				System.out.println("NAMESPACE");
				int pocet = stary.getNumberOfNames();
				for(int h=0;h<pocet;h++){
					Name jmeno = stary.getNameAt(h);
					String formule = jmeno.getRefersToFormula();
					String nazev = jmeno.getNameName();

					System.out.println(nazev + " = " + formule);
				}
				System.out.println("**************");
				
				
				
				int start = sheet1.getFirstRowNum();
				int konec = sheet1.getLastRowNum();
				for(int i=start;i<konec;i++){
					Row radek = sheet1.getRow(i);
					System.out.println("############");
					if(radek == null){

						System.out.println("RADEK CHYBI!!!");
						continue;
					}
					
					int prvaBunka = radek.getFirstCellNum();
					int posledniBunka = radek.getLastCellNum();
					
					//prvni a posledni
					//System.out.println(prvaBunka);
					//System.out.println(posledniBunka);
					
					for(int j=prvaBunka;j<posledniBunka;j++){
						Cell bunka = radek.getCell(j);
						
						if(bunka == null){
							continue;
						}
						
							int typ = bunka.getCellType();
							if(typ == 0){
								double cislo = bunka.getNumericCellValue();
								System.out.println(cislo);
							}
							if(typ == 1){
								Object x = bunka.getRichStringCellValue();

								System.out.println(x);
							}
							if(typ == 2){
								int typx = bunka.getCachedFormulaResultType();
								if(typx == 0){
								double cislo = bunka.getNumericCellValue();
								System.out.println(cislo);}
								
								
							}
					
						
						
					}//konec prochazeni bunek
					
						
					
					
					
				}//konec prochazeni radku
				
				//PREPSANI TRETIHO SLOUPCE NA SAME TROJKY A ZMENA GLOBALNI PROMENNE
				System.out.println("//////////////////");
				Name jmeno = stary.getNameAt(0);
				String formule = jmeno.getRefersToFormula();
				String formuleNova = formule + "6";
				System.out.println(formule + " stara");
				
				String nazev = jmeno.getNameName();
				stary.removeName(0);
				Name noveJmeno = stary.createName();
				noveJmeno.setRefersToFormula(formuleNova);
				noveJmeno.setNameName(nazev);
				
				jmeno = stary.getNameAt(0);
				formule = jmeno.getRefersToFormula();
				System.out.println(formule + " nova");
				
				//zmena globalni promenne splneno
				sheet1 = stary.getSheetAt(0);
				for(int i=1;i<8;i++){
					Row radek = sheet1.getRow(i);
					Cell bunkaZmenit = radek.getCell(3);
					bunkaZmenit.setCellValue(3);
					System.out.println("Zmena" + i);
				}
				
				//prepsani tretiho sloupce splneno
				
				//ULOZENI VYSLEDKU DO NOVEHO SOUBORU
				File vysledny = new File("vysledek.xls");
				vysledny.createNewFile();
				
			    FileOutputStream fileOut = new FileOutputStream(vysledny);
			    stary.write(fileOut);
			    fileOut.close();
			    //vysledek uspesne otevren v MS excel a overen, konec cviceni


				 
				
				
				
				
				
				
				
				
				
				
				
				
				
				
				
			} catch (EncryptedDocumentException e) {
				
				e.printStackTrace();
			}
			
			
	
		
	
	
		
	}
	
	
	
	
	
	
	
	
	
}
