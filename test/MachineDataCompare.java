import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Scanner;

import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import jxl.Cell;
import jxl.CellType;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;


public class MachineDataCompare {

	static Scanner sc;
	
	

	public static void main (String [] args) throws IOException{
		//Declaring variables
		boolean again = true;
		String modelNumber;
		String screwDiameter;
		int modelColomnNumber = -1, colomnNumber = -1;
		ArrayList <String> specSheetValues = new ArrayList<String>();
		ArrayList <String> specSheetVariableNames = new ArrayList<String>();
		ArrayList <String> xmlValues = new ArrayList<String>();
		ArrayList <String> xmlVariableNames = new ArrayList<String>();
		JFileChooser chooser;
		sc = new Scanner (System.in);

		
		
		//Spec sheet portion
		while (again){
			
			String path;
			chooser = new JFileChooser();
			chooser.setCurrentDirectory(new File("."));


			System.out.println("Select the specsheet");
			int r = chooser.showOpenDialog(new JFrame());
			//path = chooser.getSelectedFile().getName();

			File specSheet = chooser.getSelectedFile().getAbsoluteFile();
			System.out.println(chooser.getSelectedFile().getAbsolutePath());

			System.out.println("Select the machine data");
			r = chooser.showOpenDialog(new JFrame());
			
			File fXmlFile = chooser.getSelectedFile().getAbsoluteFile();
			System.out.println(chooser.getSelectedFile().getAbsolutePath());

			System.out.print("Model Number: ");
			modelNumber = sc.nextLine();

			System.out.print("Screw Diameter: ");
			screwDiameter = sc.nextLine();

			System.out.println();

			Workbook w;
			try{
				w = Workbook.getWorkbook(specSheet);
				Sheet sheet = w.getSheet(2);
				//Checking for Model Number
				for (int i = 0; i < sheet.getColumns(); i++){
					Cell cell = sheet.getCell(i, 2);
					CellType type = cell.getType();
					if (modelNumber.equals(cell.getContents())){
						modelColomnNumber = i;
					}
				}
				System.out.println();

				//Checking for screw diameter
				if (modelColomnNumber != -1){
					for (int i = 0; i < 4; i++){
						Cell cell = sheet.getCell(modelColomnNumber + i, 8);
						if (cell.getContents().equals(screwDiameter)){
							colomnNumber = modelColomnNumber + i;
						}
					}
				}

				//Printing out values
				if (colomnNumber != -1 && modelColomnNumber != -1){
					for (int i = 462; i < sheet.getRows(); i++){

						Cell varCell = sheet.getCell(1, i);
						Cell rawCell = sheet.getCell(modelColomnNumber, i);
						Cell cell = sheet.getCell(colomnNumber, i);
						CellType type = cell.getType();


						if (!cell.getContents().equals("") && varCell.getContents().matches(".*[a-zA-Z]+.*")){
							if (cell.getContents().equals("YES"))
								specSheetValues.add("true");
							else
								specSheetValues.add(cell.getContents());
							specSheetVariableNames.add(varCell.getContents());
						}
						else if (!rawCell.getContents().equals("") && varCell.getContents().matches(".*[a-zA-Z]+.*")){
							specSheetValues.add(rawCell.getContents());
							specSheetVariableNames.add(varCell.getContents());
						}
					}
				}


				if (colomnNumber == -1){
					System.out.println("**********************************");
					System.out.println("Invalid Model Number or Screw Size");
					System.out.println("**********************************");
				}

			}catch(BiffException e){
				e.printStackTrace();
			}


			//XML portion

			try {

				
				DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
				DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
				Document doc = dBuilder.parse(fXmlFile);

				//optional, but recommended
				//read this - http://stackoverflow.com/questions/13786607/normalization-in-dom-parsing-with-java-how-does-it-work
				doc.getDocumentElement().normalize();

				System.out.println("Root element :" + doc.getDocumentElement().getNodeName());

				NodeList nList = doc.getElementsByTagName("Variable");

				System.out.println("----------------------------");

				for (int temp = 0; temp < nList.getLength(); temp++) {

					Node nNode = nList.item(temp);



					Element eElement = (Element) nNode;

					xmlVariableNames.add(eElement.getAttribute("Name"));
					xmlValues.add(eElement.getElementsByTagName("Value").item(0).getTextContent());

				}
			} catch (Exception e) {
				e.printStackTrace();
			}

			//Comparing both

			for (int i = 0; i < specSheetValues.size(); i++){
				for (int j = 0; j < xmlValues.size(); j++){
					if (specSheetVariableNames.get(i).equals(xmlVariableNames.get(j))){
						if (specSheetValues.get(i).matches(".*[a-zA-Z]+.*")){
							if (specSheetValues.get(i).equals(xmlValues.get(j))){
								System.out.println("Correct");
								System.out.println("----------------------------------------------------");
							}
							else{
								System.out.println(specSheetVariableNames.get(i));
								System.out.println("Spec Sheet:   " + specSheetValues.get(i));
								System.out.println("Machine Data: " + xmlValues.get(j));
								System.out.println("----------------------------------------------------");
							}
						}
						else{
							if (Double.parseDouble(specSheetValues.get(i).replaceAll(",", "")) == Double.parseDouble(xmlValues.get(j).replaceAll(",", ""))){
								System.out.println("Correct");
								System.out.println("----------------------------------------------------");
							}
							else{
								System.out.println(specSheetVariableNames.get(i));
								System.out.println("Spec Sheet:   " + specSheetValues.get(i));
								System.out.println("Machine Data: " + xmlValues.get(j));
								System.out.println("----------------------------------------------------");
							}
						}
						//												System.out.println(specSheetValues.get(i) + " Compared to " + xmlValues.get(j));
						//												System.out.println(specSheetVariableNames.get(i));
						//												System.out.println();
					}
				}
			}


			//Asking to perform the process again
			System.out.println("Would you like to do this again? (Yes/No)");
			if (sc.nextLine().equalsIgnoreCase("no"))
				again = false;

		}
	}
}
