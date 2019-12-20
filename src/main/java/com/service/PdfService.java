package com.service;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileSystemView;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.fasterxml.jackson.annotation.JsonInclude.Include;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.service.Category;
import com.service.FieldType;

public class PdfService {
	final static Logger logger = Logger.getLogger(PdfService.class);

	public PdfService() {
		super();
	}

	/**
	 * Create json with Ihm swing
	 */
	public static void createJson() throws IOException {
		JFileChooser jfc = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());
		jfc.setDialogTitle("Multiple file and directory selection:");
		jfc.setMultiSelectionEnabled(true);
		jfc.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
		int returnValue = jfc.showOpenDialog(null);
		String pdfName = "";
		String excelPath = "";
		String pdfPath = "";
		if (returnValue == JFileChooser.APPROVE_OPTION) {
			File[] files = jfc.getSelectedFiles();

			for (File file : files) {

				int index = file.getName().lastIndexOf('.');
				if (file.getName().substring(index + 1).equals("pdf")) {
					pdfName = file.getName();
					pdfPath = file.getAbsolutePath();
				}
				if (file.getName().substring(index + 1).equals("xlsx")) {
					excelPath = file.getAbsolutePath();
				}

			}
			if (files.length == 2) {
				create(pdfName, pdfPath, excelPath);
			} else {
				JOptionPane.showMessageDialog(null, "vous devriez saisir deux fichiers", "Erreur",
						JOptionPane.ERROR_MESSAGE);
			}

		}
	}

	public static void create(String pdfName, String pdfPath, String pathExcel) {
		InputStream excelFileToRead;
		XSSFWorkbook wb = null;
		try {
			excelFileToRead = new FileInputStream(pathExcel);
			try {
				wb = new XSSFWorkbook(excelFileToRead);
			} catch (IOException e) {
				e.printStackTrace();
			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		XSSFSheet sheet = wb.getSheetAt(1);
		Map<String, Object> jsonFile = new HashMap<>();
		List<Map<String, Object>> defaultList = new ArrayList<>();
		List<Map<String, Object>> defaultListOrder = new ArrayList<>();
		//transport container
		List<Map<String, Object>> transport = new ArrayList<>();
		//marchandise container
		List<Map<String, Object>> marchandise = new ArrayList<>();
		//Certification container
		List<Map<String, Object>> certification = new ArrayList<>();
		List<Map<String, Object>> tableFields = new ArrayList<>();

		// Create containers
		Map<String, Object> transportContainer = new HashMap<>();
		transportContainer.put("key", "Transport");
		transportContainer.put("component", "multiField");
		transportContainer.put("editable", true);
		transportContainer.put("index", "");
		transportContainer.put("label", "");
		transportContainer.put("description", "");
		transportContainer.put("placeholder", "");
		transportContainer.put("options", Arrays.asList());
		transportContainer.put("required", false);
		transportContainer.put("validation", "/.*/");
		transportContainer.put("id", "default-multiField-1");
		transportContainer.put("isContainer", true);
		transportContainer.put("expressionProperties", "");
		transportContainer.put("noFormControl", true);
		Map<String, Object> configTransport = new HashMap<>();
		configTransport.put("wrapper", "[panel]");
		configTransport.put("panelEntete", "TRANSPORT");
		Map<String, Object> templateOptionsTransport = new HashMap<>();
		templateOptionsTransport.put("config", configTransport);
		transportContainer.put("templateOptions", templateOptionsTransport);

		Map<String, Object> marchandiseContainer = new HashMap<>();
		marchandiseContainer.put("key", "Marchandises");
		marchandiseContainer.put("component", "multiField");
		marchandiseContainer.put("editable", true);
		marchandiseContainer.put("index", "");
		marchandiseContainer.put("label", "");
		marchandiseContainer.put("description", "");
		marchandiseContainer.put("placeholder", "");
		marchandiseContainer.put("options", Arrays.asList());
		marchandiseContainer.put("required", false);
		marchandiseContainer.put("validation", "/.*/");
		marchandiseContainer.put("id", "default-multiField-2");
		marchandiseContainer.put("isContainer", true);
		marchandiseContainer.put("expressionProperties", "");
		marchandiseContainer.put("noFormControl", true);
		Map<String, Object> configMarchandise = new HashMap<>();
		configMarchandise.put("wrapper", "[panel]");
		configMarchandise.put("panelEntete", "MARCHANDISE");
		Map<String, Object> templateOptionsMarchandise = new HashMap<>();
		templateOptionsMarchandise.put("config", configMarchandise);
		marchandiseContainer.put("templateOptions", templateOptionsMarchandise);

		Map<String, Object> certificationContainer = new HashMap<>();
		certificationContainer.put("key", "Certification");
		certificationContainer.put("component", "multiField");
		certificationContainer.put("editable", true);
		certificationContainer.put("index", "");
		certificationContainer.put("label", "");
		certificationContainer.put("description", "");
		certificationContainer.put("placeholder", "");
		certificationContainer.put("options", Arrays.asList());
		certificationContainer.put("required", false);
		certificationContainer.put("validation", "/.*/");
		certificationContainer.put("id", "default-multiField-3");
		certificationContainer.put("isContainer", true);
		certificationContainer.put("expressionProperties", "");
		certificationContainer.put("noFormControl", true);
		Map<String, Object> configCertification = new HashMap<>();
		configCertification.put("wrapper", "[panel]");
		configCertification.put("panelEntete", "CERTIFICATION");
		Map<String, Object> templateOptionsCertification = new HashMap<>();
		templateOptionsCertification.put("config", configCertification);
		certificationContainer.put("templateOptions", templateOptionsCertification);
        // Get data from dictionary
		String model = "";
		String certificatName = "";
		String fieldName = "";
		String label;
		String fieldType = "";
		String category = "";
		String format = "";
		int index = 0;
		int defaultIndex = 0;
		boolean isPdfFileExist = false;
		int idTable = 4;
		for (int rowNumber = sheet.getFirstRowNum() + 1; (rowNumber <= sheet.getLastRowNum()); rowNumber++) {

			model = "";
			certificatName = "";
			label = "";
			fieldType = "";
			category = "";
			format = "";
			fieldName = "";

			Row currentRow = sheet.getRow(rowNumber);

			if (currentRow != null && currentRow.getFirstCellNum() >= 0 && currentRow.getLastCellNum() >= 0) {
				for (int cellNumber = currentRow.getFirstCellNum(); cellNumber <= currentRow
						.getLastCellNum(); cellNumber++) {
					Cell currentCell = currentRow.getCell(cellNumber);

					if (currentCell != null && cellNumber >= 0) {
						if (cellNumber == 0) {
							model = currentCell.toString();
						}

						if (cellNumber == 1) {
							certificatName = currentCell.toString();
						}

						if (cellNumber == 2) {
							fieldName = currentCell.toString();

						}

						if (cellNumber == 3) {
							label = currentCell.toString();

						}
						if (cellNumber == 4) {
							fieldType = currentCell.toString();

						}
						if (cellNumber == 5) {
							format = currentCell.toString();

						}

						if (cellNumber == 10) {
							category = currentCell.toString();
						}
					}

				}
				if (certificatName.equals(pdfName)) {
					index++;
					isPdfFileExist = true;
					logger.info("Model : " + model + " Nom du champs :" + fieldName + " Type :" + fieldType
							+ " Categorie :" + category);

					String typeField = getFieldType(fieldType);

					if (fieldType.toUpperCase().equals(FieldType.DATATABLE)) {
						logger.info("-------------Nom du tableau trouvé est : ==> " + fieldName);
						idTable++;
						format = format.replaceAll("\\s", "");
						int repeatMin = 0;
						int repeatMax = 0;
						if (!format.isEmpty()) {
							char ch = format.charAt(0);
							repeatMin = Integer.parseInt(String.valueOf(ch));
							repeatMax = Integer.parseInt(format.substring(2));
						} else {
							logger.warn("Attention ! Le nombre de ligne du tableau n'est pas renseigné");
						}
						//Create table container
						Map<String, Object> tableContainer = new HashMap<>();
						tableContainer.put("key", fieldName);
						tableContainer.put("component", "multiField");
						tableContainer.put("editable", true);
						tableContainer.put("index", "");
						tableContainer.put("label", "");
						tableContainer.put("description", "");
						tableContainer.put("options", Arrays.asList());
						tableContainer.put("required", false);
						tableContainer.put("validation", "/.*/");
						tableContainer.put("id", "default-multiField-" + idTable);
						tableContainer.put("isContainer", true);
						tableContainer.put("expressionProperties", "");
						tableContainer.put("noFormControl", true);
						Map<String, Object> configTable = new HashMap<>();
						configTable.put("wrapper", "[panel]");
						configTable.put("panelEntete", fieldName);
						configTable.put("isRepeater", true);
						configTable.put("isTableau", true);
						configTable.put("repeatMin", repeatMin);
						configTable.put("repeatMax", repeatMax);
						Map<String, Object> templateOptionsTable = new HashMap<>();
						templateOptionsTable.put("config", configTable);
						tableContainer.put("templateOptions", templateOptionsTable);
                       
						//Get fields for table columns
						tableFields = getTableColumns(wb, model, fieldName);
						jsonFile.put("default-multiField-" + idTable, tableFields);
						if (category.toUpperCase().equals(Category.MARCHANDISES)) {
							marchandise.add(tableContainer);
						} else if (category.toUpperCase().equals(Category.TRANSPORT)) {
							transport.add(tableContainer);
						} else if (category.toUpperCase().equals(Category.CERTIFICATION)) {
							certification.add(tableContainer);
						} else {
							defaultList.add(tableContainer);
						}

					} else {

						Map<String, Object> field = new HashMap<>();
						Map<String, Object> chaine = new HashMap<>();
						Map<String, Object> style = new HashMap<>();
						Map<String, Object> enume = new HashMap<>();
						Map<String, Object> config = new HashMap<>();
						Map<String, Object> templateOptions = new HashMap<>();

						field.put("key", fieldName);
						field.put("component", "FamInput");
						field.put("editable", true);
						field.put("index", "");
						field.put("label", label);
						field.put("description", "");
						field.put("placeholder", "");
						field.put("options", Arrays.asList());
						field.put("required", false);
						field.put("validation", "/.*/");
						field.put("isContainer", false);
						field.put("expressionProperties", "");
						field.put("noFormControl", true);
						chaine.put("regExp", "none");
						style.put("bootstrapColXs", 6);
						style.put("bootstrapColSm", 6);
						style.put("bootstrapColMd", 6);
						style.put("bootstrapColLg", 6);
						enume.put("list",
								"[{\"label\":\"ORIGINAL\",\"value\":\"O\"} ,{\"label\":\"DUPLICATA\",\"value\":\"D\"}]");
						enume.put("type", "Liste");
						config.put("requiredValidationMessage", "Le champ est obligatoire !");
						config.put("messageFailTypeValidation", "La valeur du champ n\\\"est pas valide");

						config.put("inputType", typeField);
						config.put("chaine", chaine);
						config.put("style", style);
						if (fieldType.toUpperCase().equals(FieldType.RADIO)) {
							config.put("enum", enume);
						}
						templateOptions.put("config", config);
						field.put("templateOptions", templateOptions);

						if (category.toUpperCase().equals(Category.TRANSPORT)) {
							field.put("id", "default-multiField-1-FamInput-" + index);
							transport.add(field);
						} else if (category.toUpperCase().equals(Category.MARCHANDISES)) {
							field.put("id", "default-multiField-2-FamInput-" + index);
							marchandise.add(field);
						} else if (category.toUpperCase().equals(Category.CERTIFICATION)) {
							// certification.add(field);
							if (!fieldName.equals("CASE QUALITE_CERTIFICAT_OPT_COMP")
									&& !fieldName.equals("NOMBRE_TOTAL_DUPLICATA_CERTIFICAT_OBLIG_COMP")
									&& !fieldName.equals("NUMÉRO_CERTIFICAT_UNIQUE_OBLIG_COMP")) {
								field.put("id", "default-multiField-3-FamInput-" + index);
								certification.add(field);
							} else {
								defaultIndex++;
								field.put("id", "FamInput-" + defaultIndex);
								defaultListOrder.add(field);
							}

						} else {
							defaultListOrder.add(field);
						}

					}
				}

			}

		}

		if (!isPdfFileExist) {
			logger.error("Attention ! Le nom du pdf n'existe pas dans le dictionnaire,veuillez verifier les fichiers ");
		}

		if (!certification.isEmpty()) {
			defaultList.add(certificationContainer);
		}
		if (!marchandise.isEmpty()) {
			defaultList.add(marchandiseContainer);
		}
		if (!transport.isEmpty()) {
			defaultList.add(transportContainer);
		}

		for (Map<String, Object> mp : defaultListOrder) {
			defaultList.add(mp);
		}

		jsonFile.put("default", defaultList);

		jsonFile.put("default-multiField-1", transport);
		jsonFile.put("default-multiField-2", marchandise);
		jsonFile.put("default-multiField-3", certification);

		//Writing the json file

		String fileDestination = pdfPath.substring(0, pdfPath.length() - 4) + ".json";

		try {
			BufferedWriter writer = new BufferedWriter(new FileWriter(new File(fileDestination)));
			ObjectMapper mapper = new ObjectMapper();
			mapper.setSerializationInclusion(Include.NON_NULL);
			writer.write(mapper.writeValueAsString(jsonFile));
			writer.close();
			logger.info("Fin d'écriture du fichier json");
			JOptionPane.showMessageDialog(null, "Fin d'écriture du fichier json", "Information",
					JOptionPane.INFORMATION_MESSAGE);

		} catch (Exception ex) {
			ex.printStackTrace();
		}

	}

	/**
	 * Create fields for table columns
	 */
	public static List<Map<String, Object>> getTableColumns(XSSFWorkbook wb, String model, String nameTable) {

		XSSFSheet sheet2 = wb.getSheetAt(2);
		List<Map<String, Object>> columns = new ArrayList<Map<String, Object>>();
		String tableName = "";
		String fieldName = "";
		String label = "";
		String typeField = "";
		int index = 0;
		String codeModel = model + "." + nameTable;
		int columnsCount = getColumsCount(wb, model, nameTable);
		logger.info("Le nombre de colonnes du tableau est = " + columnsCount);
		for (int rowNumber = sheet2.getFirstRowNum() + 1; rowNumber <= sheet2.getLastRowNum(); rowNumber++) {
			tableName = "";
			fieldName = "";
			label = "";
			typeField = "";
			Row currentRow = sheet2.getRow(rowNumber);
			if (currentRow != null) {

				for (int cellNumber = currentRow.getFirstCellNum(); cellNumber <= currentRow
						.getLastCellNum(); cellNumber++) {

					Cell currentCell = currentRow.getCell(cellNumber);

					if (currentCell != null) {

						if (cellNumber == 0) {
							tableName = currentCell.toString();

						}

						if (cellNumber == 1) {
							fieldName = currentCell.toString();
						}

						if (cellNumber == 2) {
							label = currentCell.toString();

						}
						if (cellNumber == 3) {
							typeField = currentCell.toString();

						}
					}
				}
			}
			if (tableName.equals(codeModel)) {
				index++;
				logger.info("Nom de la Colonne " + index + "==>" + fieldName);
				typeField = getFieldType(typeField);
				Map<String, Object> field = new HashMap<String, Object>();
				Map<String, Object> chaine = new HashMap<String, Object>();
				Map<String, Object> style = new HashMap<String, Object>();
				Map<String, Object> styleLabel = new HashMap<String, Object>();
				Map<String, Object> config = new HashMap<String, Object>();
				Map<String, Object> templateOptions = new HashMap<String, Object>();

				field.put("key", fieldName);
				field.put("component", "FamInput");
				field.put("editable", true);
				field.put("index", index);
				field.put("label", label);
				field.put("description", "");
				field.put("placeholder", "");
				field.put("options", Arrays.asList());
				field.put("required", false);
				field.put("validation", "/.*/");
				field.put("id", "default-multiField-4-FamInput-" + index);
				field.put("isContainer", false);
				field.put("expressionProperties", "");
				field.put("noFormControl", true);
				chaine.put("regExp", "none");
                //Style according to the number of columns
				switch (columnsCount) {
				case 1:

					style.put("bootstrapColXs", 12);
					style.put("bootstrapColSm", 12);
					style.put("bootstrapColMd", 12);
					style.put("bootstrapColLg", 12);

					styleLabel.put("bootstrapColXs", 12);
					styleLabel.put("bootstrapColSm", 12);
					styleLabel.put("bootstrapColMd", 12);
					styleLabel.put("bootstrapColLg", 12);

					break;

				case 2:
					style.put("bootstrapColXs", 6);
					style.put("bootstrapColSm", 6);
					style.put("bootstrapColMd", 6);
					style.put("bootstrapColLg", 6);

					styleLabel.put("bootstrapColXs", 6);
					styleLabel.put("bootstrapColSm", 6);
					styleLabel.put("bootstrapColMd", 6);
					styleLabel.put("bootstrapColLg", 6);
					break;
				case 3:
					style.put("bootstrapColXs", 4);
					style.put("bootstrapColSm", 4);
					style.put("bootstrapColMd", 4);
					style.put("bootstrapColLg", 4);

					styleLabel.put("bootstrapColXs", 4);
					styleLabel.put("bootstrapColSm", 4);
					styleLabel.put("bootstrapColMd", 4);
					styleLabel.put("bootstrapColLg", 4);
					break;
				case 4:
					style.put("bootstrapColXs", 3);
					style.put("bootstrapColSm", 3);
					style.put("bootstrapColMd", 3);
					style.put("bootstrapColLg", 3);

					styleLabel.put("bootstrapColXs", 3);
					styleLabel.put("bootstrapColSm", 3);
					styleLabel.put("bootstrapColMd", 3);
					styleLabel.put("bootstrapColLg", 3);
					break;
				case 5:
					style.put("bootstrapColXs", 2);
					style.put("bootstrapColSm", 2);
					style.put("bootstrapColMd", 2);
					style.put("bootstrapColLg", 2);

					styleLabel.put("bootstrapColXs", 2);
					styleLabel.put("bootstrapColSm", 2);
					styleLabel.put("bootstrapColMd", 2);
					styleLabel.put("bootstrapColLg", 2);
					break;
				case 6:
					style.put("bootstrapColXs", 2);
					style.put("bootstrapColSm", 2);
					style.put("bootstrapColMd", 2);
					style.put("bootstrapColLg", 2);

					styleLabel.put("bootstrapColXs", 2);
					styleLabel.put("bootstrapColSm", 2);
					styleLabel.put("bootstrapColMd", 2);
					styleLabel.put("bootstrapColLg", 2);
					break;

				default:
					if (columnsCount == 9 || columnsCount == 7 || columnsCount == 10 || columnsCount == 11) {
						style.put("bootstrapColXs", 1);
						style.put("bootstrapColSm", 1);
						style.put("bootstrapColMd", 1);
						style.put("bootstrapColLg", 1);

						styleLabel.put("bootstrapColXs", 1);
						styleLabel.put("bootstrapColSm", 1);
						styleLabel.put("bootstrapColMd", 1);
						styleLabel.put("bootstrapColLg", 1);
					}
					break;

				}

				config.put("inputType", typeField);
				config.put("requiredValidationMessage", "Le champ est obligatoire !");
				config.put("messageFailTypeValidation", "La valeur du champ n\\\"est pas valide");
				config.put("chaine", chaine);
				config.put("style", style);
				config.put("alignRight", true);
				config.put("stylelabel", styleLabel);
				templateOptions.put("config", config);
				field.put("templateOptions", templateOptions);
				columns.add(field);
			}

		}

		return columns;

	}

	/**
	 * Compute the number of columns
	 * @param wb
	 * @param model
	 * @param nameTable
	 * @return
	 */
	
	public static int getColumsCount(XSSFWorkbook wb, String model, String nameTable) {
		XSSFSheet sheet2 = wb.getSheetAt(2);
		String tableName = "";
		int counter = 0;
		String codeModel = model + "." + nameTable;

		for (int rowNumber = sheet2.getFirstRowNum() + 1; rowNumber <= sheet2.getLastRowNum(); rowNumber++) {
			Row currentRow = sheet2.getRow(rowNumber);
			if (currentRow != null && currentRow.getFirstCellNum() >= 0 && currentRow.getLastCellNum() >= 0) {
				for (int cellNumber = currentRow.getFirstCellNum(); cellNumber <= currentRow
						.getLastCellNum(); cellNumber++) {
					Cell currentCell = currentRow.getCell(cellNumber);
					if (currentCell != null) {

						if (cellNumber == 0) {
							tableName = currentCell.toString();

						}

					}

				}
			}

			if (tableName.equals(codeModel)) {
				counter++;
			}

		}
		return counter;

	}
/**
 * Get type of the field 
 * @param type
 * @return
 */
	public static String getFieldType(String type) {
		String typeField = "";
		switch (type.toUpperCase()) {
		case FieldType.TEXTFIELD:
			typeField = "[chaine]";
			break;
		case FieldType.TEXTAREA:
			typeField = "[text]";
			break;
		case FieldType.RADIO:
			typeField = "[enum]";
			break;
		case FieldType.DATE:
			typeField = "[date]";
			break;
		case FieldType.BOOLEEN:
			typeField = "[bool]";
			break;
		}

		return typeField;

	}

}
