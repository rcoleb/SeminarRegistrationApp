package com.fhs;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.Properties;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.itextpdf.forms.PdfAcroForm;
import com.itextpdf.forms.fields.PdfFormField;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfReader;
import com.itextpdf.kernel.pdf.PdfWriter;

public class SeminarRegistrationApp {
  static Properties config = loadProperties();

  public static void main(String[] args) {
    RegistrationProcessor processor = new RegistrationProcessor(config.getProperty("excelFilePath"), config.getProperty("pdfTemplatePath"), config.getProperty("outputFolderPath"));
    processor.processRegistrations();
  }

  private static Properties loadProperties() {
    Properties props = new Properties();
    try (InputStream input = new FileInputStream("config.properties")) {
      props.load(input);
    } catch (IOException ex) {
      ex.printStackTrace();
    }
    return props;
  }
}

class RegistrationProcessor {
  private String excelFilePath;
  private String pdfTemplatePath;
  private String outputFolderPath;

  public RegistrationProcessor(String excelFilePath, String pdfTemplatePath, String outputFolderPath) {
    this.excelFilePath = excelFilePath;
    this.pdfTemplatePath = pdfTemplatePath;
    this.outputFolderPath = outputFolderPath;
  }

  public void processRegistrations() {
    ExcelReader excelReader = new ExcelReader(excelFilePath);
    PDFWriter pdfWriter = new PDFWriter(pdfTemplatePath, outputFolderPath);

    try {
      List<RegistrationData> registrations = excelReader.readRegistrations();
      for (RegistrationData registration : registrations) {
        String outputFileName = registration.getSafeFileName();
        pdfWriter.fillForm(registration, outputFileName);
      }
    } catch (IOException e) {
      System.err.println("Failed to process registrations: " + e.getMessage());
      e.printStackTrace();
    }
  }
}

class ExcelReader {
  private String filePath;

  public ExcelReader(String filePath) {
    this.filePath = filePath;
  }

  public List<RegistrationData> readRegistrations() throws IOException {
    List<RegistrationData> registrations = new ArrayList<>();
    try (FileInputStream excelFile = new FileInputStream(new File(filePath)); Workbook workbook = new XSSFWorkbook(excelFile)) {
      Sheet sheet = workbook.getSheetAt(0);
      Map<String, Integer> columnMap = getColumnMap(sheet.getRow(0));

      for (int i = 1; i <= sheet.getLastRowNum(); i++) {
        Row row = sheet.getRow(i);
        if (row == null) continue;

        RegistrationData registrationData = new RegistrationData(row, columnMap);
        registrations.add(registrationData);
      }
    }
    return registrations;
  }

  private Map<String, Integer> getColumnMap(Row headerRow) {
    Map<String, Integer> map = new HashMap<>();
    for (Cell cell : headerRow) {
      if (cell.getCellType() == CellType.STRING) {
        map.put(cell.getStringCellValue(), cell.getColumnIndex());
      }
    }
    return map;
  }
}

enum FieldConstants {
  FULL_NAME("Full Name", "Name", new DefaultFieldHandler()),
  DOJO("Dojo (if applicable)", "Dojo", new DefaultFieldHandler()),
  RANK("Rank (if applicable)", "Rank", new DefaultFieldHandler()),
  EMAIL("E-mail", "Email", new DefaultFieldHandler()),
  PHONE_NUMBER("Phone Number", "Phone", new DefaultFieldHandler()),
  NUM_PEOPLE("Number of people attending:", "NumPpl", new NumberFieldHandler()),
  DAYS_ATTENDING("Day(s) attending:", "Day", new DayFieldHandler());

  final String excelHeader;
  final String fieldName;
  final FieldHandler handler;

  FieldConstants(String excelHeader, String fieldName, FieldHandler handler) {
    this.excelHeader = excelHeader;
    this.fieldName = fieldName;
    this.handler = handler;
  }

  public static Optional<FieldConstants> getHandlerForExcelField(String excelHeader) {
    return Arrays.stream(FieldConstants.values()).filter(fc -> fc.excelHeader.equalsIgnoreCase(excelHeader)).findFirst();
  }

}

class RegistrationData {
  Map<String, String> fields = new HashMap<>();

  public RegistrationData(Row row, Map<String, Integer> columnMap) {
    columnMap.forEach((fieldName, index) -> {
      Cell cell = row.getCell(index, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
      String value = (cell != null && cell.getCellType() == CellType.STRING) ? cell.getStringCellValue() : "";
      Optional<FieldConstants> fieldConstant = FieldConstants.getHandlerForExcelField(fieldName);
      if (fieldConstant.isPresent()) {
        FieldHandler handler = fieldConstant.get().handler;
        handler.processField(fields, fieldConstant.get().fieldName, value);
      }
    });
  }

  public String getSafeFileName() {
    String name = getField(FieldConstants.FULL_NAME.fieldName).replaceAll("[^a-zA-Z0-9_]", "_");
    if (name.isEmpty()) {
      return "default_name.pdf";
    }
    return "filled_form_" + name + ".pdf";
  }

  public String getField(String key) {
    return fields.getOrDefault(key, "");
  }
}

interface FieldHandler {
  void processField(Map<String, String> fields, String fieldName, String value);
}

class DefaultFieldHandler implements FieldHandler {
  @Override
  public void processField(Map<String, String> fields, String fieldName, String value) {
    fields.put(fieldName, value);
  }
}

class DayFieldHandler implements FieldHandler {
  @Override
  public void processField(Map<String, String> fields, String fieldName, String value) {
    String key = switch (value) {
      case "Saturday and Sunday", "Saturday and Sunday (w/o Banquet)" -> "Sat/Sun People";
      case "Saturday Only", "Saturday Only (w/o Banquet)" -> "Sat People";
      case "Sunday Only" -> "Sun People";
      default -> "";
    };
    if (!key.isEmpty()) {
      fields.put(key, fields.getOrDefault(FieldConstants.NUM_PEOPLE.fieldName, "1")); // Default to
                                                                                      // 1 if no
                                                                                      // value is
                                                                                      // set
    }
  }
}

class NumberFieldHandler implements FieldHandler {
  private static final String NUM_PPL = "NumPpl";

  @Override
  public void processField(Map<String, String> fields, String fieldName, String value) {
    try {
      int num = Integer.parseInt(value);
      fields.put(NUM_PPL, Integer.toString(num));
    } catch (NumberFormatException e) {
      fields.put(NUM_PPL, "1"); // Default to 1 if parsing fails
    }
  }
}

class PDFWriter {
  private String templatePath;
  private String outputFolderPath;

  public PDFWriter(String templatePath, String outputFolderPath) {
    this.templatePath = templatePath;
    this.outputFolderPath = outputFolderPath;
  }

  public void fillForm(RegistrationData registrationData, String outputFileName) throws IOException {
    try (PdfDocument pdfDoc = new PdfDocument(new PdfReader(templatePath), new PdfWriter(outputFolderPath + File.separator + outputFileName))) {
      PdfAcroForm form = PdfAcroForm.getAcroForm(pdfDoc, true);
      registrationData.fields.forEach((key, value) -> {
        PdfFormField field = form.getField(key);
        if (field != null) {
          field.setValue(value);
          field.setFontSize(Math.max(field.getFontSize() - 2, 10));
        }
      });
    }
  }
}