/*
 * Copyright (C) 2022 jtalbut
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <http://www.gnu.org/licenses/>.
 */
package uk.co.spudsoft.xlsx.sandbox;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.util.Arrays;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;
import org.junit.jupiter.api.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 *
 * @author jtalbut
 */
public class SynchronousTest {
  
  @SuppressWarnings("constantname")
  private static final Logger logger = LoggerFactory.getLogger(SynchronousTest.class);

  //some static parts of the XLSX file:
  private static final String content_types_xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"no\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default ContentType=\"application/vnd.openxmlformats-package.relationships+xml\" Extension=\"rels\"/><Default ContentType=\"application/xml\" Extension=\"xml\"/><Override ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\" PartName=\"/docProps/app.xml\"/><Override ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\" PartName=\"/docProps/core.xml\"/><Override ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\" PartName=\"/xl/sharedStrings.xml\"/><Override ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\" PartName=\"/xl/styles.xml\"/><Override ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\" PartName=\"/xl/workbook.xml\"/><Override ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\" PartName=\"/xl/worksheets/sheet1.xml\"/></Types>";

  private static final String docProps_app_xml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"><Application>" + "Created Low level From Scratch" + "</Application></Properties>";

  private static final String docProps_core_xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"no\"?><cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"><dcterms:created xsi:type=\"dcterms:W3CDTF\">" + java.time.Instant.now().truncatedTo(java.time.temporal.ChronoUnit.SECONDS).toString() + "</dcterms:created><dc:creator>" + "Axel Richter from scratch" + "</dc:creator></cp:coreProperties>";

  private static final String _rels_rels_xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"no\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Target=\"xl/workbook.xml\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\"/><Relationship Id=\"rId2\" Target=\"docProps/app.xml\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\"/><Relationship Id=\"rId3\" Target=\"docProps/core.xml\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\"/></Relationships>";

  private static final String xl_rels_workbook_xml_rels_xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"no\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Target=\"sharedStrings.xml\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\"/><Relationship Id=\"rId2\" Target=\"styles.xml\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\"/><Relationship Id=\"rId3\" Target=\"worksheets/sheet1.xml\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\"/></Relationships>";

  private static final String xl_sharedstrings_xml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><sst count=\"0\" uniqueCount=\"0\" xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"/>";

  private static final String xl_styles_xml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><numFmts count=\"0\"/><fonts count=\"1\"><font><sz val=\"11.0\"/><color indexed=\"8\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font></fonts><fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"darkGray\"/></fill></fills><borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs><cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs></styleSheet>";

  private static final String xl_workbook_xml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><workbookPr date1904=\"false\"/><bookViews><workbookView activeTab=\"0\"/></bookViews><sheets><sheet name=\"" + "Sheet1" + "\" r:id=\"rId3\" sheetId=\"1\"/></sheets></workbook>";

  // private static final String xl_worksheets_sheet1_xml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><dimension ref=\"A1\"/><sheetViews><sheetView workbookViewId=\"0\" tabSelected=\"true\"/></sheetViews><sheetFormatPr defaultRowHeight=\"15.0\"/><sheetData/><pageMargins bottom=\"0.75\" footer=\"0.3\" header=\"0.3\" left=\"0.7\" right=\"0.7\" top=\"0.75\"/></worksheet>";
  private static final String xl_worksheets_sheet1_xml_start = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetViews><sheetView workbookViewId=\"0\" tabSelected=\"true\"/></sheetViews><sheetFormatPr defaultRowHeight=\"15.0\"/><sheetData>";
  private static final String xl_worksheets_sheet1_xml_end = "</sheetData><pageMargins bottom=\"0.75\" footer=\"0.3\" header=\"0.3\" left=\"0.7\" right=\"0.7\" top=\"0.75\"/></worksheet>";

  @Test
  public void outputWorkbook() throws Exception {
    File file = new File("target/temp/SynchronousTest.xlsx");
    file.getParentFile().mkdirs();

    ZipOutputStream zipout = createFile(file);
    for (int i = 1; i < 100; ++i) {
      outputRow(zipout, Arrays.asList(i, "Text", 100.0 / i));
    }
    closeFile(zipout);

  }

  private int r = 0;
  
  ZipOutputStream createFile(File outputFile) throws Exception {

    // result goes into a ByteArrayOutputStream
    OutputStream resultout = new FileOutputStream(outputFile);

    // needed objects
    ZipEntry zipentry;
    byte[] data;

    // create ZipOutputStream
    ZipOutputStream zipout = new ZipOutputStream(resultout);

    // create the static parts of the XLSX ZIP file:
    zipentry = new ZipEntry("[Content_Types].xml");
    zipout.putNextEntry(zipentry);
    data = content_types_xml.getBytes();
    zipout.write(data, 0, data.length);
    zipout.closeEntry();

    zipentry = new ZipEntry("docProps/app.xml");
    zipout.putNextEntry(zipentry);
    data = docProps_app_xml.getBytes();
    zipout.write(data, 0, data.length);
    zipout.closeEntry();

    zipentry = new ZipEntry("docProps/core.xml");
    zipout.putNextEntry(zipentry);
    data = docProps_core_xml.getBytes();
    zipout.write(data, 0, data.length);
    zipout.closeEntry();

    zipentry = new ZipEntry("_rels/.rels");
    zipout.putNextEntry(zipentry);
    data = _rels_rels_xml.getBytes();
    zipout.write(data, 0, data.length);
    zipout.closeEntry();

    zipentry = new ZipEntry("xl/_rels/workbook.xml.rels");
    zipout.putNextEntry(zipentry);
    data = xl_rels_workbook_xml_rels_xml.getBytes();
    zipout.write(data, 0, data.length);
    zipout.closeEntry();

    zipentry = new ZipEntry("xl/sharedStrings.xml");
    zipout.putNextEntry(zipentry);
    data = xl_sharedstrings_xml.getBytes();
    zipout.write(data, 0, data.length);
    zipout.closeEntry();

    zipentry = new ZipEntry("xl/styles.xml");
    zipout.putNextEntry(zipentry);
    data = xl_styles_xml.getBytes();
    zipout.write(data, 0, data.length);
    zipout.closeEntry();

    zipentry = new ZipEntry("xl/workbook.xml");
    zipout.putNextEntry(zipentry);
    data = xl_workbook_xml.getBytes();
    zipout.write(data, 0, data.length);
    zipout.closeEntry();
    
    // create the xl/worksheets/sheet1.xml
    zipentry = new ZipEntry("xl/worksheets/sheet1.xml");
    zipout.putNextEntry(zipentry);
    data = xl_worksheets_sheet1_xml_start.getBytes(StandardCharsets.UTF_8);
    zipout.write(data, 0, data.length);
    
    return zipout;
  }
  
  private static int toNumber(String name) {
    int number = 0;
    for (int i = 0; i < name.length(); i++) {
      number = number * 26 + (name.charAt(i) - ('A' - 1));
    }
    return number;
  }

  private static String toName(int number) {
    StringBuilder sb = new StringBuilder();
    while (number-- > 0) {
      sb.append((char) ('A' + (number % 26)));
      number /= 26;
    }
    return sb.reverse().toString();
  }

  private void outputRow(ZipOutputStream zipout, List<Object> values) throws Exception {
    StringBuilder rowString = new StringBuilder();
    rowString.append("<row r=\"").append(++r).append("\">");
    
    int colNum = 0;
    for (Object cellData : values) {
      rowString.append("<c r=\"").append(toName(++colNum)).append(r).append('"');

      if (cellData instanceof String && ((String) cellData).startsWith("=")) {
        rowString.append("><f>" + ((String) cellData).replace("=", "") + "</f></c>");
      } else if (cellData instanceof String) {
        rowString.append(" t=\"inlineStr\"><is><t>" + cellData + "</t></is></c>");
      } else if (cellData instanceof Double) {
        rowString.append("><v>" + cellData + "</v></c>");
      } else if (cellData instanceof Integer) {
        rowString.append("><v>" + cellData + "</v></c>");
      }
    }
    rowString.append("</row>");
    logger.debug("Row: {}", rowString);

    byte[] data = rowString.toString().getBytes(StandardCharsets.UTF_8);
    zipout.write(data, 0, data.length);
  }
  
  private void closeFile(ZipOutputStream zipout) throws Exception {
    byte[] data = xl_worksheets_sheet1_xml_end.getBytes(StandardCharsets.UTF_8);
    zipout.write(data, 0, data.length);
    zipout.closeEntry();

    zipout.finish();
  }
}
