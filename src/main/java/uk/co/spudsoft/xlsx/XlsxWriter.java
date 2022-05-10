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
package uk.co.spudsoft.xlsx;

import java.io.Closeable;
import java.io.IOException;
import java.io.OutputStream;
import java.math.RoundingMode;
import java.nio.charset.StandardCharsets;
import java.text.DecimalFormat;
import java.time.LocalDate;
import java.time.LocalTime;
import java.time.temporal.ChronoField;
import java.time.temporal.ChronoUnit;
import java.time.temporal.Temporal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/**
 * Output an XLSX file one row at a time, streaming the output so that it is written as the rows come in.
 * 
 * The Excel file is intended to contain large feeds of data (hence streaming is important) whilst enabling sufficient formatting to look nice.
 * 
 * It is important to note that although the OutputStream class specifies blocking operations, nothing in this class (nor in ZipOutputStream) actually blocks.
 * If the OutputStream passed in to the constructor can be guaranteed to not block then so can this class.
 * 
 * @author jtalbut
 */
@SuppressWarnings("checkstyle:membername")
public class XlsxWriter implements Closeable {
  
  /**
   * The default application name that will be reported in the document properties.
   */
  public static final String DEFAULT_APP_NAME = "Streaming XLSX Writer";
  
  /**
   * The date of the epoch for Excel date values (1 January 1970).
   */
  public static final LocalDate EPOCH_DATE = LocalDate.of(1900, 1, 1);
  
  /**
   * The data of the epoch for Excel time values (midnight).
   */
  public static final LocalTime EPOCH_TIME = LocalTime.of(0, 0);

  /**
   * The font to use if no alternative is specified.
   */
  public static final String DEFAULT_FONT_NAME = "Calibri";
  
  /**
   * The size of the font to use is not alternative is specified.
   */
  public static final int DEFAULT_FONT_SIZE = 11;
  
  private static final DecimalFormat DATE_FORMAT = prepareDateFormat();
  
  private final TableDefinition defn;
  private final int colCount;
  
  private final String contentTypes = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"bin\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings\"/><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/><Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/><Override PartName=\"/xl/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/><Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/><Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/><Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/><Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/></Types>";
  private final String rels_rels = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>";
  private final String docProps_app;
  private final String docProps_core;
  private final String xl_rels_workbook = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/><Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/></Relationships>";
  private final String xl_theme_theme1;  
  private final String xl_sharedstrings = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><sst count=\"0\" uniqueCount=\"0\" xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"/>";
  private final String xl_styles;
  private final String xl_workbook;
  private final String xl_worksheets_sheet1_start = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetViews><sheetView workbookViewId=\"0\" tabSelected=\"true\"/></sheetViews><sheetFormatPr defaultRowHeight=\"15.0\"/>";
  private final String xl_worksheets_sheet1_end = "</sheetData><pageMargins bottom=\"0.75\" footer=\"0.3\" header=\"0.3\" left=\"0.7\" right=\"0.7\" top=\"0.75\"/></worksheet>";
  
  private Map<String, Integer> numFmtIdMap = new HashMap<>();
  private ZipOutputStream zipout;
  private int r = 0;

  private String coalesce(String value1, String value2) {
    if (value1 == null || value1.isEmpty()) {
      return value2;
    }
    return value1;
  }
  
  /**
   * Constructor.
   * 
   * @param defn The definition of the formatting required in the workbook.
   */
  public XlsxWriter(TableDefinition defn) {
    this.defn = defn;
    this.colCount = defn.columns.size();

    this.docProps_app = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"><Application>" + coalesce(defn.application, DEFAULT_APP_NAME) + "</Application></Properties>";
    this.docProps_core = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"no\"?><cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"><dcterms:created xsi:type=\"dcterms:W3CDTF\">" + java.time.Instant.now().truncatedTo(java.time.temporal.ChronoUnit.SECONDS).toString() + "</dcterms:created><dc:creator>" + coalesce(defn.creator, DEFAULT_APP_NAME) + "</dc:creator></cp:coreProperties>";
    this.xl_workbook = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><workbookPr date1904=\"false\"/><bookViews><workbookView activeTab=\"0\"/></bookViews><sheets><sheet name=\"" + coalesce(defn.name, "Sheet1") + "\" r:id=\"rId1\" sheetId=\"1\"/></sheets></workbook>";
    this.xl_theme_theme1 = buildTheme(defn);
    this.xl_styles = buildStyles(defn);
  }    
  
  /**
   * Start outputting the metadata to the OutputStream.
   * @param stream The output stream that will be written to.
   * @throws IOException if something goes wrong - this should only happen if "stream" throws an exception.
   */
  public void startFile(OutputStream stream) throws IOException {
    // needed objects
    ZipEntry zipentry;
    byte[] data;

    // create ZipOutputStream
    zipout = new ZipOutputStream(stream);

    // create the static parts of the XLSX ZIP file:
    zipentry = new ZipEntry("[Content_Types].xml");
    zipout.putNextEntry(zipentry);
    data = contentTypes.getBytes();
    zipout.write(data, 0, data.length);
    zipout.closeEntry();

    zipentry = new ZipEntry("docProps/app.xml");
    zipout.putNextEntry(zipentry);
    data = docProps_app.getBytes();
    zipout.write(data, 0, data.length);
    zipout.closeEntry();

    zipentry = new ZipEntry("docProps/core.xml");
    zipout.putNextEntry(zipentry);
    data = docProps_core.getBytes();
    zipout.write(data, 0, data.length);
    zipout.closeEntry();

    zipentry = new ZipEntry("_rels/.rels");
    zipout.putNextEntry(zipentry);
    data = rels_rels.getBytes();
    zipout.write(data, 0, data.length);
    zipout.closeEntry();

    zipentry = new ZipEntry("xl/theme/theme1.xml");
    zipout.putNextEntry(zipentry);
    data = xl_theme_theme1.getBytes();
    zipout.write(data, 0, data.length);
    zipout.closeEntry();

    zipentry = new ZipEntry("xl/_rels/workbook.xml.rels");
    zipout.putNextEntry(zipentry);
    data = xl_rels_workbook.getBytes();
    zipout.write(data, 0, data.length);
    zipout.closeEntry();

    zipentry = new ZipEntry("xl/sharedStrings.xml");
    zipout.putNextEntry(zipentry);
    data = xl_sharedstrings.getBytes();
    zipout.write(data, 0, data.length);
    zipout.closeEntry();

    zipentry = new ZipEntry("xl/styles.xml");
    zipout.putNextEntry(zipentry);
    data = xl_styles.getBytes();
    zipout.write(data, 0, data.length);
    zipout.closeEntry();

    zipentry = new ZipEntry("xl/workbook.xml");
    zipout.putNextEntry(zipentry);
    data = xl_workbook.getBytes();
    zipout.write(data, 0, data.length);
    zipout.closeEntry();
    
    // create the xl/worksheets/sheet1.xml
    zipentry = new ZipEntry("xl/worksheets/sheet1.xml");
    zipout.putNextEntry(zipentry);
    data = xl_worksheets_sheet1_start.getBytes(StandardCharsets.UTF_8);
    zipout.write(data, 0, data.length);    
    
    if (anyColumnSpecifiesWidth()) {
      outputColumns();
    }
    
    data = "<sheetData>".getBytes(StandardCharsets.UTF_8);
    zipout.write(data, 0, data.length);    
    
    if (defn.headers) {
      outputHeaders();
    }
  }
  
  boolean anyColumnSpecifiesWidth() {
    for (ColumnDefinition col : defn.columns) {
      if (col.width != null) {
        return true;
      }
    }
    return false;
  }
  
  void outputColumns() throws IOException {
    StringBuilder bldr = new StringBuilder();
    bldr.append("<cols>");
    int colNum = 0;
    for (ColumnDefinition col : defn.columns) {
      ++colNum;
      Double width = col.width;
      if (width == null) {
        width = 11.0;
      }
      int s = colNum;
      bldr.append("<col min=\"").append(colNum).append("\" max=\"").append(colNum).append("\" width=\"").append(width).append("\" style=\"").append(s).append("\" customWidth=\"1\" />");
    }
    bldr.append("</cols>");
    
    byte[] data = bldr.toString().getBytes(StandardCharsets.UTF_8);
    zipout.write(data, 0, data.length);
  }
  
  void outputHeaders() throws IOException {
    StringBuilder rowString = new StringBuilder();
    rowString.append("<row r=\"").append(++r).append("\">");
    
    int colNum = 0;
    for (ColumnDefinition col : defn.columns) {
      ++colNum;
      int s = 1 + colCount + colNum;
      rowString.append("<c r=\"").append(toName(colNum)).append(r).append('"').append(" s=\"").append(s).append('"');
      rowString.append(" t=\"inlineStr\"><is><t>" + encodeSpecialCharacters(coalesce(col.name, "")) + "</t></is></c>");
    }
    rowString.append("</row>");

    byte[] data = rowString.toString().getBytes(StandardCharsets.UTF_8);
    zipout.write(data, 0, data.length);
  }
  
  static String toName(int number) {
    StringBuilder sb = new StringBuilder();
    while (number-- > 0) {
      sb.append((char) ('A' + (number % 26)));
      number /= 26;
    }
    return sb.reverse().toString();
  }
  
  private static DecimalFormat prepareDateFormat() {
    DecimalFormat df = new DecimalFormat("#.##################");
    df.setRoundingMode(RoundingMode.HALF_UP);    
    return df;
  }    

  static String temporalToExcelValue(Temporal ip) {
    boolean hasValue = false;
    double value = 0.0;
    if (ip.isSupported(ChronoField.DAY_OF_YEAR) && ip.isSupported(ChronoField.YEAR)) {
      // Add because Excel date is inclusive at both ends and one based
      value += 2 + ChronoUnit.DAYS.between(EPOCH_DATE, ip);
      hasValue = true;
    }
    if (ip.isSupported(ChronoField.MICRO_OF_DAY)) {
      value += ChronoUnit.MILLIS.between(EPOCH_TIME, ip) / (24.0 * 60 * 60 * 1000);
      hasValue = true;
    }
    if (hasValue) {
      return DATE_FORMAT.format(value);
    } else {
      return ip.toString();
    }
  }
  
  /**
   * Output a row of data to the output stream.
   * 
   * Note that, because of the buffering inherent in ZipOutputStream, this method may not result in a call to OutputStream.write.
   * 
   * The values are handling according to the following rules:
   * <ul>
   * <li>Null values are output as empty cells.
   * <li>String values starting with '=' are output as formulae.
   * <li>Other strings values are output as inline strings.
   * <li>Temporal values are output as numeric values complying with Excel data/time formatting.
   * <li>Number values are output as numeric values.
   * <li>Anything else is output as an inline string after calling toString() on it.
   * </ul>
   * 
   * Note that the handling of Temporal values should work for any jsr310 classes (ignoring time zones) but will not work for Date, or SQL Timestamp values.
   * 
   * @param values The values to add to the output, one column at a time.
   * @throws IOException if something goes wrong, this should only happen if the OutputStream throws.
   */
  public void outputRow(List<Object> values) throws IOException {
    StringBuilder rowString = new StringBuilder();
    rowString.append("<row r=\"").append(++r).append("\">");
    
    int colNum = 0;
    int colCount = defn.columns.size();
    for (Object cellData : values) {
      ++colNum;
      int s = (2 + r % 2) * (colCount + 1) + (colNum > colCount ? 0 : colNum);
      rowString.append("<c r=\"").append(toName(colNum)).append(r).append('"').append(" s=\"").append(s).append('"');

      if (cellData == null) {
        rowString.append("></c>");
      } else if (cellData instanceof String) {
        String cellString = (String) cellData;
        if (cellString.startsWith("=")) {
          rowString.append("><f>" + encodeSpecialCharacters(cellString.substring(1)) + "</f></c>");
        } else {
          rowString.append(" t=\"inlineStr\"><is><t>" + encodeSpecialCharacters(cellString) + "</t></is></c>");
        }
      } else if (cellData instanceof Temporal) {
        rowString.append("><v>").append(temporalToExcelValue((Temporal) cellData)).append("</v></c>");
      } else if (cellData instanceof Number) {
        rowString.append("><v>").append(cellData.toString()).append("</v></c>");
      } else {
        rowString.append(" t=\"inlineStr\"><is><t>" + encodeSpecialCharacters(cellData.toString()) + "</t></is></c>");
      }
    }
    rowString.append("</row>");

    byte[] data = rowString.toString().getBytes(StandardCharsets.UTF_8);
    zipout.write(data, 0, data.length);
  }

  @Override
  public void close() throws IOException {
    byte[] data = xl_worksheets_sheet1_end.getBytes(StandardCharsets.UTF_8);
    zipout.write(data, 0, data.length);
    zipout.closeEntry();

    zipout.finish();
  }
  
  String buildTheme(TableDefinition defn) {
    StringBuilder bldr = new StringBuilder();
    bldr.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>")
            .append("<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office Theme\">")

            .append("<a:themeElements>")
            
            .append("<a:clrScheme name=\"Office\">")
            .append("<a:dk1><a:sysClr val=\"windowText\" lastClr=\"000000\"/></a:dk1>")
            .append("<a:lt1><a:sysClr val=\"window\" lastClr=\"FFFFFF\"/></a:lt1>")            
            .append("<a:dk2><a:srgbClr val=\"44546A\"/></a:dk2>")            
            .append("<a:lt2><a:srgbClr val=\"E7E6E6\"/></a:lt2>")           
            .append("<a:accent1><a:srgbClr val=\"4472C4\"/></a:accent1>")            
            .append("<a:accent2><a:srgbClr val=\"ED7D31\"/></a:accent2>")
            .append("<a:accent3><a:srgbClr val=\"A5A5A5\"/></a:accent3>")
            .append("<a:accent4><a:srgbClr val=\"FFC000\"/></a:accent4>")
            .append("<a:accent5><a:srgbClr val=\"5B9BD5\"/></a:accent5>")
            .append("<a:accent6><a:srgbClr val=\"70AD47\"/></a:accent6>")
            .append("<a:hlink><a:srgbClr val=\"0563C1\"/></a:hlink>")
            .append("<a:folHlink><a:srgbClr val=\"954F72\"/></a:folHlink>")
            .append("</a:clrScheme>")
            
            .append("<a:fontScheme name=\"Office\">")
            .append("<a:majorFont>")
            .append("<a:latin typeface=\"Calibri Light\" panose=\"020F0302020204030204\"/>")
            .append("<a:ea typeface=\"\"/><a:cs typeface=\"\"/>")
            .append("<a:font script=\"Jpan\" typeface=\"游ゴシック Light\"/>")
            .append("<a:font script=\"Hang\" typeface=\"맑은 고딕\"/>")
            .append("<a:font script=\"Hans\" typeface=\"等线 Light\"/>")
            .append("<a:font script=\"Hant\" typeface=\"新細明體\"/>")
            .append("<a:font script=\"Arab\" typeface=\"Times New Roman\"/>")
            .append("<a:font script=\"Hebr\" typeface=\"Times New Roman\"/>")
            .append("<a:font script=\"Thai\" typeface=\"Tahoma\"/>")
            .append("<a:font script=\"Ethi\" typeface=\"Nyala\"/>")
            .append("<a:font script=\"Beng\" typeface=\"Vrinda\"/>")
            .append("<a:font script=\"Gujr\" typeface=\"Shruti\"/>")
            .append("<a:font script=\"Khmr\" typeface=\"MoolBoran\"/>")
            .append("<a:font script=\"Knda\" typeface=\"Tunga\"/>")
            .append("<a:font script=\"Guru\" typeface=\"Raavi\"/>")
            .append("<a:font script=\"Cans\" typeface=\"Euphemia\"/>")
            .append("<a:font script=\"Cher\" typeface=\"Plantagenet Cherokee\"/>")
            .append("<a:font script=\"Yiii\" typeface=\"Microsoft Yi Baiti\"/>")
            .append("<a:font script=\"Tibt\" typeface=\"Microsoft Himalaya\"/>")
            .append("<a:font script=\"Thaa\" typeface=\"MV Boli\"/>")
            .append("<a:font script=\"Deva\" typeface=\"Mangal\"/>")
            .append("<a:font script=\"Telu\" typeface=\"Gautami\"/>")
            .append("<a:font script=\"Taml\" typeface=\"Latha\"/>")
            .append("<a:font script=\"Syrc\" typeface=\"Estrangelo Edessa\"/>")
            .append("<a:font script=\"Orya\" typeface=\"Kalinga\"/>")
            .append("<a:font script=\"Mlym\" typeface=\"Kartika\"/>")
            .append("<a:font script=\"Laoo\" typeface=\"DokChampa\"/>")
            .append("<a:font script=\"Sinh\" typeface=\"Iskoola Pota\"/>")
            .append("<a:font script=\"Mong\" typeface=\"Mongolian Baiti\"/>")
            .append("<a:font script=\"Viet\" typeface=\"Times New Roman\"/>")
            .append("<a:font script=\"Uigh\" typeface=\"Microsoft Uighur\"/>")
            .append("<a:font script=\"Geor\" typeface=\"Sylfaen\"/>")
            .append("<a:font script=\"Armn\" typeface=\"Arial\"/>")
            .append("<a:font script=\"Bugi\" typeface=\"Leelawadee UI\"/>")
            .append("<a:font script=\"Bopo\" typeface=\"Microsoft JhengHei\"/>")
            .append("<a:font script=\"Java\" typeface=\"Javanese Text\"/>")
            .append("<a:font script=\"Lisu\" typeface=\"Segoe UI\"/>")
            .append("<a:font script=\"Mymr\" typeface=\"Myanmar Text\"/>")
            .append("<a:font script=\"Nkoo\" typeface=\"Ebrima\"/>")
            .append("<a:font script=\"Olck\" typeface=\"Nirmala UI\"/>")
            .append("<a:font script=\"Osma\" typeface=\"Ebrima\"/>")
            .append("<a:font script=\"Phag\" typeface=\"Phagspa\"/>")
            .append("<a:font script=\"Syrn\" typeface=\"Estrangelo Edessa\"/>")
            .append("<a:font script=\"Syrj\" typeface=\"Estrangelo Edessa\"/>")
            .append("<a:font script=\"Syre\" typeface=\"Estrangelo Edessa\"/>")
            .append("<a:font script=\"Sora\" typeface=\"Nirmala UI\"/>")
            .append("<a:font script=\"Tale\" typeface=\"Microsoft Tai Le\"/>")
            .append("<a:font script=\"Talu\" typeface=\"Microsoft New Tai Lue\"/>")
            .append("<a:font script=\"Tfng\" typeface=\"Ebrima\"/>")
            .append("</a:majorFont>")
            .append("<a:minorFont>")
            .append("<a:latin typeface=\"Calibri\" panose=\"020F0502020204030204\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/><a:font script=\"Jpan\" typeface=\"游ゴシック\"/><a:font script=\"Hang\" typeface=\"맑은 고딕\"/><a:font script=\"Hans\" typeface=\"等线\"/><a:font script=\"Hant\" typeface=\"新細明體\"/><a:font script=\"Arab\" typeface=\"Arial\"/><a:font script=\"Hebr\" typeface=\"Arial\"/><a:font script=\"Thai\" typeface=\"Tahoma\"/><a:font script=\"Ethi\" typeface=\"Nyala\"/><a:font script=\"Beng\" typeface=\"Vrinda\"/><a:font script=\"Gujr\" typeface=\"Shruti\"/><a:font script=\"Khmr\" typeface=\"DaunPenh\"/><a:font script=\"Knda\" typeface=\"Tunga\"/><a:font script=\"Guru\" typeface=\"Raavi\"/><a:font script=\"Cans\" typeface=\"Euphemia\"/><a:font script=\"Cher\" typeface=\"Plantagenet Cherokee\"/><a:font script=\"Yiii\" typeface=\"Microsoft Yi Baiti\"/><a:font script=\"Tibt\" typeface=\"Microsoft Himalaya\"/><a:font script=\"Thaa\" typeface=\"MV Boli\"/><a:font script=\"Deva\" typeface=\"Mangal\"/><a:font script=\"Telu\" typeface=\"Gautami\"/><a:font script=\"Taml\" typeface=\"Latha\"/><a:font script=\"Syrc\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Orya\" typeface=\"Kalinga\"/><a:font script=\"Mlym\" typeface=\"Kartika\"/><a:font script=\"Laoo\" typeface=\"DokChampa\"/><a:font script=\"Sinh\" typeface=\"Iskoola Pota\"/><a:font script=\"Mong\" typeface=\"Mongolian Baiti\"/><a:font script=\"Viet\" typeface=\"Arial\"/><a:font script=\"Uigh\" typeface=\"Microsoft Uighur\"/><a:font script=\"Geor\" typeface=\"Sylfaen\"/><a:font script=\"Armn\" typeface=\"Arial\"/><a:font script=\"Bugi\" typeface=\"Leelawadee UI\"/><a:font script=\"Bopo\" typeface=\"Microsoft JhengHei\"/><a:font script=\"Java\" typeface=\"Javanese Text\"/><a:font script=\"Lisu\" typeface=\"Segoe UI\"/><a:font script=\"Mymr\" typeface=\"Myanmar Text\"/><a:font script=\"Nkoo\" typeface=\"Ebrima\"/><a:font script=\"Olck\" typeface=\"Nirmala UI\"/><a:font script=\"Osma\" typeface=\"Ebrima\"/><a:font script=\"Phag\" typeface=\"Phagspa\"/><a:font script=\"Syrn\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Syrj\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Syre\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Sora\" typeface=\"Nirmala UI\"/><a:font script=\"Tale\" typeface=\"Microsoft Tai Le\"/><a:font script=\"Talu\" typeface=\"Microsoft New Tai Lue\"/><a:font script=\"Tfng\" typeface=\"Ebrima\"/>")
            .append("</a:minorFont>")
            .append("</a:fontScheme>")
            
            .append("<a:fmtScheme name=\"Office\">")
            .append("<a:fillStyleLst>")
            .append("<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>")
            .append("<a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"110000\"/><a:satMod val=\"105000\"/><a:tint val=\"67000\"/></a:schemeClr></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"105000\"/><a:satMod val=\"103000\"/><a:tint val=\"73000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"105000\"/><a:satMod val=\"109000\"/><a:tint val=\"81000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill>")
            .append("<a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:satMod val=\"103000\"/><a:lumMod val=\"102000\"/><a:tint val=\"94000\"/></a:schemeClr></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:satMod val=\"110000\"/><a:lumMod val=\"100000\"/><a:shade val=\"100000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"99000\"/><a:satMod val=\"120000\"/><a:shade val=\"78000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill>")
            .append("</a:fillStyleLst>")
            .append("<a:lnStyleLst>")
            .append("<a:ln w=\"6350\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln>")
            .append("<a:ln w=\"12700\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln>")
            .append("<a:ln w=\"19050\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln>")
            .append("</a:lnStyleLst>")            
            .append("<a:effectStyleLst>")
            .append("<a:effectStyle><a:effectLst/></a:effectStyle>")
            .append("<a:effectStyle><a:effectLst/></a:effectStyle>")
            .append("<a:effectStyle><a:effectLst><a:outerShdw blurRad=\"57150\" dist=\"19050\" dir=\"5400000\" algn=\"ctr\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"63000\"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle>")
            .append("</a:effectStyleLst>")
            .append("<a:bgFillStyleLst>")
            .append("<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>")
            .append("<a:solidFill><a:schemeClr val=\"phClr\"><a:tint val=\"95000\"/><a:satMod val=\"170000\"/></a:schemeClr></a:solidFill>")
            .append("<a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"93000\"/><a:satMod val=\"150000\"/><a:shade val=\"98000\"/><a:lumMod val=\"102000\"/></a:schemeClr></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:tint val=\"98000\"/><a:satMod val=\"130000\"/><a:shade val=\"90000\"/><a:lumMod val=\"103000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"63000\"/><a:satMod val=\"120000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill>")
            .append("</a:bgFillStyleLst>")
            .append("</a:fmtScheme>")
            
            .append("</a:themeElements>")
            
            .append("<a:objectDefaults/>")
            .append("<a:extraClrSchemeLst/>")
            .append("<a:extLst><a:ext uri=\"{05A4C25C-085E-4340-85A3-A5531E510DB2}\"><thm15:themeFamily xmlns:thm15=\"http://schemas.microsoft.com/office/thememl/2012/main\" name=\"Office Theme\" id=\"{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}\" vid=\"{4A3C46E8-61CC-4603-A589-7422A47A8E4A}\"/></a:ext></a:extLst>")
            
            .append("</a:theme>");
    return bldr.toString();
  }
  
  private void appendFont(StringBuilder bldr, FontDefinition fontDefn, ColourDefinition colourDefn) {
    bldr.append("<font><sz val=\"");
    bldr.append(fontDefn == null ? DEFAULT_FONT_SIZE : fontDefn.size);
    bldr.append("\"/>");
    if (colourDefn != null && colourDefn.fgColour != null) {
      bldr.append("<color rgb=\"").append(colourDefn.fgColour).append("\" />");
    }
    bldr.append("<name val=\"");
    bldr.append(fontDefn == null || fontDefn.typeface == null ? DEFAULT_FONT_NAME : fontDefn.typeface);
    bldr.append("\"/></font>");
  }
      
  private void appendFill(StringBuilder bldr, ColourDefinition colourDefn) {
    if (colourDefn != null && colourDefn.bgColour != null) {
      bldr.append("<fill><patternFill patternType=\"solid\"><fgColor rgb=\"").append(colourDefn.bgColour).append("\" /><bgColor indexed=\"64\"/></patternFill></fill>");
    } else {
      bldr.append("<fill><patternFill patternType=\"none\"/></fill>");      
    }
  }

  static String encodeSpecialCharacters(String input) {
    return input.replaceAll("&", "&amp;").replaceAll("<", "&lt;");
  }
  
  String buildStyles(TableDefinition defn) {
    StringBuilder bldr = new StringBuilder();
    bldr.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac x16r2 xr\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" xmlns:x16r2=\"http://schemas.microsoft.com/office/spreadsheetml/2015/02/main\" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\">");
    buildNumFmts(numFmtIdMap, bldr, defn);

    bldr.append("<fonts count=\"4\">");    
    appendFont(bldr, defn.bodyFont, null);
    appendFont(bldr, defn.headerFont, defn.headerColours);
    appendFont(bldr, defn.bodyFont, defn.evenColours);
    appendFont(bldr, defn.bodyFont, defn.oddColours);
    bldr.append("</fonts>");

    bldr.append("<fills count=\"5\">");
    bldr.append("<fill><patternFill patternType=\"none\"/></fill>");
    bldr.append("<fill><patternFill patternType=\"gray125\"/></fill>");
    appendFill(bldr, defn.headerColours);
    appendFill(bldr, defn.evenColours);
    appendFill(bldr, defn.oddColours);
    bldr.append("</fills>");
    
    bldr.append("<borders count=\"2\">");
    bldr.append("<border><left/><right/><top/><bottom/><diagonal/></border>");
    bldr.append("<border><left style=\"thin\"><color indexed=\"64\"/></left><right style=\"thin\"><color indexed=\"64\"/></right><top style=\"thin\"><color indexed=\"64\"/></top><bottom style=\"thin\"><color indexed=\"64\"/></bottom><diagonal/></border>");
    bldr.append("</borders>");
        
    bldr.append("<cellStyleXfs count=\"1\">");
    bldr.append("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>");
    bldr.append("</cellStyleXfs>");

    bldr.append("<cellXfs count=\"17\">");

    // Default format
    bldr.append("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>");   
    
    int borderId = defn.gridLines ? 1 : 0;
    // Column formats   
    for (ColumnDefinition col : defn.columns) {      
      int numFmt = col.format != null ? this.numFmtIdMap.get(col.format) : 0;
      bldr.append("<xf numFmtId=\"").append(numFmt).append("\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\" />");         
    }
    
    // Header
    bldr.append("<xf fontId=\"1\" fillId=\"2\" borderId=\"").append(borderId).append("\" xfId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\"/>");
    for (ColumnDefinition col : defn.columns) {      
      int numFmt = col.format != null ? this.numFmtIdMap.get(col.format) : 0;
      bldr.append("<xf numFmtId=\"").append(numFmt).append("\" fontId=\"1\" fillId=\"2\" borderId=\"").append(borderId).append("\" xfId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\"/>");
    }
    
    // Even Rows
      bldr.append("<xf fontId=\"2\" fillId=\"3\" borderId=\"").append(borderId).append("\" xfId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\"/>");
    for (ColumnDefinition col : defn.columns) {      
      int numFmt = col.format != null ? this.numFmtIdMap.get(col.format) : 0;
      bldr.append("<xf numFmtId=\"").append(numFmt).append("\" fontId=\"2\" fillId=\"3\" borderId=\"").append(borderId).append("\" xfId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\"/>");
    }

    // Odd Rows
    bldr.append("<xf fontId=\"3\" fillId=\"4\" borderId=\"").append(borderId).append("\" xfId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\"/>");
    for (ColumnDefinition col : defn.columns) {      
      int numFmt = col.format != null ? this.numFmtIdMap.get(col.format) : 0;
      bldr.append("<xf numFmtId=\"").append(numFmt).append("\" fontId=\"3\" fillId=\"4\" borderId=\"").append(borderId).append("\" xfId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\"/>");
    }
    
    bldr.append("</cellXfs>");
    
    bldr.append("<cellStyles count=\"1\">");
    bldr.append("<cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>");
    bldr.append("</cellStyles>");
    
    bldr.append("<dxfs count=\"0\"/>");
    
    bldr.append("<tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium2\" defaultPivotStyle=\"PivotStyleLight16\"/>");

    bldr.append("</styleSheet>");
    return bldr.toString();
  }

  static void buildNumFmts(Map<String, Integer> numFmtIdMap, StringBuilder bldr, TableDefinition defn) {
    List<String> numFmts = new ArrayList<>();
    for (ColumnDefinition col : defn.columns) {
      if (col.format != null) {
        if (!numFmts.contains(col.format)) {
          numFmts.add(col.format);
        }
      }
    }
    if (!numFmts.isEmpty()) {
      bldr.append("<numFmts count=\"").append(numFmts.size()).append("\">");
      int id = 165;
      for (String fmt : numFmts) {
        bldr.append("<numFmt numFmtId=\"").append(id).append("\" formatCode=\"").append(fmt).append("\"/>");      
        numFmtIdMap.put(fmt, id++);
      }
      bldr.append("</numFmts>");
    }
  }
  
  
  
}
