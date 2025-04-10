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

import java.util.Collections;
import java.util.List;

/**
 * Definition of all formatting to be applied to the table.
 * 
 * It would be possible to create a more flexible API for formatting, but this should be adequate for producing a nicely formatted 'feed' type workbook.
 * 
 * 
 * @author jtalbut
 */
public class TableDefinition {

  /**
   * The application that is reported in the document properties.
   * If null the value {@link uk.co.spudsoft.xlsx.XlsxWriter#DEFAULT_APP_NAME} is used.
   */
  public final String application;
  
  /**
   * The name of the worksheet in the workbook.
   * If null the value "Sheet1" is used.
   */
  public final String name;
  
  /**
   * The name of the user creating the workbook as reported in the document properties.
   * If null the value {@link uk.co.spudsoft.xlsx.XlsxWriter#DEFAULT_APP_NAME} is used.
   * 
   */
  public final String creator;
  
  /**
   * If set to true a 'thin' (Excel's terminology) border will be applied to each cell output.
   */
  public final boolean gridLines;
  
  /**
   * If set to true a header row containing the names of each column will be generated.
   */
  public final boolean headers;
  
  /**
   * The font to use for the header row (if headers are enabled).
   * If headers are not enabled this value is ignored.
   * If this value is not specified the default value is Calibri 11pt.
   */
  public final FontDefinition headerFont;
  
  /**
   * The font to use for every row after the header row.
   * If this value is not specified the default value is Calibri 11pt.
   */
  public final FontDefinition bodyFont;
  
  /**
   * The foreground (font) and background (fill) colours to use for the header row.
   * If headers are not enabled this value is ignored.
   */
  public final ColourDefinition headerColours;

  /**
   * The foreground (font) and background (fill) colours to use for every even numbered row.
   */
  public final ColourDefinition evenColours;

  /**
   * The foreground (font) and background (fill) colours to use for every odd numbered row.
   */
  public final ColourDefinition oddColours;
  
  /**
   * Details of the columns in the output.
   * If a row contains more fields than the specified columns the output will accommodate them but they will have no header and no number format.
   */
  public final List<ColumnDefinition> columns;

  
  
  /**
   * Constructor.
   * 
   * @param application The application that is reported in the document properties.
   * @param name The name of the worksheet in the workbook.
   * @param creator The name of the user creating the workbook as reported in the document properties.
   * @param gridLines If set to true a 'thin' border will be applied to each cell output.
   * @param headers If set to true a header row containing the names of each column will be generated.
   * @param headerFont The font to use for the header row.
   * @param bodyFont The font to use for every row after the header row.
   * @param headerColours The colours to use for the header row.
   * @param evenColours The colours to use for every even numbered row.
   * @param oddColours The colours to use for every odd numbered row.
   * @param columns  Details of the columns in the output.
   */
  public TableDefinition(String application
          , String name
          , String creator
          , boolean gridLines
          , boolean headers
          , FontDefinition headerFont
          , FontDefinition bodyFont
          , ColourDefinition headerColours
          , ColourDefinition evenColours
          , ColourDefinition oddColours
          , List<ColumnDefinition> columns
  ) {
    this.application = application;
    this.name = name;
    this.creator = creator;
    this.gridLines = gridLines;
    this.headers = headers;
    this.headerFont = headerFont;
    this.bodyFont = bodyFont;
    this.headerColours = headerColours;
    this.evenColours = evenColours;
    this.oddColours = oddColours;
    this.columns = (columns == null ? Collections.emptyList() : columns);
  }
}
