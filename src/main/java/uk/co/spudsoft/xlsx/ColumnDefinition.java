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

/**
 * Details of column formatting (and header title).
 * 
 * @author jtalbut
 */
public class ColumnDefinition {
  
  /**
   * Title to use for this column if headers are enabled.
   */
  public final String name;
  
  /**
   * Excel format for the column (set for both the column and for each cell in the column).
   */
  public final String format;
  
  /**
   * Width of the column, in Excel units.
   * 
   * The formula for this is based on the number of digits that will fit and can be found <a href="https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.column?redirectedfrom=MSDN&view=openxml-2.8.1">in Microsoft's documentation</a>.
   * 
   * Summary:
   * Column width measured as the number of characters of the maximum digit width of the numbers 0, 1, 2, …, 9 as rendered in the normal style's font. There are 4 pixels of margin padding (two on each side), plus 1 pixel padding for the gridlines.
   * 
   * The "normal style's font" is the body font specified in the {@link uk.co.spudsoft.xlsx.TableDefinition} (or "Calibri" 11pt).
   * 
   */
  public final Double width;

  /**
   * Constructor.
   * 
   * @param name Title to use for this column if headers are enabled.
   * @param format Excel format for the column (set for both the column and for each cell in the column).
   * @param width Width of the column.
   */
  public ColumnDefinition(String name, String format, Double width) {
    if (width != null && width < 0.0) {
      throw new IllegalArgumentException("Width must not be negative");
    }
    if (format != null && format.isBlank()) {
      throw new IllegalArgumentException("Format must not be blank, though it may be null");
    }

    this.name = name;
    this.format = format;
    this.width = width;
  }

}
