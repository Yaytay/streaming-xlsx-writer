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
 * Details of the font to use.
 * 
 * If the font name is not recognised Excel will keep the unknown value but seems to render it as Calibri.
 * 
 * @author jtalbut
 */
public class FontDefinition {
  
  /**
   * The name of the font.
   * If the font name is not recognised Excel will keep the unknown value but seems to render it as Calibri.
   */
  public final String typeface;
  
  /**
   * The size of the font, in points.
   */
  public final int size;

  /**
   * Constructor.
   * 
   * @param typeface The name of the font.
   * @param size The size of the font, in points.
   */
  public FontDefinition(String typeface, int size) {
    if (typeface != null && typeface.isBlank()) {
      throw new IllegalArgumentException("Typeface cannot be blank (though may be null)");
    }
    if (size < 1) {
      throw new IllegalArgumentException("Size must be positive integer");
    }

    this.typeface = typeface;
    this.size = size;
  }
  
}
