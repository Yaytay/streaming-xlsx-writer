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

import java.util.regex.Pattern;

/**
 * Represents the colour that will be used in the output.
 * 
 * The colours must be specified as six or eight character [A]RGB hex strings.
 * The Alpha channel appears to be ignored by Excel, but the XML is technically invalid without it, so it should always be set to "FF".
 * 
 * @author jtalbut
 */
public class ColourDefinition {
  
  /**
   * Regular expression for validating [A]RGB colour strings (6 or 8 uppercase hexadecimal values).
   */
  public static final Pattern VALID_COLOUR = Pattern.compile("^[0-9A-F]{6}([0-9A-F]{2})?$");
  
  /**
   * The colour to be used for text.
   */
  public final String fgColour;
  
  /**
   * The colour to be used for the background fill. 
   */
  public final String bgColour;

  /**
   * Constructor.
   * @param fgColour The colour to be used for text.
   * @param bgColour The colour to be used for the background fill. 
   */
  public ColourDefinition(String fgColour, String bgColour) {
    this.fgColour = checkColour(fgColour);
    this.bgColour = checkColour(bgColour);
  }
  
  static final String checkColour(String colour) {
    if (colour == null) {
      return null;
    }
    if (!VALID_COLOUR.matcher(colour).matches()) {
      throw new IllegalArgumentException("Colours must be either 6 or 8 character RGB strings");
    }
    if (colour.length() == 6) {
      return "FF" + colour;
    }
    return colour;
  }
    
}
