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

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;

/**
 *
 * @author jtalbut
 */
public class ColourDefinitionTest {
  
  @Test
  public void testCheckColour() {
    
    Assertions.assertThrows(IllegalArgumentException.class, () -> { ColourDefinition.checkColour(""); });
    Assertions.assertThrows(IllegalArgumentException.class, () -> { ColourDefinition.checkColour("A"); });
    assertEquals("FF123456", ColourDefinition.checkColour("123456"));
    assertNull(ColourDefinition.checkColour(null));
    assertEquals("FF123456", ColourDefinition.checkColour("FF123456"));
    
  }
  
}
