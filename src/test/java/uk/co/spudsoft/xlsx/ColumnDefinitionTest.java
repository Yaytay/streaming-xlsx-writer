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

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertThrows;

/**
 *
 * @author jtalbut
 */
public class ColumnDefinitionTest {
  
  @Test
  public void testSomeMethod() {
    new ColumnDefinition(null, null, null);
    new ColumnDefinition("Name", null, null);
    new ColumnDefinition(null, "Format", null);
    new ColumnDefinition(null, null, 11.0);
    assertThrows(IllegalArgumentException.class, () -> { new ColumnDefinition(null, null, -0.1); });
    assertThrows(IllegalArgumentException.class, () -> { new ColumnDefinition(null, "", null); });
  }
  
}
