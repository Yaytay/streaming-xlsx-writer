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

import java.io.File;
import java.io.FileOutputStream;
import org.junit.jupiter.api.Test;

/**
 *
 * @author jtalbut
 */
public class XlsxWriterNoBordersTest extends AbstractXlsxWriterTest {
  
  @Test
  public void testFile() throws Exception {
    
    File file = new File("target/temp/XlsxWriterNoBordersTest.xlsx");
    file.getParentFile().mkdirs();

    try (FileOutputStream fos = new FileOutputStream(file)) {      
      TableDefinition defn = new TableDefinition(null, "My Data", "Jim", false, true
              , null
              , null
              , new ColourDefinition("FF0000", "00FF00")
              , new ColourDefinition("0000FF", "FF00FF")
              , new ColourDefinition("FFFF00", "FFFFFF")
              , getStandardColumnsDefns()
      );
      
      outputFile(defn, fos);      
    }
    
  }

  
}
