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

import java.io.IOException;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.Month;
import java.time.Year;
import java.time.ZoneId;
import java.time.ZoneOffset;
import java.time.ZonedDateTime;
import java.util.Arrays;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 *
 * @author jtalbut
 */
public class XlsxWriterTest {
  
  @Test
  public void testAnyColumnSpecifiesWidth() throws IOException {
    TableDefinition defn = new TableDefinition(null, "My Data", "Jim", true, true
            , null
            , null
            , new ColourDefinition("FF0000", "00FF00") // Red, Green
            , new ColourDefinition("0000FF", "FF00FF") // Blue, Magenta
            , new ColourDefinition("FFFF00", "00FFFF") // Yellow, Cyan
            , Arrays.asList(
                    new ColumnDefinition("Index", null, null)
                    , new ColumnDefinition("Day", null, null)
                    , new ColumnDefinition("Date", "yyyy-mm-dd", 25.0)
                    , new ColumnDefinition("Fixed Format Double", "0.000", 30.0)
            )
    );
    
    XlsxWriter writer = new XlsxWriter(defn);
    assertTrue(writer.anyColumnSpecifiesWidth());

    defn = new TableDefinition(null, "My Data", "Jim", true, true
            , null
            , null
            , new ColourDefinition("FF0000", "00FF00") // Red, Green
            , new ColourDefinition("0000FF", "FF00FF") // Blue, Magenta
            , new ColourDefinition("FFFF00", "00FFFF") // Yellow, Cyan
            , Arrays.asList(
                    new ColumnDefinition("Index", null, null)
                    , new ColumnDefinition("Day", null, null)
                    , new ColumnDefinition("Date", "yyyy-mm-dd", null)
                    , new ColumnDefinition("Fixed Format Double", "0.000", null)
            )
    );
    writer = new XlsxWriter(defn);
    assertFalse(writer.anyColumnSpecifiesWidth());
  }

  @Test
  public void testToName() {
    assertEquals("A", XlsxWriter.toName(1));
    assertEquals("T", XlsxWriter.toName(20));
    assertEquals("Z", XlsxWriter.toName(26));
    assertEquals("AA", XlsxWriter.toName(27));
    assertEquals("AB", XlsxWriter.toName(28));
    assertEquals("AZ", XlsxWriter.toName(52));
    assertEquals("BA", XlsxWriter.toName(53));
  }

  @Test
  public void testTemporarlToExcelValue() {
    assertEquals("26059.421527777777", XlsxWriter.temporalToExcelValue(LocalDateTime.of(1971, Month.MAY, 6, 10, 7)));
    assertEquals("0.4217939814814815", XlsxWriter.temporalToExcelValue(LocalTime.of(10, 7, 23)));
    assertEquals("26059.421527777777", XlsxWriter.temporalToExcelValue(ZonedDateTime.of(1971, 5, 6, 10, 7, 0, 0, ZoneId.ofOffset("", ZoneOffset.ofHours(7)))));
    assertEquals("1968", XlsxWriter.temporalToExcelValue(Year.of(1968)));
  }
  

}
