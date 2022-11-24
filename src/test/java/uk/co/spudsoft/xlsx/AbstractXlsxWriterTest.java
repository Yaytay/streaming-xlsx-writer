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

import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

/**
 *
 * @author jtalbut
 */
public class AbstractXlsxWriterTest {
  
  private static final String[] DAYS_OF_WEEK = {"Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"};
  
  protected List<ColumnDefinition> getStandardColumnsDefns() {
    return Arrays.asList(
            new ColumnDefinition("Index", null, null)
            , new ColumnDefinition("Day", null, null)
            , new ColumnDefinition("Date", "yyyy-mm-dd", 25.0)
            , new ColumnDefinition("Date & time", "yyyy-mm-dd hh:mm:ss", 40.0)
            , new ColumnDefinition("Fixed Format Double", "0.###", 30.0)
            , new ColumnDefinition("Fixed Format Integer", "0.###", 30.0)
    );    
  }
  protected void outputFile(TableDefinition defn, final FileOutputStream fos) throws IOException, Exception {
    XlsxWriter writer = new XlsxWriter(defn);
    writer.startFile(fos);
    for (int i = 1; i < 10; ++i) {
      writer.outputRow(
              Arrays.asList(
                      i
                      , DAYS_OF_WEEK[i % DAYS_OF_WEEK.length]
                      , LocalDate.of(2022, 5, i)
                      , LocalDateTime.of(1971, 5, i, i, i)
                      , i == 4 ? null : 1.0 / i
                      , i == 2 ? null : i * i
                      , new Date()
                      , "=INDIRECT(\"A\" & ROW()) * INDIRECT(\"D\" & ROW())"
                      , i % 2 == 0
              )
      );
    }
    writer.close();
  }
  
}
