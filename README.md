# About streaming-xlsx-writer
The streaming-xlsx-writer is a minimal jar (no runtime dependencies)to enable the output of a single sheet XLSX file on an OutputStream.
The file is generated as it is output, there is no buffering beyond that built into a ZipOutputStream and no blocking beyond that inherent in the OutputStream.

# Build Status
![example workflow](https://github.com/Yaytay/streaming-xlsx-writer/actions/workflows/maven.yml/badge.svg)

# How to build streaming-xlsx-writer
streaming-xlsx-writer uses Maven as its build tool.

The Maven version must be 3.6.2 or later and the JDK version must be 11 or later (the jar built will be targetted to JDK 11).

# Usage
The basic usage pattern is:
1. Create an OutputStream.
2. Create a TableDefinition object to define the formatting you want.
3. Create the XlsxWriter, passing in the TableDefinition.
4. Call startFile, passing in the OutputStream, to output all of the file metadata.
5. Call outputRow for each row required on the worksheet.
6. Call close to end the XLSX document.
7. Call close to close the OutputStream.

It is recommended that try-with-resources be used for closing the XlsxWriter and the OutputStream if possible.

```java
    String[] DAYS_OF_WEEK = {"Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"};

    File file = new File("target/temp/XlsxWriterFileTest.xlsx");
    file.getParentFile().mkdirs();

    try (FileOutputStream fos = new FileOutputStream(file)) {      
      TableDefinition defn = new TableDefinition(null, "My Data", "Jim", true, true
              , new FontDefinition("Comic Sans MS", 15)
              , new FontDefinition("Consolas", 15)
              , new ColourDefinition("FF0000", "00FF00") // Red, Green
              , new ColourDefinition("0000FF", "FF00FF") // Blue, Magenta
              , new ColourDefinition("FFFF00", "00FFFF") // Yellow, Cyan
              , getStandardColumnsDefns()
      );
      
      try (XlsxWriter writer = new XlsxWriter(defn)) {
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
                  )
          );
        }
      }
    }
```