# About streaming-xlsx-writer
The streaming-xlsx-writer is a minimal jar (no runtime dependencies)to enable the output of a single sheet XLSX file on an OutputStream.
The file is generated as it is output, there is no buffering beyond that built into a ZipOutputStream and no blocking beyond that inherent in the OutputStream.

# Build Status
![example workflow](https://github.com/Yaytay/streaming-xlsx-writer/actions/workflows/maven.yml/badge.svg)

# How to build streaming-xlsx-writer

streaming-xlsx-writer uses Maven as its build tool.

The Maven version must be 3.6.2 or later and the JDK version must be 11 or later (the jar built will be targetted to JDK 11).
