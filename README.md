# Excel Writer

This project is a good starting point for anybody wanting to generate an MS Excel containing headers and tabluar data using the Apache POI library. A typical MS Excel report contains headers and data.

## Use it!
There are three basic intefaces [ExcelFileWriter](https://github.com/aashaysaralkar/excel-writer/blob/master/src/main/java/com/ekam/utilities/excelwriter/base/ExcelFileWriter.java), [HeaderWriter](https://github.com/aashaysaralkar/excel-writer/blob/master/src/main/java/com/ekam/utilities/excelwriter/base/HeaderWriter.java) and [SheetDataWriter](https://github.com/aashaysaralkar/excel-writer/blob/master/src/main/java/com/ekam/utilities/excelwriter/base/SheetDataWriter.java)

### ExcelFileWriter
This is the base writer interface and contains the below methods
```java
// Contains all the excel writing logic. Any Java class using this writer 
// should invoke this method to start the writing process
public void write(WriteConfig writeConfig) throws FileNotFoundException, IOException 

// Release all the resources a writer holds
public void close()
```
### GenericFileWriter
[GenericFileWriter](https://github.com/aashaysaralkar/excel-writer/blob/master/src/main/java/com/ekam/utilities/excelwriter/impl/GenericFileWriter.java) implements [ExcelFileWriter](https://github.com/aashaysaralkar/excel-writer/blob/master/src/main/java/com/ekam/utilities/excelwriter/base/ExcelFileWriter.java) interface and provides an implementation of the `write(...)` method which delegates writing to implementations of [HeaderWriter](https://github.com/aashaysaralkar/excel-writer/blob/master/src/main/java/com/ekam/utilities/excelwriter/base/HeaderWriter.java) and [SheetDataWriter](https://github.com/aashaysaralkar/excel-writer/blob/master/src/main/java/com/ekam/utilities/excelwriter/base/SheetDataWriter.java) 

### WriteConfig
Stores all the configurations for the writer and acts as a tranfer object to pass data between the header and content writer

### HeaderWriter
Represents a component that is responsible for writing the header information
```java
// invoked during writer initialization stage. 
// implement this method and use writeConfig for custom init logic like creating excel styles,font instances
void init(WriteConfig writeConfig)

// should contain logic to write header data
void write(WriteConfig writeConfig)

 // Should contain resource cleanup logic
void close()
```

###SheetDataWriter
Represents a component that is responsible for writing data values after the header
```java
// invoked during writer initialization stage. 
// implement this method and use writeConfig for custom init logic like creating excel styles,font instances
void init(WriteConfig writerContext);

// Used by the GenericFileWriter to decide how many sheets are to be created. 
int totalSheetsToWrite();
 
// Should contain logic to be executed before file writing starts 
void beforeFileWritingStart();
 
// Should contains logic to be executed before writing of a sheet starts
void beforeSheetWritingStart();
 
// Should contain logic to write data below headers. 
void writeSheetData(WriteConfig writeConfig);

// Should contain logic to be invoked every time writing of a sheet completes
void afterSheetWritingEnd();

// Shold contain logic to be invoked after file writing ends
	void afterFileWritingEnd();
 
 // Should contain resource cleanup logic
	void close();
 ```
 
### Prerequisites
* JAVA 1.6 or higher
* Eclipse IDE with Maven(M2E plugin)

### Installing & Running

* Import the project to your Eclipse IDE
* Do clean build
* Add a server(Tomcat/others) & add the application to it.
* Run the server.

## Running the tests

Execute the JUnit test cases in com/ekam/utilities/excelwriter/impl/GenericWriterTests.java

## Built With

* [Apache POI](http://www.dropwizard.io/1.0.2/docs/) - Library to write the file in MS Excel format
* [Maven](https://maven.apache.org/) - Dependency Management


## License

MIT

 
