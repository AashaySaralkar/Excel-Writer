# Excel Writer

This project is a good starting point for anybody wanting to generate an MS Excel containing headers and tabular data. Base implementation of writers use Apache POI.

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
Stores all the configurations for the writer and acts as a tranfer object to pass data between the header and content writer. Data is stored as a key value pair. It contains some basic keys required for the GenericFileWriter implementation
```java
//key to store the absolute path to the location where the generated
MS Excel file is written in the WriteConfig.Required.
public static final String CONFIG_OUTPUT_FILE_LOCATION

// key to store the name of the generated MS Excel file. Required.
public static final String CONFIG_OUTPUT_FILE_NAME

// key to retrieve the current Sheet being written to
public static final String CONFIG_CURRENT_SHEET

// key to retrieve the current workbook to which sheets are being added. Value set by GenericFileWriter
public static final String CONFIG_WORKBOOK

// key to retrieve the RowCellCounter instance. Value set by GenericFileWriter
public static final String CONFIG_CURRENT_ROW_CELL_COUNTER

```

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

### SheetDataWriter
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
### RowCellCounter
Apache POI requires integer row number and cell number values while creating and accessing new cells in a Excel Sheet.Maintaining local counters for keeping track of current row and cell can be a headache.Code sprinkled with conditional counter++ can be difficult to debug. `RowCellCounter` class helps you avoid all that. It wraps current cell and row number generation.`GenericFileWriter` creates an instance of `RowCellCounter` and sets it in `WriteConfig` under `CONFIG_CURRENT_ROW_CELL_COUNTER` key. The class contains following methods:
```java
// Returns next value of the row number counter. 
public int nextRowNum()

// Returns next value for cell number counter.
public int nextCellNum()

//Resets internal counters to zero
public void reset();
```

 
### Sample use
```java
TestHeaderWriter headerWriter = new TestHeaderWriter();
TestDataWriter dataWriter = new TestDataWriter();
GenericFileWriter genericFileWriter = new GenericFileWriter(headerWriter,dataWriter);
WriteConfig config = new WriteConfig();
genericFileWriter.write(config);
genericFileWriter.close();
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
It generates a sample Excel containing basic data about airports (Many thanks to [jbrooksuk/JSON-Airports](https://github.com/jbrooksuk/JSON-Airports)) at D:\\ExcelWriter location. Kindly change output path to your desired location.

## Built With

* [Apache POI](http://www.dropwizard.io/1.0.2/docs/) - Library to write the file in MS Excel format
* [Maven](https://maven.apache.org/) - Dependency Management


## License

MIT

 
