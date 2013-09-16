# plum spreadsheet converter

This simple application allows editing of plum database files using 
standard database tools by converting them to spreadsheets which may be
edited and then converted back in plum database text files.

## Compilation

### Build Dependencies

* java (version 5 or higher)
* maven (version 3 or higher)

To build the application, simply type: 

	mvn package

	
## Running

	cd target
	java -jar plum-0.0.1-SNAPSHOT-jar-with-dependencies.jar

On certain platforms, you may simply be able to double-click on the jar
file in the application.

## Usage

The application will load a simple file selector dialog.  Either select
plum database files that will be converted to Excel spreadsheets or 
select spreadsheets that will be converted back.

### Caveats

* All cells in the spreadsheet are formatted as strings (whether they're numeric or not) and should be left that way.
* The spreadsheet rows should be sorted by "original order" column if order is to be retained.
