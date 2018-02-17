# bibliophilly-omeka
Code for preparing BiblioPhilly data for import into Omeka

# Instructions

1. Download "BiblioPhilly: Manuscripts included in the project" spreadsheet from Google Drive as an Excel spreadsheet, open it in Excel, save it as an XML file with the name bibliophilly-ms-in-projects.xml

2. Process bibliophilly_tei2csv.xsl against bibliophilly-ms-in-projects.xml. This process will read data from the TEI files (in the TEI folder) and combine it with data from bibliophilly-ms-in-projects.xml. Any manuscript that does not have a TEI file will have data pulled from the XML spreadsheet.

3. The resulting file spreadsheet.csv can be ingested into Omeka (using the CSV Import plugin)

The official BiblioPhilly Omeka instance is at http://bibliophilly.omeka.net/
