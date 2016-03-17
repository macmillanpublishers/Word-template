Most of the useful info about specific bits of code is available in comments it in code, but a few macros have outside dependencies as detailed below.

# Bookmaker Check macro
The Bookmaker Check macro requires a CSV of the styles currently supported by Bookmaker. This file is attached to [the production downloads page](https://confluence.macmillan.com/display/PBL/Word+Template+downloads+-+production), but as of right now not via the Git for Confluence connector.


# Castoff macro
The Macmillan Castoff Macro requires at least one CSV file of design info to be uploaded as an attachment to [Word Template downloads - production](https://confluence.macmillan.com/display/PBL/Word+Template+downloads+-+production). 

NOTE: If you create and/or edit any CSV from Excel 2011 for Mac, be sure to save the file as "Windows Comma Separated (csv)" or it won't display correctly when embedded on Confluence pages (though it will still work in the macro).

## Design info CSV file
Each publisher listed as an option on the castoff form should have its own design info CSV uploaded to Confluence. The file name must consist of "Castoff_" plus the short code for that publisher used in the macro code itself. It must have both a heading row and a heading column.

Currently we are supporting the following publishers:

Publisher | Design info file
----------|------------------
St. Martin's Press | Castoff_SMP.csv
Tor.com | Castoff_torDOTcom.csv

NOTE: If no CSV is available for a publisher, the macro defaults to using the data for SMP.

### Information contained
Each MUST contain the following information for both the 5-1/2 x 8-1/4 and the 6-1/8 x 9-1/4 trim sizes, in this order.

Info | Description
-----|-------------
Design character count | The average number of characters per page for the interior text design, for a loose design, an average design, and a tight design.
Notes character count | The average number of characters per page for the interior text design of any notes or bibliography sections.
Lines per page | The standard number of lines per page for the interior text design at that trim size.
Overflow pages | The number of pages over a signature under which the castoff should round down to the nearest signature rather than round up to the next signature.
 
## Spine size CSV file
If the spine size is to be calculated as part of the castoff, a CSV needs to be uploaded that lists the spine size for each page count. It must have a heading row but no header column. Currently we only support spine sizes for POD titles, which must be named `POD_Spines.csv`.

### Information contained
Any spine size CSV must contain the following columns:

Info | Description
-----|-------------
Page count | A list of every possible page count
Spine size | The spine size at that page count, in inches