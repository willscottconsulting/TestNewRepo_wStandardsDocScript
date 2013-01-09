#!/usr/bin/perl -w

###################################################################################
# This script prints details from an Excel spreadsheet to a txt file. It is       #
# intended to work with Field-level data standards spreadsheets at Penn Museum to #
# provide HTML formatting that can be copied into WordPress.                      #
# While those files have a specific column structure, options are provided to     #
# allow some flexibility with those and possibly with other Excel files.          #
###################################################################################

# Revision and Version Details                                                    
# Created by: Will Scott 2/26/2012                                                
# Version 1.1, 2/29/2012, Changes:
# -Added @endreport that includes notice of critical blank values in Named Anchors.
# -Tested on live Excel data standards file
# -Fixed errors created by empty cell values being treated as undefined.
# -Added additional comments and documentation.

use strict;
use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';

$Win32::OLE::Warn = 3;                                # die on errors...

# Get Active Excel Application.
my $Excel = Win32::OLE->GetActiveObject('Excel.Application');

###########################################################################
# BASIC CONFIGURATION: Environment & File variables, column selections    #
###########################################################################

# Set file name and location, define a target path and filename for the txt results file.
my $filepath='E:\HPPavilionBackup9Sept2011_FULL\Users\Will\Documents\Business\UPM\Migration\DataStandardsEMu';
my $filename='CatalogFieldDocumentation_OngoingRevs.xlsx';     #NOTE!!! 17May2011 is an old version for testing, move to ongoingrevs file for final.
my $targetfilepath='E:\HPPavilionBackup9Sept2011_FULL\Users\Will\Documents\Business\UPM\Migration\DataStandardsEMu\DataStandards_FieldLevelScript\Target_ResultFiles';
my $targetfilename='DocScriptResults.txt';
my $worksheetname='Catalog'; # The name of the worksheet/tab within the Excel file to use.

# Configure default column selection & order and the columns & orderto use in compiling a named anchor.

my @defaultfieldselection=(1,2,5,6,10);
my @namedanchorcols=(3,2,1);   #Note there is no input option to change this. Any changes must be made here.

# open Excel file
my $Book = $Excel->Workbooks->Open("${filepath}\\${filename}"); 

# Select worksheet (you can also do this by number rather than name).
# NOTE this will presently only work for a hard-coded worksheet name as configured with other environment & file variables above.
# Only one worksheet can be used at a time (unless further dev work allows additional ones). Worksheets can be referenced here by number
# rather than by name, in case that is useful in future development or working around whitespace issues.

my $Sheet = $Book->Worksheets($worksheetname);

# Find Last Column and Row. Note this may have no tolerance for completely empty cols or rows in the context of data. If a column is unused, it should
# at least have a Label in the first row.

my $LastRow = $Sheet->UsedRange->Find({What=>"*",
    SearchDirection=>xlPrevious,
    SearchOrder=>xlByRows})->{Row};

my $LastCol = $Sheet->UsedRange->Find({What=>"*", 
                  SearchDirection=>xlPrevious,
                  SearchOrder=>xlByColumns})->{Column};

#Report range limits found to screen for testing and confirmation.
printf "\n--LastRow and LastCol Values--\n";
printf "LastRow: $LastRow\n";
printf "LastCol: $LastCol\n\n";


######################################################################################
# User Input: Get column selection, allow for a filter, verify overwrite of txt file #
######################################################################################

# Show some information to the user on screen, including an option to select field order
printf "NOTE: This script likely has no tolerance for rows or columns left blank in the context of data. A field label must be included in the source file at the least.\n\n";
printf "These are the Column Names in your Spreadsheet based on the top row:\n\n";

foreach my $namcol (1..$LastCol)
{
    printf "$namcol -";
    printf $Sheet->Cells(1,$namcol)->{'Value'};
    printf "\n";
}

printf "\nThe default fields order in the script are: @defaultfieldselection\n\n";
printf "Enter D to use the default or enter field numbers in desired order separated by commas for a custom selection \n\n==> ";

# Declare some variables for processing column selection and named anchor compilation.
my @usecolarr;
my @inputarr;
my $val;
my $i;
my $namedanchor='';  #Named Anchor string, compiled.
my $naraw;   #Named Anchor Raw.
my @endreport; # An array of messages provided upon script completion noting any items that may need attention.

# Get and process user input for column selection. Input validation is not great here, DEV: request a valid option instead of dying if !valid.
my $input = <STDIN>;
chomp($input);

if (!$input)
    {
        die "\nNOTHING ENTERED! \nScript Terminated\nRun the script again and select an option.\n";
    }
    
elsif ($input ne "D" && $input ne "d")
    {
    
    @inputarr = split(',', $input);
  
        # Convert input values to integers. OLE requires this. Also checks that all numbers are less than LastCol.
        foreach $i(@inputarr)
        {
            if ($i=~/\D/) {die("ERROR: Value $i could not be converted to an integer");}  #Checks for any non-numeric character including period/decimal.
            $val = int($i);
            if ($val>$LastCol) {die("ERROR: Value $val is greater than the number of used columns in LastCol.");}
            @usecolarr=(@usecolarr,$val); # Inserts valid values into an array.
        }
    
    # Display the selection on screen if not the default.
    printf "\nYou have selected @usecolarr \n\n";
    @endreport=(@endreport,"\nCustom Columns Selected: @usecolarr\n");
        
    }

else 
    {
    print "\nOK, I'll use the Default order\n\n";
    @usecolarr = @defaultfieldselection;
    
    @endreport=(@endreport,"\nDefault Columns Selected: @usecolarr\n");
    }

# Allow option of a single column / single value FILTER based on stdin. This is primarily here to allow for filtering by
# Tab names. Checking of input data is very bare-bones, so the data needs to be clean.

# Declare some variables for processing filter input.
my $filtercolin;
my $filtercol;
my $filtervalin;
my $filterval;
my $usefilter=0; # Determines whether the filter option was selected. Somewhat redundant, but used for clarity.
my $celldata;  # Used in processing the filter specifically for defining undefined values as empty strings.
my @filterrows;  # Row numbers that match the defined filter.
my @userows;  # Either filterrows or full range of available rows for use in scrolling through data for printing.

# Get user input of filter specifications if any and process. Input value checking is dumb here. Basically it dies on any invalid value.
# Could be much better, but does the trick for now.
printf "\nTo add a filter for a specific value in a column, enter the column number\nOr just press Enter to continue.\n\n==>";
$filtercolin=<STDIN>;
chomp($filtercolin);

if ($filtercolin)
{
# Value must be clearly set as an integer.
if ($filtercolin=~/\D/) {die("ERROR: Value $filtercolin could not be converted to an integer");}  #Checks for any non-numeric character including period/decimal.
$filtercol = int($filtercolin);
if ($filtercolin>$LastCol) {die("ERROR: Filter Column selected is not valid - exceeds LastCol value.");}
else
{
$filtercol=$filtercolin;

printf "\n\nWhat value in Column $filtercol would you like to include?\n(Must match exactly the header, whitespace is OK, quotes are not required)\n\n==>";
$filtervalin=<STDIN>;
chomp($filtervalin);
if (!$filtervalin) {die("Filter option selected but no filter value provided.");}
$filterval=$filtervalin;
$usefilter=1;

@endreport=(@endreport,"Filter Selected on Column: $filtercol, Value: $filterval\n");
}
} # I belong to the IF ($filtercolin)condition.


#############################################################################
# Processing: Loop through rows and columns. Print the results to a file.   #
#############################################################################

#Open a file to which to write the results.

my $owverif;  #Overwrite verification, set by user input... confirms that it is ok to clear previous text from the target file.

printf "\n-----File Overwrite Verification-----\n";
printf "Target File: ${targetfilename}\n";
printf "Target Directory: ${targetfilepath}\n\n";
printf "Any data in the specified file will be overwritten!\n\nPress any key to continue, or N to cancel.\n\n==>";

$owverif=<STDIN>;
chomp($owverif);
if ($owverif eq "N" || $owverif eq "n") {die("Script stopped by user at file overwrite verification.");}

open (TARGETFILE, ">${targetfilepath}\\${targetfilename}");
print TARGETFILE;  #Overwrites any data in the file, by writing nothing to it.

# Loop through rows and columns and print formatted information to the txt file.

# If a filter was set, loop through for matches first, read the matching rows into an array.

foreach my $row (2..$LastRow)   #Note that the first row will always contain headers/field labels, so is skipped in the loop, but still accessible.
{
    if ($usefilter==1)  #Processing if a filter option was selected.
    {
    $celldata=$Sheet->Cells($row, $filtercol)->{'Value'};
    if (!$celldata)  # Force an empty string if the call returns nothing. Empty cells are handled as undefined and causes errors.
        {
            $celldata='';
        } 
    if ($celldata eq $filterval)
    {@filterrows=(@filterrows,$row);}  # Add all matching rows to an array to be used when looping through the data to print.
    
    }
}

# Figure out which rows to use in printing, filtered rows or the full range.

if (!@filterrows && $usefilter==1)  # Filter selected, but found no rows.
{
    printf "\nNO ROWS FOUND MATCHING THE SPECIFIED FILTER.\n";
}
elsif (@filterrows && $usefilter==1)   # Filter selected, and DID find rows
{
    @userows=@filterrows;
}
elsif ($usefilter==0)    # Filter not selected, use the full range.
{
    @userows=(2..$LastRow)   
}
else
{
    die ("Script stopped. Could not determine what rows to include based either on matching filter rows or a range of row 2 to LastRow.");
}


# Loop through rows (either filtered or all), build named anchors, and print to screen and to file.

foreach my $row (@userows)
{
    # Generate Named Anchors.     
    foreach my $nacol(@namedanchorcols)
        {
            # Skip blank cells.
            # next unless defined $Sheet->Cells($row,$nacol);
            
            if (!$Sheet->Cells($row, $nacol)->{'Value'})
                {
                    $naraw='BLANKNAMEDANCHORFIELD';
                    @endreport=(@endreport,"\nNEEDS ATTENTION: Empty Named Anchor value found and given a default - Row: $row, Col: $nacol\n");
                }
            else {$naraw=$Sheet->Cells($row, $nacol)->{'Value'};}
            
            #Replace some punctuation and any white space with an underscore.
            $naraw=~s/[\.,-\/#!$%\^&\*;:{}=\-`~()\s+]/_/g;
            $namedanchor.=$naraw;
        }
    
    printf "<a name=\"".$namedanchor."\"></a>\n";
    print TARGETFILE "<a name=\"".$namedanchor."\"</a>\n";
    
    # Reset the named anchor variable to empty for processing the next row.
    $namedanchor='';
  
 foreach my $col(@usecolarr)
 {
    
  # Skip blank cells. From sample, not used here, but may be eventually.
  # next unless defined $Sheet->Cells($row,$col);
  
  # Here we generate HTML from the contents.
  
    # Print formatted values to screen for verification.   
    printf '<strong>';  
    printf $Sheet->Cells(1, $col)->{'Value'};   # Prints labels regardless of the presence of data. 
    printf '</strong>: ';
    if ($Sheet->Cells($row, $col)->{'Value'})   # Must check that value is defined or errors.
        {
            printf $Sheet->Cells($row, $col)->{'Value'};
        }
   
    printf "\n\n";
    
    # Print formatted values to file. Note that newlines are included, but no specifically-defined <BR> tags. Wordpress
    # interprets newlines correctly. 
    print TARGETFILE '<strong>';  
    print TARGETFILE $Sheet->Cells(1, $col)->{'Value'};    # Prints labels regardless of the presence of data. 
    print TARGETFILE '</strong>: ';
    if ($Sheet->Cells($row, $col)->{'Value'}) {print TARGETFILE $Sheet->Cells($row, $col)->{'Value'};}  # Must check that value is defined or errors.
    print TARGETFILE "\n";
      
 }
 print TARGETFILE "\n"; # Add a newline to the end of each row. The last of these could be stripped eventually, but not urgent.
 
} # I belong to the foreach row loop for printing.

printf "\nScript Completion Report: @endreport";  # Print the script completion report including any errors that need attention.

# clean up 
close TARGETFILE;
$Book->Close;

