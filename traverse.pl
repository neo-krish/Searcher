use File::Find qw(find);
use Spreadsheet::ParseExcel;
use strict;
use warnings;

package main;
my $fileCount = 0;			# To keep count of number of files in the folders
my $excelFileCount = 0;		# To keep count of number of .XLS files in folders
my $excelFilesWithCDM = 0;	# To keep count of number of .XLS files with CDM in their filename
my $wordFileCount = 0;		# To keep count of number of .DOC files in folders
my $dirCount = 0;			# To keep count of number of directories
my $percentCDM;				# Stores the % value of CMD XL vs XL files.

my $parser = Spreadsheet::ParseExcel->new();	# Parser object to parse excel files
my $currentWorkbook;							# Handle for the excel file being processed
my @worksheets;									# Array to hold the worksheets in the excel file
my $numOfWorksheets;							# To keep count of number of worksheets in the excel file
my $worksheetToProcess;							# Handle for the worksheet that needs to be checked for data
my %excelRowColCount;							# Hash to hold the excel file name and the number of Rows and Columns present in it

# Starting Directory
my $myDir = '/Users/narayanavenkatesh/Downloads/Perl script project/California/2006';

# Subroutine that performs function on each File found
sub readEachFile {

	# Check if the current items being processed is a file
	if (!-d){
		# Check if the current file is an excel sheet
		if ($_ =~ /.*xls$/ ){ 
			$excelFileCount++;

			# Check if the excel sheet has 'CDM' or 'cdm' in its filename
			if (($_ =~ /.*CDM*/) || ($_ =~ /.*cdm*/)){
				$excelFilesWithCDM++;
				# Open the excel file with CDM in its filename
				$currentWorkbook = $parser->parse($File::Find::name);
				# If the file can't be opened, die gracefully
				if ( !defined $currentWorkbook ) {
					die $parser->error(), ".\n";
				}
				# Store all the worksheets in the excel file in an array
				@worksheets = $currentWorkbook->worksheets();
				# Find the number of worksheets
				$numOfWorksheets = $currentWorkbook->worksheet_count();
				#print $numOfWorksheets, "\n";
				# If the number of excel sheets is more than 1, pick the one with CDM / cdm in its name
				foreach my $currentWorkSheet (@worksheets) {
					if (($currentWorkSheet->get_name() =~ /.*CDM*/) || ($currentWorkSheet->get_name() =~ /.*cdm*/)){

						# Array to keep the Row and Column count for each excel file (selected worksheet)
						my @worksheetRowColStat;
						# Fetch the Max and Min value for the rows used
						my ( $row_min, $row_max ) = $currentWorkSheet->row_range();

						# Function row_range returns row_max less than row_min if there are no rows used in the excel sheet
						if ($row_max > $row_min) {
							$worksheetRowColStat[0] = $row_max - $row_min;
							} else {
								$worksheetRowColStat[0] = 0;
							}
							# Fetch the Max and Min value for the columns used
							my ( $col_min, $col_max ) = $currentWorkSheet->col_range();

							# Function col_range returns col_max less than col_min if there are no rows used in the excel sheet
							if ($col_max > $col_min) {
								$worksheetRowColStat[1] = $col_max - $col_min;
								} else {
									$worksheetRowColStat[1] = 0;
								}

								# Store the information in the hash
								push(@{$excelRowColCount{$_}}, @worksheetRowColStat);

								# Leave this loop soon as a worksheet is found with CDM / cdm in its name
								last;
							}
						}
					}
				}
				# Check if the current file is an word document
				if ($_ =~ /.*doc$/){
					$wordFileCount++;
				}
			}
			# If the item is a Directory
			else
			{
				$dirCount++;
			}
		}

		# Traverse the folder structure
		find(\&readEachFile, $myDir);

		# Housekeeping calculations
		$percentCDM = ($excelFilesWithCDM/$excelFileCount)*100;

		# Outputs
		printf ("The number of excel files in the folders is %d\n", $excelFileCount);
		printf ("The number of excel files with CDM / cdm in their filename is %d\n", $excelFilesWithCDM);
		printf ("The Percent of excel files with CDM in file name is %d\n", $percentCDM);
		printf ("The number of word documents in the folders is %d\n", $wordFileCount);
		printf ("The number of directories is %d\n", $dirCount);
		#print %excelRowColCount, "\n";
		print "Excel File\t\t\t\t\t\t\t\t\t\t\tRows		Columns\n";
		foreach my $key (keys %excelRowColCount) {
			my @rowColDetail = @{$excelRowColCount{$key}};
			print "$key\t\t\t\t\t\t\t\t\t\t\t$rowColDetail[0]	$rowColDetail[1]";
			print "\n";
		}
