#!/usr/bin/perl -w
########################################################################
# Purpose: Converts Symphony output into excel spreadsheet automatically.
# Method:  Perl library dependency.
#
# Converts Symphony output into excel spreadsheet automatically.
#    Copyright (C) 2014  Andrew Nisbet
#
# This program is free software; you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation; either version 2 of the License, or
# (at your option) any later version.
# 
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
# 
# You should have received a copy of the GNU General Public License
# along with this program; if not, write to the Free Software
# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston,
# MA 02110-1301, USA.
#
# Author:  Andrew Nisbet, Edmonton Public Library
# Date:    April 10, 2012
# Rev:     
#          2.3 - June 11, 2014 - Added to documentation prior to upload to community.
#          2.3 - June 6, 2014 - Fix date output.
#          2.2 - June 5, 2014 - Added date field separator character selection.
#          2.1 - June 3, 2013 - Fixed 0 date field to match excels handling (1900-01-00).
#          1.0 - develop
#          1.1 - April 27, 2012 - Commments added.
#          1.2 - June 21, 2012 - Cleaned up data input code, fixed warnings
#          that were issued when NOT using '-f' flag.
#          2.0 - July 11, 2012 - Fixed bug that issued a warning when using
#          -fn and the expected number is an empty string.
#
########################################################################

#The following methods are available through a new worksheet:
#
#write()
#write_number()
#write_string()
#write_unicode()
#write_unicode_le()
#keep_leading_zeros()
#write_blank()
#write_row()
#write_col()
#write_date_time()
#write_url()
#write_url_range()
#write_formula()
#store_formula()
#repeat_formula()
#write_comment()
#show_comments()
#add_write_handler()
#insert_bitmap()
#get_name()
#activate()
#select()
#hide()
#set_first_sheet()
#protect()
#set_selection()
#set_row()
#set_column()
#outline_settings()
#freeze_panes()
#thaw_panes()
#merge_range()
#set_zoom()
#right_to_left()
#hide_zero()
#set_tab_color()
#Border     Cell border       border          set_border()
#               Bottom border     bottom          set_bottom()


use strict;
use lib "/s/sirsi/Perl/lib/perl5/site_perl/5.8.8"; # This is the standard install on our system you mileage may vary.
require Spreadsheet::WriteExcel;
use vars qw/ %opt /;
use Switch;

my $excelFile   = "excel.xls";
my $excelField;
my $delim       = '\|';
my $inputFile;
my $colHeadings = "";
my $DATE_DELIM  = "-";
my $DEBUG       = 0;
my $VERSION     = qq{2.4};
#
# Message about this program and how to use it
#
sub usage()
{
    print STDERR << "EOF";
	usage: [[echo|cat] <api_output.txt> |] $0 [-d delimiter] [-i input] [-o file] [-t title row] [-x]
This program takes delimited Symphony output from API and converts it into a MS excel file.
Version $VERSION
 -d delim  : changes the delimiter from the standard '|' pipe character.
 -f "cols" : specifies the data types allowed for columns. Valid types are
             'g'-general, 'd'-date, 'n'-number, 'u'-url and 's'-string. The default is
             'g' so if a spreadsheet has data like |12|Andrew|1988/08/22|some text|
             you can specify -f "ngdg" for each, but the last 'g' is not expressly
             necessary.
 -i file   : specifies to take input from file rather than stdin.
 -o file   : writes the output to the argument file.
 -s char   : Alternate date field separator, like 1900/04/21. Default '-'. Note that 
             Windows will coerce the separator to the local settings of the computer
             that opens the spreadsheet for display.
 -t heading: uses delimited sting as titles for the columns.
 -x        : print help messages to stderr.
example: 
   $0 -d ',' -i 'c:/temp/file.txt' -o 'c:/temp/out.xls' -t 'Date,Cost,Tax,Total'
   echo "1|22|333|20140605|" | $0 -otest.xls -fnnnd -s'/'
   cat test.txt | $0 -o deleteme.xls -d"\\^"
EOF
    exit;
}

sub init()
{
    use Getopt::Std;
    my $opt_string = 'd:f:i:o:s:t:x';
    getopts( "$opt_string", \%opt ) or usage();
    usage() if $opt{x};
    $delim       = $opt{'d'} if $opt{'d'};
    $colHeadings = $opt{'t'} if $opt{'t'};
    $inputFile   = $opt{'i'} if $opt{'i'};
    $excelFile   = $opt{'o'} if $opt{'o'};
    $excelField  = $opt{'f'} if $opt{'f'};
    $DATE_DELIM  = $opt{'s'} if $opt{'s'};
}

#####
# Start here
init();
# Create a new Excel workbook
my $workbook = Spreadsheet::WriteExcel->new($excelFile);
my $headingFormat = $workbook->add_format(); # Add a format
my $dateFormat = $workbook->add_format(num_format => 'yyyy'.$DATE_DELIM.'mm'.$DATE_DELIM.'dd'); # Date format if used
$headingFormat->set_bold();
$headingFormat->set_align('center');
#$headingFormat->set_color('red');
my $worksheet = $workbook->addworksheet();
my @lines;
my $rowIndex = 0;
#
# Output the column headings if any.
#
if ($colHeadings ne "")
{
    # if we include a heading then the data needs to be written on
    # row 1, not 0 as we do when we have no title. Forget it and we
    # end up with an error message when we open excel.
    $rowIndex = 1;
    my $colIndex = 0;
    my @colhead = split($delim, $colHeadings);
    # don't output new line chars where we split on a pipe at the end of the line.
    # with strict we the $_ might not be initialized (it comes from command line).
    if ($_ and (ord($_) == 13 or ord($_) == 10))
    {
        next;
    }
    foreach (@colhead)
    {
        $worksheet->write(0, $colIndex, $_, $headingFormat);
        $colIndex++;
    }
}

#
# Open the appropriate input stream.
#
open(STDIN, "<$inputFile") if ($opt{i});
@lines = <STDIN>;
close(STDIN) if ($opt{i});

if ($DEBUG)
{
    print "delimiter:  '$delim'\n";
    print "colheading: '$colHeadings'\n";
    print "excelFile:  '$excelFile'\n";
    print "inputFile:  '$inputFile'\n";
    if ($opt{i})
    {
        print "input coming from file.\n";
    }
    else
    {
        print "input coming from stdin.\n";
    }
}

#
# Output the data.
#
foreach (@lines)
{
    # row and column are zero indexed.
    my $colIndex = 0;
    my @coldata = split($delim, $_);
	my @fieldTypes = ();
	@fieldTypes = split('', $excelField) if (defined($excelField));
    foreach (@coldata)
    {
        if ($DEBUG)
        {
            print "$rowIndex, $colIndex, $_\n";
        }
        # so we don't output newline chars to excel
        # chomp still allows an extra char to make it through to the excel
        # file and it shows up as a box when it should be blank.
        chomp;
		if (defined($fieldTypes[$colIndex]))
		{
			switch ($fieldTypes[$colIndex])
			{
				case "n"
				{
					if ($_ eq "")
					{
						$worksheet->write($rowIndex, $colIndex, $_);
					}
					else
					{
						$worksheet->write_number($rowIndex, $colIndex, $_);
					}
				}
				case "d"	
				{
					my @date  = split('',$_);
					if (scalar(@date) != 8)
					{
						@date = ("1","9","0","0","0","1","0","0");
					}
					my $year  = join('',@date[0..3]);
					my $month = join('',@date[4..5]);
					my $day   = join('',@date[6..7]);
					# $worksheet->write_date_time($rowIndex, $colIndex, "$year$DATE_DELIM$month$DATE_DELIM$day");
					my $dateString = sprintf "%4d-%02d-%02dT", $year, $month, $day;
					$worksheet->write_date_time($rowIndex, $colIndex, $dateString, $dateFormat);
				}
				case "u"	{$worksheet->write_url($rowIndex, $colIndex, $_);}
				case "s"	{$worksheet->write_string($rowIndex, $colIndex, $_);}
				case "f"	{$worksheet->write_formula($rowIndex, $colIndex, $_);} # experimental
				else		{$worksheet->write($rowIndex, $colIndex, $_);} # general
			}
		}
		else
		{
			$worksheet->write($rowIndex, $colIndex, $_);
		}
        $colIndex += 1;
    }
    $rowIndex++;
}

