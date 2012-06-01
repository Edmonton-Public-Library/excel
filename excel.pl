#!/usr/bin/perl -w
########################################################################
# Purpose: Makes educated guess at which DISCARD cards to convert.
# Method:  This script recommends candidate DISCARD cards based on user
#          Take a pipe delimited file and turn it into an excel file.
#          Currently only single page excel file output is supported.
#          -i pipe delimited file path else take from stdin
#          -t optional title for the first row
#          -o name of the output file
#          -d specific delimiter other than '|'
#          -x help
#
# Author:  Andrew Nisbet
# Date:    April 10, 2012
# Rev:     1.0 - develop
#          1.1 - April 27, 2012 - Commments added.
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


use strict;
use lib "/s/sirsi/Perl/lib/perl5/site_perl/5.8.8";
require Spreadsheet::WriteExcel;
use vars qw/ %opt /;

my $excelFile   = "excel.xls";
my $delim       = '\|';
my $inputFile;
my $colHeadings = "";
my $DEBUG       = 0;

#
# Message about this program and how to use it
#
sub usage()
{
    print STDERR << "EOF";
This program take a pipe delimited input and turns it into a single sheet MS excel file.
usage: $0 [-d delimiter] [-i input] [-o file] [-t title row] [-x]
 -d delim  : changes the delimiter from the standard '|' pipe character.
 -i file   : specifies to take input from file rather than stdin.
 -o file   : writes the output to the argument file.
 -t heading: uses delimited sting as titles for the columns.
 -x        : print help messages to stderr.
example: $0 -d ',' -i 'c:/temp/file.txt' -o 'c:/temp/out.xls' -t 'Date,Cost,Tax,Total'
EOF
    exit;
}

sub init()
{
    use Getopt::Std;
    my $opt_string = 'd:i:o:t:x';
    getopts( "$opt_string", \%opt ) or usage();
    usage() if $opt{x};
    $delim       = $opt{'d'} if $opt{d};
    $colHeadings = $opt{'t'} if $opt{t};
    $inputFile   = $opt{'i'} if $opt{i};
    $excelFile   = $opt{'o'} if $opt{o};
}

#####
# Start here
init();
# Create a new Excel workbook
my $workbook = Spreadsheet::WriteExcel->new($excelFile);
my $headingFormat = $workbook->add_format(); # Add a format
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
if ($opt{i})
{
    open(IN, "<$inputFile");
    @lines = <IN>;
}
else
{
    @lines = <STDIN>;
}

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
        $worksheet->write($rowIndex, $colIndex, $_);
        $colIndex += 1;
    }
    $rowIndex++;
}

close(IN);