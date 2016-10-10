#!/usr/bin/perl -w

#source from http://search.cpan.org/~hmbrand/Text-CSV_XS/MANIFEST
#March 2001, John McNamara, jmcnamara@cpan.org

#* you need first download Excel::Writer::XLSX moduel from 
#* http://search.cpan.org/~jmcnamara/Excel-Writer-XLSX-0.95/lib/Excel/Writer/XLSX.pm#Excel::Writer::XLSX_and_Spreadsheet::WriteExcel
#* and decompress the *.tar.gz and put this PERL file in */example/ and run the perl file with the absolutely directory

use strict;
use Getopt::Long;
use File::Basename;
use Spreadsheet::WriteExcel;
use Excel::Writer::XLSX;

my $excel;
my $help;
GetOptions(
	"excel|e:s"	=>\$excel,
	"help|h!"	=>\$help,
);

# Check for valid number of arguments
if (($#ARGV < 1) || ($#ARGV > 2) or $help) {
  print<<DO;
	Usage: perl $0 [-excel 2007/2003] <tabfile.txt> <newfile.xls>
	Option:
  excel	excel version 2007/2003 (default:2007)
	help|h	help option
	[]	optional
	<>	required and ordered
DO
exit;
};

my $tab=&ABSOLUTE_DIR($ARGV[0]);
my $out=(dirname $tab)."/".$ARGV[1];

# Open the tab delimited file
open (TABFILE, $tab) or die "$ARGV[0]: $!";
if($ARGV[1]=~/\//){
	$out=&ABSOLUTE_DIR($ARGV[1]);
}

# Create a new Excel workbook
my $workbook;
if($excel eq "2003"){
	$workbook  = Spreadsheet::WriteExcel->new($out);
}
else{
	$workbook  = Excel::Writer::XLSX->new($out);
}
my $worksheet = $workbook->add_worksheet();

# Row and column are zero indexed
my $row = 0;

while (<TABFILE>) {
    chomp;
    # Split on single tab
    my @Fld = split('\t', $_);

    my $col = 0;
    foreach my $token (@Fld) {
        $worksheet->write($row, $col, $token);
        $col++;
    }
    $row++;
}

sub ABSOLUTE_DIR{
	my $cur_dir=`pwd`;chomp($cur_dir);
	my ($in)=@_;
	my $return="";
	
	if(-f $in)
	{
		my $od=dirname($in);
		my $file=basename($in);
		chdir $od;$od=`pwd`;chomp $od;
		$return="$od/$file";
	}
	elsif(-d $in)
	{
		chdir $in;$return=`pwd`;chomp $return;
	}
	else
	{
		warn "Warning just for file and od in [sub ABSOLUTE_DIR]\n";
		exit;
	}
	
	chdir $cur_dir;
	return $return;
}
