#!/usr/bin/perl -w
#use strict;

# Declare the subroutines
sub trim($);
sub ltrim($);
sub rtrim($);

my $linenum = 1;
my $operation = 1;
my $repaircode = "";
my $desc = "";
open FILE1, "repair_codes.old" or die $!;
while (my $line = <FILE1>) {
  chomp ($line);
  $line = trim($line);
  my $string  = substr $line, 0, 1;
  if($string  =~ m/;/){
    $line = " ";
  }
  if($line =~ m/(\d+)/) {
	$repaircode = $1;
  }
  if($line =~ m/'(.*)'/) {
    $desc = '"'.$1.'"';
  }
  # Format:
  # Line Operation Code "Desc"
  print "$linenum,$operation,$repaircode,$desc\n";
  open FILE2, ">>repair_codes.new" or die $!;
  print FILE2 "$linenum,$operation,$repaircode,$desc\n";
  close FILE2;
}
close(FILE1);

# Perl trim function to remove whitespace from the start and end of the string
sub trim($)
{
	my $string = shift;
	$string =~ s/^\s+//;
	$string =~ s/\s+$//;
	return $string;
}
# Left trim function to remove leading whitespace
sub ltrim($)
{
	my $string = shift;
	$string =~ s/^\s+//;
	return $string;
}
# Right trim function to remove trailing whitespace
sub rtrim($)
{
	my $string = shift;
	$string =~ s/\s+$//;
	return $string;
}