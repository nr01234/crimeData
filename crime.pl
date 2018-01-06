#!/usr/bin/perl
use strict;
use warnings;
use Text::CSV;

my $fn = "APD_Incident_Extract_YTD.csv";
my $csv = Text::CSV->new();

open(my $fh, '<', $fn) or die "Cannot open file: ";
#my $x = 0;
#while($x < 5){
#  print "$x\n";
#print $fh;
#  $x++;
#  }

my $A = "Incident Report Number";
my $B = "Crime Type";
my $C = "Date";
my $D = "Time";
my $E = "Location_Type";
my $F = "ADDRESS";

my $RBT = "Robbery By Threat";
my $AM = "AM";
my $PM = "PM";

$csv->column_names($csv->getline($fh));
print "$RBT\n";
while(my $row = $csv->getline_hr($fh)){
$row->{$D} = ($row->{$D} > 1200 ? $row->{$D} % 1200 . $PM : $row->{$D}.$AM);
#print "$row->{$D}\n";

  if($row->{$B} =~ /$RBT/i){ 
      if($row->{$D} =~ /(\d{1,2})(\d{2})([a-z]+)/i){print "$row->{$F}\t$row->{$C}\t$row->{$D}\t$1:$2 $3\n"; }}
  }

