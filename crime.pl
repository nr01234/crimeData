#!/usr/bin/perl
use strict;
use warnings;
use DBI;
use DBD::Pg;
use Excel::Writer::XLSX;

##Columns##
my $A = "Incident Report Number";
my $B = "Crime Type";
my $BB = "Number of Crimes";
my $C = "Date";
my $D = "Time";
my $E = "Location_Type";
my $F = "ADDRESS";

##Crimes##
my $RBT = "Robbery By Threat";
my $theft = "theft";

##########
my $AM = "AM";
my $PM = "PM";


##Database Connect##
my $db = "austin_data";
my $host = "";
my $port = "5432";
my $driver = "Pg";
my $dsn = "DBI:$driver:database=$db;host=$host;port=$port";
my $user = "postgres";
my $pswd = "postgres";

my $dbh = DBI->connect($dsn,$user,$pswd) 
          or die "Can't connect to Database: $DBI::errstr\n";

my $crimes = qw /theft auto hate/;

sub qwriter
{
  return "Select crime_type,
           count(report_number)
		  From public.crime_incidents
		  Where crime_type ~* '$theft'
		  Group By crime_type
				 ";
}

sub x
{
  return "Select crime_type,
           count(report_number) as number
          From public.crime_incidents
		  Where crime_type ~* '$theft'
		  Group By crime_type
		  Order By number Desc
		  Limit 10
		  ";
}


my $query = qwriter();	
$query = x();	  

my $sth = $dbh->prepare($query);
$sth->execute();
my $fn = 'output.xlsx';
my $wb = Excel::Writer::XLSX->new($fn);
my $ws = $wb->add_worksheet('Theft');
my $bar = $wb->add_chart( type => 'column', name => 'Top 10 Types of Theft', embedded=> 1, title => { name => 'chartzzz' });

$ws->set_column(0, 0, 35);


my $row = 0;
my @header = ($B,$BB);

while(my @fields = $sth->fetchrow_array())
{
  for (@fields) { 
    my $last = @fields - 1;
	foreach my $col (0..$last){ $ws->write($row, $col,$fields[$col]); }
	}
  $row++;
}

$bar->add_series(
  categories => '=sheet1!$A$1:$A$5',
  values => '=sheet1!$B$1:$B$5'
);

$ws->insert_chart( 'C1', $bar);

