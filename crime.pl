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
		  
my $fn = 'output.xlsx';
my $wb = Excel::Writer::XLSX->new($fn);

my @crimes = qw /theft auto hate assault intoxication/;

for my $crime (@crimes){ 

  my $query = "Select crime_type,
           count(report_number) as number
          From public.crime_incidents
		  Where crime_type ~* '$crime'
		  Group By crime_type
		  Order By number Desc
		  Limit 10
		  ";

  my $sth = $dbh->prepare($query);
  $sth->execute();

my $ws = $wb->add_worksheet($crime);
my $bar = $wb->add_chart( type => 'column', embedded=> 1, title => { name => 'chartzzz' });

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
  name => $crime,
  categories => '='.$crime.'!$A$1:$A$5',
  values => '='.$crime.'!$B$1:$B$5'
);
$bar->set_title(name => 'Top 5 Types of '.$crime );
$bar->set_x_axis(name => 'Type of '.$crime);
$bar->set_y_axis(name => 'Number of Crimes');
$ws->insert_chart( 'C1', $bar);
}

