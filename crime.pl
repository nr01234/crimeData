#!/usr/bin/perl
use strict;
use warnings;
use Excel::Writer::XLSX;
use DBI;
use DBD::Pg;

##Database Connect##
my $db = "data";
my $host = "";
my $port = "5432";
my $driver = "Pg";
my $dsn = "DBI:$driver:database=$db;host=$host;port=$port";
my $user = "";
my $pswd = "";

my $dbh = DBI->connect($dsn,$user,$pswd) 
          or die "Can't connect to Database: $DBI::errstr\n";

sub qwriter
{
  return "Select crime_type,
                 count(report_number)
		  From public.crime_incidents
		  Group By crime_type
				 ";
}


my $query = qwriter();		  

my $sth = $dbh->prepare($query);
$sth->execute();
my $fn = 'output.xlsx';
my $wb = Excel::Writer::XLSX->new($fn);

my $row = 0;

while(my @fields = $sth->fetchrow_array())
{
  for (@fields) { 
    my $size = @fields;
	}
  $row++;
}

