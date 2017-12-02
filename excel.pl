#!/usr/bin/perl -w

use DBI;
use Excel::Writer::XLSX;

my $query = $ARGV[0];
my $dbname = '';
my $host = 'localhost';
my $port = '5432';
my $username = '';
my $password = '';

my $dbh = DBI->connect("dbi:Pg:dbname=$dbname;host=$host;port=$port",
  $username,
  $password,
  {AutoCommit => 0, RaiseError => 1, PrintError => 0}
);

my $sth = $dbh->prepare($query);

$sth->execute or die $sth->errstr;

my $row = 0;
my $col = 0;

my $workbook = Excel::Writer::XLSX->new( '/tmp/resultado.xlsx' );
$worksheet = $workbook->add_worksheet();

$worksheet->write_row($row++,$col,['Col 1', 'Col 2', 'Col 3', 'Col 4', 'Col 5', 'Col 6', 'Col 7', 'Col 8', 'Col 9', 'Col 10', 'Col 11', 'Col 12']);
while(my @data = $sth->fetchrow_array)
{
  $worksheet->write_row($row++,$col,\@data);
}

$sth->finish;
$dbh->disconnect;
