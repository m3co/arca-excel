#!/usr/bin/perl -w

use DBI;
use Excel::Writer::XLSX;
use Config::Simple;

my $query = $ARGV[0];

$cfg = new Config::Simple();
$cfg->read('excel/db.ini'); #revisar como hacer esta direccion relativa

my $dbname = $cfg->param("Pg.dbname");
my $host = $cfg->param("Pg.host");
my $port = $cfg->param("Pg.port");
my $username = $cfg->param("Pg.username");
my $password = $cfg->param("Pg.password");

my $dbh = DBI->connect("dbi:Pg:dbname=$dbname;host=$host;port=$port",
  $username,
  $password,
  {AutoCommit => 0, RaiseError => 1, PrintError => 0}
);

my $sth = $dbh->prepare($query);

$sth->execute or die $sth->errstr;
my $fields = $sth->{NAME};

my $row = 0;
my $col = 0;

my $workbook = Excel::Writer::XLSX->new( '/tmp/resultado.xlsx' );
$worksheet = $workbook->add_worksheet();

$worksheet->write_row($row++,$col,$fields);
while(my @data = $sth->fetchrow_array)
{
  $worksheet->write_row($row++,$col,\@data);
}

$sth->finish;
$dbh->disconnect;
