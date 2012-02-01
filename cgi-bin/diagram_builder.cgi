#!"E:/xampp/perl/bin/perl.exe"
push(@INC,"./");

use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';
require("parse_xls.pl");
$query=&query_parser();

if($query->{"page"} eq "diagram" and $query->{"file"}){
	$toswap->{"wtitle"}="diagram";
	$toswap->{"title"}="Diagram based on the values in the file";
	
	$diagram_info=&parse_xls($query->{"file"});		#here I resive the onfo needed to draw the Chart
	#and then I will make the data suitable for drawing the Charts
	
	&makepage("diagram");
}elsif($query->{"page"} eq "diagram_settings"){
	$toswap->{"wtitle"}="diagram settings";
	$toswap->{"title"}="Asign the settings for the diagram";
	&makepage("diagram_settings");
}else{
	$toswap->{"wtitle"}="welcome";
	$toswap->{"title"}="Welcome";
	&makepage("welcome");
}


print "Content-Type: text/html\n\n";
swap($toswap,"../htdocs/s-html/template.html");

sub swap(){
	my ($hash,$filename)=@_;
	open(FILE,$filename);
	while(<FILE>){
		$_=~ s/%(\w+)/$hash->{$1} || ""/ge;
		print($_);
	}
	close(FILE);
}

sub makepage{
	$toswap->{"content"}=&swap_string($page,$html_template{"$_[0]"});
}

sub swap_string{
	my ($hash,$string)=@_;
	$string=~ s/%(\w+)/$hash->{$1} || ""/ge;
	return $string;
}

sub query_parser(){
	my $query=$ENV{"QUERY_STRING"};
	$query=~ s/&amp;/&/gi;
	$query=~ tr/+/ /;
	my @aquery=split(/&/,$query);
	for($i=0;$i <=$#aquery;$i++){
		($name,$value)=split(/=/,$aquery[$i]);
		# $name=~ s/%(..)/pack("c",hex($1))/ge;
		# $value=~ s/%(..)/pack("c",hex($1))/ge;
		$value=~ s/\'/\&\#39\;/g;
		$value=~ s/\\/\&\#92\;/g;
		$hquery{$name} and $hquery{$name}.=";; $value" or $hquery{$name}=$value;
	}
	return \%hquery;
}

sub html_templates(){
	$html_temlate->{"diagram"}="";						#the diagram I would prefer make a Flot Chart if this is possible :)
	$html_temlate->{"diagram_settings"}="";				#menu for inserting a file and some possible options for the chart
	$html_temlate->{"welcome"}="Добре Дошли :)";		#a welcoming screen
}