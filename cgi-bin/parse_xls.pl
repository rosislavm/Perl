#parse_xls.pl

sub parse_xls(){
	$filename=shift;
	# $filename="test.xls";

	$Win32::OLE::Warn = 3;

	# get already active Excel application or open new
	$Excel = Win32::OLE->GetActiveObject('Excel.Application') || Win32::OLE->new('Excel.Application', 'Quit');  

	# open Excel file
	$Book = $Excel->Workbooks->Open("$filename"); 

	# select worksheet number 1 (you can also select a worksheet by name)
	$Sheet = $Book->Worksheets(1);

	while($col){
		push(@columns,$Sheet->Cells(1,$col)->{'Value'});
	}
	while ($rows){
		my $i=1;
		for (@columns){
			my $data;
			$data->{"$_"}=$Sheet->Cells($row,$i)->{'Value'};
			push(@dataset,$data);
			$i++;
		}
	}
	$Book->Close();
	
	return @dataset;
}