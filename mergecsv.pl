#<!-- -*- encoding: utf-8n -*- -->
use strict;
use warnings;
use utf8;

use Encode;
use Encode::JP;

use Data::Dumper;
use Cwd;

use Spreadsheet::ParseExcel;
use Spreadsheet::WriteExcel;

# ファイル入出力を制御する
use open IN  => ":encoding(cp932)";
use open OUT => ":encoding(cp932)";
# 標準入出力を制御する
binmode STDIN, ':encoding(cp932)';
binmode STDOUT, ':encoding(cp932)';
binmode STDERR, ':encoding(cp932)';

my $gCSVPattern = '[^\",]*|\"(?:[^\"]|\"\")*\"';

print "mergecsv ver. 0.13.08.19.\n";
my ($argv, $gOptions) = getOptions(\@ARGV); # オプションを抜き出す
if( 2 > scalar(@{$argv}) ){
	print "Usage: mergecsv [options] <pattern file name> <search directory>\n";
	print "  options : -expected ... expected directory. need same hierarchy of search directory\n";
	print "          : -xls      ... expected xls file. need sheet name is \"expected\"\n";
	print "          : -o        ... output file.\n";
	print "          : -dbg      ... debug mode\n";
	exit;
}

my $patternFile  = $argv->[0];
my $targetDir = $argv->[1];

# 出力ファイル
my $outFile = $patternFile;
$outFile =~ s/\e|\*|\n//g;
$outFile =~ s/\./_/g;
$outFile = $outFile.'.xls';
if( exists $gOptions->{'o'} ){
	$outFile = $gOptions->{'o'};
}

my $actualFileList = getFileList($targetDir, $patternFile);
my $actualFilesValue = getFilesValue($targetDir, $actualFileList);

my $expectedFilesValue = undef;
if( exists $gOptions->{'expected'} ){
	my $expectedFileList = getFileList($gOptions->{'expected'}, $patternFile);
	$expectedFilesValue = getFilesValue($gOptions->{'expected'}, $expectedFileList);
}
elsif( exists $gOptions->{'xls'} ){
	# エクセルからexpectedを取得する
	$expectedFilesValue = getFilesValueFromXLS( ".", $gOptions->{'xls'});
}
#print Dumper($expectedFilesValue);
#exit;

my $ssdata = makeSpreadsheetData($actualFilesValue, $expectedFilesValue);
exportExcel($outFile,$ssdata);
print "complate.\n";
exit;

sub getOptions
{
	my ($argv) = @_;
	my %options = ();
	my @newAragv = ();
	for(my $i=0; $i< @{$argv}; $i++){
		my $key = decode('cp932', $argv->[$i]);
		if( $key =~ /^-(expected|xls|o)$/ ){
			$options{$1} = decode('cp932', $argv->[$i+1]);
			$i++;
		}
		elsif( $key =~ /^-(dbg)$/ ){
			$options{$1} = 1;
		}
		elsif( $key =~ /^-/ ){
			die "illigal parameter with options ($key)";
		}
		else{
			push @newAragv, $key;
		}
	}
	return (\@newAragv, \%options);
}

sub dbg_print
{
	my ($string) = @_;
	if( defined $gOptions->{'dbg'} ){
		print encode('cp932', $string);
	}
}

# ファイルリスト取得
sub getFileList
{
	my ($path, $patternFile) = @_;
	my $path_sjis = encode('cp932', $path);

	my $regex = qr/$patternFile/;
	
	my @outFileList = ();
	opendir(DIR, $path_sjis);
	my @files_sjis = readdir(DIR);
	closedir(DIR);
	
	foreach my $file_sjis (@files_sjis){
		my $file = decode('cp932', $file_sjis);
		next if $file =~ /^\.{1,2}$/; # '.','..'は読み飛ばし
		my $filePath = $path.'\\'.$file;
		my $filePath_sjis = encode('cp932', $filePath);
		if( -d $filePath_sjis ){
			my $tempFileList = getFileList( $filePath, $patternFile);
			push @outFileList,@{$tempFileList};
			next;
		}
		if( $file =~ /$regex$/ ){
			push @outFileList, $filePath;
			dbg_print "file[$filePath]\n";
		}
	}
	return \@outFileList;
}

#my $actualFileValue = getFilesValue($actualFileList);
sub getFileValue
{
	my ($file) = @_;
	my $file_sjis = encode('cp932',$file);
	open(IN ,"<$file_sjis") or die("error[$file_sjis]:$!");
	my @array = <IN> ;
	close IN;
	my %fileValue = ();
	foreach my $line (@array){
		chomp $line;
		my $tmp = $line;
		$tmp =~ s/(?:\x0D\x0A|[\x0D\x0A])?$/,/;
		my @values = map {/^\"(.*)\"$/ ? scalar($_ = $1, s/\"\"/\"/g, $_) : $_}
		                  ($tmp =~ /(\"[^\"]*(?:\"\"[^\"]*)*\"|[^,]*),/g);
		push @{$fileValue{'keys'}} ,$values[0];
		for(my $i=1; $i<scalar(@values); $i++){
			push @{$fileValue{'data'}{$values[0]}} , $values[$i];
		}
		# 最大要素数を取得
		my $count = scalar(@values) - 1; # key 要素除く
		if( defined $fileValue{'element_count'} ){
			if( $count > $fileValue{'element_count'} ){
				$fileValue{'element_count'} = $count;
			}
		}
		else{
			$fileValue{'element_count'} = $count;
		}
	}

	return (\%fileValue);
}


sub getFilesValue
{
	my ($basePath, $fileList) = @_;
	
	my %filesValue = ();
	$filesValue{'base_path'} = $basePath;
	foreach my $file (@{$fileList}){
		$filesValue{'files'}{$file} = getFileValue($file);
		dbg_print "[$file] elemnt[$filesValue{'files'}{$file}{'element_count'}]\n";
	}
	return \%filesValue;
}


# $filesValue->{'base_path'}                             : ベースパス
# $filesValue->{'files'}{$ファイル名}{'keys'}[]          : 要素名リスト
#                                    {'data'}{$要素名}[] : 要素
#                                    {'element_count'}   : 最大要素数
sub getFilesValueFromXLS
{
	my ($basePath,$xlsfile) = @_;
	
	my %filesValue = ();
	$filesValue{'base_path'} = $basePath;
	%{$filesValue{'files'}} = ();

#	print "[$basePath][$xlsfile]\n";
	my $file_sjis = encode('cp932',$xlsfile);
	my $book = Spreadsheet::ParseExcel::Workbook->Parse($file_sjis);
	if( !$book ){
		die "[$xlsfile] not found.\n";
	}
#	print "book[$book]\n";
	my $sheet = $book->worksheet('expected');
	if( !$sheet ){
		die "[expected] sheet not found.\n";
	}
#	print "sheet[$sheet]\n";
#	print Dumper($sheet);
#	print "row[$sheet->{MaxRow}]\n";
#	print "col[$sheet->{MaxCol}]\n";
	
	# 要素名取得
	my @keys = ();
	for(my $row=1;$row<=$sheet->{MaxRow};$row++){
		my $cell = $sheet->get_cell($row,0);
		if( !defined $cell){
			last;
		}
		push @keys, $cell->value;
	}

	# ファイル名+要素取得
	# 未実装：フォーマット異常を検出する必要がある
	my @fileList = ();
	my $element_count = 0;
	my $fileName = undef;
	for(my $col=1;$col<=$sheet->{MaxCol};$col++){
		my $cell = $sheet->get_cell(0,$col);
		if( !defined $cell){
			last;
		}
		if( $cell->value ){
			# 新ファイル
			$fileName = $basePath."\\".$cell->value;
			push @fileList, $fileName;
			$element_count = 0;
			
			%{$filesValue{'files'}{$fileName}} = ();
			my $fileValue = $filesValue{'files'}{$fileName};
			push @{$fileValue->{'keys'}}, @keys;
			%{$fileValue->{'data'}} = ();
		}
		my $fileValue = $filesValue{'files'}{$fileName};
		my $data = $fileValue->{'data'};
		for(my $row=1;$row<=$sheet->{MaxRow};$row++){
			my $cell = $sheet->get_cell($row,$col);
			if( !defined $cell){
				last;
			}
			print "$fileName\:$row\:$col=$keys[$row-1]\[$element_count\],".$cell->value."\n";
			$data->{$keys[$row-1]}[$element_count] = $cell->value;
		}
		$element_count++;
		$fileValue->{'element_count'} = $element_count;
	}
#	foreach my $file(@fileList){
#		print "file:$file\n";
#	}
#	foreach my $value(@keys){
#		print "key:$value\n";
#	}
	return \%filesValue;
}



# sub getFiles
# {
# 	my ($path, $outArray, $patternFile) = @_;
# 	my $path_sjis = encode('cp932', $path);
#
# 	my %hash = ();
# 	opendir(DIR, $path_sjis);
# 	my @files_sjis = readdir(DIR);
# 	closedir(DIR);
# 	
# 	foreach my $file_sjis (@files_sjis){
# 		my $file = decode('cp932', $file_sjis);
# 		next if $file =~ /^\.{1,2}$/; # '.','..'は読み飛ばし
# 		my $filePath = $path.'\\'.$file;
# 		my $filePath_sjis = encode('cp932', $filePath);
# 		if( -d $filePath_sjis ){
# 			my %tempHash = getFiles( $filePath, $outArray, $patternFile);
# 			%hash = (%hash, %tempHash);
# 			next;
# 		}
# 		if( $file =~ /^$patternFile$/ ){
# 			push @{$outArray}, $filePath;
# 			$hash{$filePath}{'exists'} = 1;
# 			print "file[$filePath]\n";
# 		}
# 	}
# 	return %hash;
# }

#fileValue{'keys'}[]        : 要素名リスト
#         {'data'}{$要素名} : 要素
#         {'element_count'} : 最大要素数
sub isSameFileValue
{
	my ($baseFileValue, $fileValue) = @_;

	# 要素数チェック
	if( $baseFileValue->{'element_count'} != $fileValue->{'element_count'} ){
		return 0;
	}
	# 要素名数チェック
	if( scalar(@{$baseFileValue->{'keys'}}) != scalar(@{$fileValue->{'keys'}}) ){
		return 0;
	}
	# 要素名チェック
	for(my $i=0; $i<scalar(@{$baseFileValue->{'keys'}}); $i++){
		if( $baseFileValue->{'keys'}[$i] ne $fileValue->{'keys'}[$i] ){
			return 0;
		}
	}
	
	return 1;
}


#            |  title_label
# -----------+--------------
# item_label |  values
#            |
sub makeSpreadsheetData
{
	my ($actualFilesValue, $expectedFilesValue) = @_;

	my %ssd = ();
	my $isUseExpected = 0;
	if( defined $expectedFilesValue ){
		$isUseExpected = 1;
	}
	
	# $filesValue->{'base_path'}                             : ベースパス
	# $filesValue->{'files'}{$ファイル名}{'keys'}[]          : 要素名リスト
	#                                    {'data'}{$要素名}[] : 要素
	#                                    {'element_count'}   : 最大要素数
	my $baseFileValue = undef;
	my $fileCnt = 0;
	foreach my $fileName ( keys %{$actualFilesValue->{'files'}} ){
		my $relativeFileName = $fileName;
		my $tempPath = $actualFilesValue->{'base_path'};
		$tempPath =~ s/\\/\\\\/g;
		$relativeFileName =~ s/^$tempPath\\(.+?)$/$1/;
		#print "temp[$tempPath]\nrelative[$relativeFileName]\n";
		
		my $fileValue = $actualFilesValue->{'files'}{$fileName};
		if( defined $baseFileValue ){
			# keys と 要素数があっているか確認
			if( !isSameFileValue($baseFileValue, $fileValue) ){
				print "[$fileName]ファイルフォーマット異常skip...\n";
				next;
			}
		}
		else{
			# ベースファイルを先頭ファイルにする
			$baseFileValue = $fileValue;
			# item_label 作成
			foreach my $key (@{$fileValue->{'keys'}}){
				push @{$ssd{'item_label'}}, $key;
				#print "title_label=$key\n";
			}
			@{$ssd{'values'}} = ();
			@{$ssd{'values'}[0]} = ();
		}
		# title_label 作成
		push @{$ssd{'title_label'}}, $relativeFileName;
		for(my $i=0; $i<$fileValue->{'element_count'}-1;$i++){
			push @{$ssd{'title_label'}}, ''; # 要素数分進める
		}
		my $offset = $fileValue->{'element_count'} * 1 * $fileCnt;
		if( $isUseExpected ){
			# 想定結果がある場合(要素数x2の領域が足される)
			my $expectedFileName = "$expectedFilesValue->{'base_path'}\\$relativeFileName";
			my $expectedFileValue = $expectedFilesValue->{'files'}{$expectedFileName};
			#print "expectedFileName[$expectedFileName]\n";
			push @{$ssd{'title_label'}}, "expected\n$expectedFileName";
			for(my $i=0; $i<$expectedFileValue->{'element_count'}-1;$i++){
				push @{$ssd{'title_label'}}, ''; # 要素数分進める
			}
			push @{$ssd{'title_label'}}, "合否";
			for(my $i=0; $i<$expectedFileValue->{'element_count'}-1;$i++){
				push @{$ssd{'title_label'}}, ''; # 要素数分進める
			}
			$offset = $expectedFileValue->{'element_count'} * 3 * $fileCnt;
		}
		
		# values 作成
		my $row = 0;
		my $base_col= $offset;
		foreach my $key (@{$fileValue->{'keys'}}){
			my $cnt = 0;
			foreach my $value (@{$fileValue->{'data'}{$key}}){
				$ssd{'values'}[$row][$base_col+$cnt] = $value;
				$cnt++;
			}
			$row++;
		}
		
		if( $isUseExpected ){
			# 想定結果がある場合
			my $expectedFileName = "$expectedFilesValue->{'base_path'}\\$relativeFileName";
			my $expectedFileValue = $expectedFilesValue->{'files'}{$expectedFileName};
			my $row = 0;
			my $base_col= $offset + $fileValue->{'element_count'};
			foreach my $key (@{$expectedFileValue->{'keys'}}){
				my $cnt = 0;
				# 結果要素
				foreach my $value (@{$expectedFileValue->{'data'}{$key}}){
					$ssd{'values'}[$row][$base_col+$cnt] = $value;
					$cnt++;
				}
				# 合否判定
				for(my $i=0; $i<$fileValue->{'element_count'}; $i++){
					my $drow = $row + 2;
					my $act = excel_num2col(1+$offset+$i);
					my $exp = excel_num2col(1+$offset+$fileValue->{'element_count'}+$i);
					my $cmd = "=IF(EXACT($act$drow,$exp$drow)=TRUE,\"○\",\"×\")";
					$ssd{'values'}[$row][$base_col+$cnt] = $cmd;
					$cnt++;
				}
				$row++;
			}
		}

		
		$fileCnt++;
	}

# 	my $i=0;
# 	foreach my $values (@{$ssd{'values'}}){
# 		print "$ssd{'item_label'}[$i]:";
# 		foreach my $value (@{$values}){
# 			print "$value,";
# 		}
# 		print "\n";
# 		$i++;
# 	}
	
	return \%ssd;
}


#            |  title_label
# -----------+--------------
# item_label |  values
#            |
#
# $ssd->{'title_label'}[] : タイトルリスト
# $ssd->{'item_label'}[] : 要素名リスト
# $ssd->{'values'}[][] : 要素
sub exportExcel
{
	my ($file,$ssd) = @_;

	print "exportExcel[$file]\n";
	my $file_sjis = encode('cp932', $file);
	my $workbook = Spreadsheet::WriteExcel->new($file_sjis); # Create a new Excel workbook
	my $worksheet = $workbook->add_worksheet(); # Add a worksheet
	{
		# title_label
		my $format = $workbook->add_format(); # Add and define a format
		$format->set_bg_color('silver');
		$format->set_bold();
		$format->set_align('left');
		$format->set_align('top');
		$format->set_text_wrap();
		$worksheet->set_column( 1, scalar(@{$ssd->{'title_label'}}), 20);
		for(my $x=0;$x<scalar(@{$ssd->{'title_label'}});$x++){
			$worksheet->write( 0, 1+$x, $ssd->{'title_label'}[$x], $format);
		}
	}
	{
		# item_label
		my $format = $workbook->add_format(); # Add and define a format
		#$format->set_bg_color('silver');
		$format->set_bold();
		$format->set_align('left');
		$format->set_align('top');
		$format->set_text_wrap();
		$worksheet->set_column( 0, 0, 10);
		for(my $y=0;$y<scalar(@{$ssd->{'item_label'}});$y++){
			$worksheet->write( 1+$y, 0, $ssd->{'item_label'}[$y], $format);
		}
	}
	{
		# values
		my $format = $workbook->add_format(); # Add and define a format
		$format->set_align('left');
		$format->set_align('top');

# 		if(0){
# 			my $y = 1;
# 			foreach my $values (@{$ssd->{'values'}}){
# 				my $x = 1;
# 				foreach my $value (@{$values}){
# 					$worksheet->write( $y, $x, $value, $format);
# 					$x++;
# 				}
# 				$y++;
# 			}
# 		}else{
			for(my $y=0;$y<scalar(@{$ssd->{'values'}});$y++){
				for(my $x=0;$x<scalar(@{$ssd->{'values'}[$y]});$x++){
					$worksheet->write( 1+$y, 1+$x, $ssd->{'values'}[$y][$x], $format);
				}
			}
#		}
	}
	$workbook->close;
}

# num to A-Z for excel.
sub excel_num2col
{
	my ($bb) = @_;
	
	my @dst = ();
	do {
		my $mod = ($bb % 26);
		push @dst, unpack("C", 'A') + $mod;
		$bb = int( $bb / 26) - 1;
	}while( $bb >= 0 );

#	return join '',@dst;
# #	printf("[%s][%s]\n",chr($dst[0]),chr($dst[1]));
 	my $line = '';
 	foreach my $str (@dst){
 		$line = chr($str).$line;
 	}
# #	printf("sub : $line\n");
	return $line;
}

#----------------------------------------------------------------
# A-Z to num for excel.
sub excel_col2num
{
	my ($a) = @_;
	my @strs = split //, $a;
	my $digit = 0;
	my $ans = 0;
	foreach my $num (@strs){
		my $tmp = unpack("C", $num) - unpack("C", 'A') + 1;
		$ans = ($ans * 26);
#		print "$digit : \($num == $tmp\) += $ans\n";
		$ans += $tmp;
		$digit++;
	}
	return $ans - 1;
#	print "ans = $ans\n";
}

#----------------------------------------------------------------
# col = col + col
sub excel_add_col
{
	my ($a,$b) = @_;

	my $aa = excel_col2num($a);
	my $bb = excel_col2num($b);
	my $ans = $aa + $bb;
	$ans = excel_num2col($ans);
	return $ans;
}

#----------------------------------------------------------------
# col = col + num
sub excel_add_num
{
	my ($a,$b) = @_;

	my $aa = excel_col2num($a);
	my $ans = $aa += $b;
	$ans = excel_num2col($ans);
	return $ans;
}

