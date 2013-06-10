#<!-- -*- encoding: utf-8n -*- -->
use strict;
use warnings;
use utf8;

use Encode;
use Encode::JP;

# use Data::Dumper;
# use Cwd;

use Spreadsheet::WriteExcel;

# ファイル入出力を制御する
use open IN  => ":encoding(cp932)";
use open OUT => ":encoding(cp932)";
# 標準入出力を制御する
binmode STDIN, ':encoding(cp932)';
binmode STDOUT, ':encoding(cp932)';
binmode STDERR, ':encoding(cp932)';

my $gCSVPattern = '[^\",]*|\"(?:[^\"]|\"\")*\"';

my $args = @ARGV;
if( $args < 2 ){
    warn "Usage: margecsv <パターンファイル名> <検索開始ディレクトリ>\n";
    exit;
}

my $patternFile  = decode('cp932', $ARGV[0] ); # in file # 入力アーギュメントは文字コードを変更してやらないとダメっぽい。
my $targetDir = decode('cp932', $ARGV[1] ); # out file
my $outFile = $patternFile;
$outFile =~ s/\e|\*|\n//g;
$outFile =~ s/\./_/g;
$outFile = $outFile.'.xls';

print "パターンファイル名[$patternFile]\n";
print "検索開始ディレクトリ[$targetDir]\n";
print "出力ファイル名[$outFile]\n";

my @fileArray = ();
getFiles($targetDir, \@fileArray, $patternFile);
writeExcel($outFile,\@fileArray);
print "complate.\n";
exit;

# ファイルリスト取得
sub getFiles
{
	my ($path, $outArray, $patternFile) = @_;
	my $path_sjis = encode('cp932', $path);

	opendir(DIR, $path_sjis);
	my @files_sjis = readdir(DIR);
	closedir(DIR);
	
	foreach my $file_sjis (@files_sjis){
		my $file = decode('cp932', $file_sjis);
		next if $file =~ /^\.{1,2}$/; # '.','..'は読み飛ばし
		my $filePath = $path.'\\'.$file;
		my $filePath_sjis = encode('cp932', $filePath);
		if( -d $filePath_sjis ){
			getFiles( $filePath, $outArray, $patternFile);
			next;
		}
		if( $file =~ /^$patternFile$/ ){
			push @{$outArray}, $filePath;
			print "file[$filePath]\n";
		}
	}
}


sub getCSV
{
	my ($file,$values,$itemCount) = @_;
	print "getCSV file[$file]\n";
	my $file_sjis = encode('cp932',$file);
	open(IN ,"<$file_sjis") or die("error[$file_sjis]:$!");
	my @array = <IN> ;
	close IN;
	my $loop = $itemCount - 1;
	my $index = 0;
	my %keys = ();
	foreach my $line (@array){
		chomp $line;
		# パターンマッチの結果を配列アクセスで取得できる方法があったような気がするがわからん。。。
		if( $line =~ /^(?:($gCSVPattern),){$loop}($gCSVPattern)$/ ){
			$keys{$index} = $1;
# 			$values->[$index][0] = $1;
# 			$values->[$index][1] = $2;
			$values->[0][$index] = $1;
			$values->[1][$index] = $2;
#			print "key[$keys{$index}], 1[$values->[0][$index]], 2[$values->[1][$index]]\n";
		}
		else{
			die "getCSV::フォーマット異常[$file]\n";
		}
		$index++;
	}
	
	return (\%keys, $index);
}

sub isSameFormat
{
	my ($keys1, $keys2, $cnt1, $cnt2) = @_;

	if( $cnt1 != $cnt2){
		return 0; # 違うフォーマット
	}
	my $i = 0;
	for(; $i<$cnt1; $i++){
		if( !exists( $keys1->{$i}) ){
			last;
		}
		if( !exists( $keys2->{$i}) ){
			last;
		}
		if( $keys1->{$i} ne $keys2->{$i} ){
			last;
		}
	}
	if( $i != $cnt1 ){
		#print "i[$i\/$cnt1\:$cnt2] key1[$keys1->{$i}] key2[$keys2->{$i}]\n";
		return 0;
	}
	return 1;
}

sub setCell
{
	my ($worksheet, $format, $offset_x,$offset_y,$values) = @_;
	
	# $values->[$x][$y];
	my $width = @{$values};
	my $height = @{$values->[0]};

	print "[$width/$height]\n";

	for(my $y=0;$y<$height;$y++){
		for(my $x=0;$x<$width;$x++){
			$worksheet->write( $y + $offset_y, $x + $offset_x, $values->[$x][$y], $format);
#			$worksheet->write_comment( $y + $offset_y, $x + $offset_x, $values->[$x][$y], $format, $option);
		}
	}
}

# エクセルへ書込
# http://search.cpan.org/~jmcnamara/Spreadsheet-WriteExcel-2.38/lib/Spreadsheet/WriteExcel.pm
sub writeExcel
{
	my ($file,$files) = @_;

	print "writeExcel[$file]\n";
	my $file_sjis = encode('cp932', $file);
	# Create a new Excel workbook
	my $workbook = Spreadsheet::WriteExcel->new($file_sjis);
	# Add a worksheet
	my $worksheet = $workbook->add_worksheet();
	#  Add and define a format
	my $format = $workbook->add_format(); # Add a format
	# 一個目をベースとする
	my @baseValues = ();
	my ($baseKeys, $baseCnt) = getCSV( $files->[0], \@baseValues, 2);
	# それ以外
	for(my $y=1;$y<@{$files};$y++){
		my @tempValues = ();
		my ($keys, $cnt) = getCSV( $files->[$y], \@tempValues, 2);
		if( !isSameFormat($baseKeys, $keys, $baseCnt, $cnt) ){
			print "writeExcel::フォーマット異常[$files->[$y]]\n";
			next;
		}
		push @{$baseValues[$y+1]}, @{$tempValues[1]};
	}
 	$format->set_align('left');
 	$format->set_align('top');
	setCell($worksheet, $format, 0,1,\@baseValues);

	# 見出し
	my @titles = ();
	$titles[0][0] = "項目名／ファイル名";
	for(my $i=0; $i<@{$files}; $i++){
		$titles[1+$i][0] = $files->[$i];
	}
	
	my $format2 = $workbook->add_format(); # Add a format
	$format2->set_bg_color('silver');
 	$format2->set_bold();
 	$format2->set_align('left');
 	$format2->set_align('top');
	$format2->set_text_wrap();
	#	my $option = 'x_scale => 20, y_scale => 80';
	setCell($worksheet, $format2, 0,0, \@titles);
	$worksheet->set_column( 0, eval(@{$files}), 20);

	$workbook->close;
}
