  use strict;
  use warnings;
  use utf8;
  use Encode;
  use Archive::Zip qw(:ERROR_CODES);

#
# .odsファイルのセルを取得する
#   Usage:
#     @cell = @{&readOds('ファイル名')}; 又は
#     @cell = @{&readOds(ファイルハンドル)};
#     結果は $cell[シート番号][行][列] でアクセス出来る

sub readOds {
  my ($file, $zip, $data, $ns, @used, @rv);
  $file = $_[0];

  if (!defined($file)) {
    print STDERR "Undefined filename or filehandle\n";
    exit 1;
  }
  elsif (ref($file) eq '') { # ファイル名の場合
    unless ($zip = Archive::Zip -> new($file)) {
      print STDERR 'Failed to read file ' . $file . "\n";
      exit 1;
    }
  }
  else { # ファイルハンドルの場合
    $zip = Archive::Zip -> new();
    if ($zip -> readFromFileHandle($file) != AZ_OK) {
      print STDERR "Failed to read filehandle\n";
      exit 1;
    }
  }

  $data = $zip -> contents('content.xml');
  @rv = ();
  if ($data) { $data = Encode::decode('utf8', $data); } else { return \@rv; }

  @used = ();
  $ns = 0;  # シート番号
  while ($data =~ /<table:table\s.*?>(.*?)<\/table:table>/sg) {
    my ($sht, $y);
    $sht = $1;
    $y = 0;  # 行
    while ($sht =~ /<table:table-row\s.*?>(.*?)<\/table:table-row>/sg) {
      my ($row, $x);
      $row = $1;
      $x = 0;  # 列
      while ($row =~ /(?|<table:table-cell([^>]*?)()\/>|<table:table-cell\s(.*?)>(.*?)<\/table:table-cell>)/sg) {
        my ($attr, $cell, $colspan, $rowspan, $repeat, $st, $flag, $str);
        $attr = $1;
        $cell = $2;
        if ($attr =~ /number-columns-spanned[\s="']*(\d+)/) {
          $colspan = $1;
        }
        else {
          $colspan = 1;
        }
        if ($attr =~ /number-rows-spanned[\s="']*(\d+)/) {
          $rowspan = $1;
        }
        else {
          $rowspan = 1;
        }
        if ($attr =~ /number-columns-repeated[\s="']*(\d+)/) {
          $repeat = $1;
        }
        else {
          $repeat = 1;
        }

        $st = '';
        while ($cell =~ /<text:p>(.*?)<\/text:p>/sg) {
          my ($lin);
          $lin = $1;
          if ($st eq '') { $st = $lin; } else { $st = $st . "\n" . $lin; }
        }
        $flag = 1;
        $str = '';
        while ($st =~ /<(.*?)>/) {
          my ($tag);
          $str .= $` if ($flag);
          $tag = lc($1);
          $st = $';
          if (index($tag, 'text:ruby-text') == 0) {
            $flag = 0;
          }
          else {
            $flag = 1;
          }
        }
        $str .= $st if ($flag);
        $str =~ s/&lt;/</g;
        $str =~ s/&gt;/>/g;
        $str =~ s/&quot;/"/g;
        $str =~ s/&apos;/'/g;
        $str =~ s/&amp;/&/g;

        while ($repeat > 0) {
          while (defined $used[$ns][$y][$x] && $used[$ns][$y][$x] > 0) {
            $x ++;
          }
          $rv[$ns][$y][$x] = $str;
          for (my $y0 = $y; $y0 < $y + $rowspan; $y0 ++) {
            for (my $x0 = $x; $x0 < $x + $colspan; $x0 ++) {
              $used[$ns][$y0][$x0] = 1;
            }
          }
          $x ++;
          $repeat --;
        }

      }
      $y ++;
    }
    $ns ++;
  }

  return \@rv;
}

1;
