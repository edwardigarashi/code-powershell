
## get-date -uform "%Y%m%d%A%Z"
$date = Get-date -uform "%Y%m%d"
$date

# robocopy D:\ B:\lp1\ /MIR /FFT /R:3 /W:10 /Z /LOG:B:\lp1\log\bk_$date.log
robocopy D:\ B:\lp1\ /MIR /FFT /R:3 /W:10 /Z 
