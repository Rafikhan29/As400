set source=%~dp0
set today=%Date:~10,4%%Date:~4,2%%Date:~7,2%
set t=%time:~0,8%
set t=%t::=%
set t=%t: =0%
set timestamp=%today%_%t%
echo %timestamp%

set executebaleDir=D:\TenX\FutureGeneraliPOC
cd %executebaleDir%
D:
set captureScreenShot=True

echo ************Executing FG Testcases***********
call pybot --variable GlobalUserName:TENX01 --variable GlobalPassword:zxcvbm8i --logtitle FutureGenrali_FirePolicy_Data_Upload_Log --reporttitle FutureGenrali_DataUpload_Report --variable globalScreenShot:%captureScreenShot% --outputdir %executebaleDir%\Results\%timestamp% TestSuites\FirePolicy.txt
