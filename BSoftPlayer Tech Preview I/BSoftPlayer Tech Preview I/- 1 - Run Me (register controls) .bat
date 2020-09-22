@Echo Off
echo BSoftPlayer Control Setup Utility
echo ===========================================================
echo "The following utility will copy all required OCX controls to their proper location and register them.
echo ===========================================================
echo Press Control+C to exit, or
pause
cls
echo BSoftPlayer Control Setup Utility
echo ===========================================================
echo Copying controls (1 of 2)...
@copy ccrpftv6.ocx C:\
cls
echo BSoftPlayer Control Setup Utility
echo ===========================================================
echo Copying controls (2 of 2)...
@copy RICHTX32.ocx C:\
cls
echo BSoftPlayer Control Setup Utility
echo ===========================================================
echo Controls successfully copied.
echo Ready to register controls. Be sure to press "OK" in the next two dialog boxes.
pause
cls
echo BSoftPlayer Control Setup Utility
echo ===========================================================
echo Registering controls (1 of 2)...
regsvr32 C:\ccrpftv6.ocx
cls
echo BSoftPlayer Control Setup Utility
echo ===========================================================
echo Registering controls (2 of 2)...
regsvr32 C:\Richtx32.ocx
cls
echo BSoftPlayer Control Setup Utility
echo ===========================================================
echo Controls registered.
pause
