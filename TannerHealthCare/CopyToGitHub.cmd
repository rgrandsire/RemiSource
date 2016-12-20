@echo off
Title Copying files real quick

echo Getting files from the Web first and copying to this Repo
cd "\Maintenance Connection\mc_iis\mapp_v12\modules\common"
copy *Tanner.asp "C:\Users\R3miG\OneDrive\Documents\MaintenanceConnection\Project\RemiSource\TannerHealthCare\"
cd ..\Asset
copy *Tanner*.* "C:\Users\R3miG\OneDrive\Documents\MaintenanceConnection\Project\RemiSource\TannerHealthCare\"
echo Are you sure you want to copy the files to GitHub?
Pause
cd "\Users\R3miG\OneDrive\Documents\MaintenanceConnection\Project\RemiSource\TannerHealthCare"
echo. 
echo First let's copy the SQL files...
copy *.sql "C:\Users\R3miG\Documents\GitHub\MaintenanceConnection\mcCoreMB\SQL Scripts\Other\TannerHealthCare\WO85594\"
copy *.txt "C:\Users\R3miG\Documents\GitHub\MaintenanceConnection\mcCoreMB\SQL Scripts\Other\TannerHealthCare\WO85594\"

echo.
echo Copy Tab file...
copy _mc_tab*_*.asp "C:\Users\R3miG\Documents\GitHub\MaintenanceConnection\mcCoreMB\mc_iis\mapp_v70\modules\common\"

echo.
echo Copy ASP, HTM files...
copy _asset_*.asp "C:\Users\R3miG\Documents\GitHub\MaintenanceConnection\mcCoreMB\mc_iis\mapp_v70\modules\asset\"
copy *.htm "C:\Users\R3miG\Documents\GitHub\MaintenanceConnection\mcCoreMB\mc_iis\mapp_v70\modules\asset\"


Echo Done
Pause
@exit

Remi G GRandsire 12/2016