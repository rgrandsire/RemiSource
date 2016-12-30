@echo off
Title Copying files real quick

echo Getting files from the Web first and copying to this Repo
cd "\Maintenance Connection\mc_iis\mapp_v12\modules\reports"
copy _*CKE.asp "C:\Users\R3miG\OneDrive\Documents\MaintenanceConnection\Project\RemiSource\CKE\WO88677\"
echo Are you sure you want to copy the files to GitHub?
Pause
cd "\Users\R3miG\OneDrive\Documents\MaintenanceConnection\Project\RemiSource\CKE\WO88677\"
echo. 
echo Copy ASP file...
copy _*CKE.asp "C:\Users\R3miG\Documents\GitHub\MaintenanceConnection\mcCoreMB\mc_iis\mapp_v70\modules\reports\"

Echo Done
Pause
@exit

Remi G GRandsire 12/2016