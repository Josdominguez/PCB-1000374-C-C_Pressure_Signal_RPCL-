@echo off

setlocal EnableDelayedExpansion
set pn1=1000374
set pn2=1000375
set variants=1
set rev=A-0

if exist %pn1%_Documents (
	echo Found %pn1%_Documents Folder !
	cd %pn1%_Documents
	if exist %pn1%-GERBERS (
		echo Cleaning %pn1%-GERBERS Folder
		cd %pn1%-GERBERS
		del *.zip
		del *.rul
		del *.APR_LIB
		del *.rep
		del *.ldp
		del *.apr
		del *.drr
		del "Status Report.txt"
		cd ..
	) ELSE (echo %pn1%-GERBERS Directory does not exist!)
	if exist %pn1%-ODB++_Files (
		echo Cleaning %pn1%-ODB++_Files Directory
		cd %pn1%-ODB++_Files
		del *.zip
		del *.rul
		del *.rep
		del *.tgz
		del "Status Report.txt"
		cd ..
	) ELSE (echo %pn1%-ODB++_Files Directory does not exist!)
	if exist %pn1%-STENCIL (
		echo Cleaning %pn1%-STENCIL Directory
		cd %pn1%-STENCIL
		del *.zip
		del *.rul
		del *.APR_LIB
		del *.rep
		del *.ldp
		del *.apr
		del *.drr
		del *.extrep
		del "Status Report.txt"
		cd ..
	) ELSE (echo %pn1%-STENCIL Directory does not exist!)
	echo Cleaning %pn1%_Documents Folder
	del "Status Report.txt"
	cd..
)
echo Please Confirm that the newest avl.xlsx is in the "Templates directory"
pause

FOR /L %%x IN (1,1,%variants%) DO (
    if exist %pn2%-%%x_Documents (
        echo Cleaning and Adding AVL Data to %pn2%-%%x_Documents
        cd %pn2%-%%x_Documents
        del "Status Report.txt"
        ..\..\..\PCB-Templates\BOM_AddAVL\get_avl %pn2%-%%x-%rev%_BOM.xlsx ..\..\..\Templates\AVL.xlsx %pn2%-%%x-%rev%_BOM_AVL.xlsx
        if	ERRORLEVEL 0 (
            del %pn2%-%%x-%rev%_BOM.xlsx
            rename %pn2%-%%x-%rev%_BOM_AVL.xlsx %pn2%-%%x-%rev%_BOM.xlsx
            cd ..
        ) ELSE (
            echo get_avl failed returned code %ERRORLEVEL%
            cd ..
            Goto :Error
        )
    ) ELSE (echo %pn2%-%%x_Documents Directory does not exist !)
)

:Error
	pause
