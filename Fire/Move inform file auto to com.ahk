#Persistent
SetTimer, CheckForNewFiles, 5000  ; Check every second
return

CheckForNewFiles:
    sourceFolder := "C:\Users\Peter\Downloads"
    destFolder := "D:\Committee\01 Inform"

    ; Create destination folder if it doesn't exist
    IfNotExist, %destFolder%
        FileCreateDir, %destFolder%

    ; Process all matching files in downloads
    filesToOpen := ""
    Loop, Files, %sourceFolder%\AUEED*.pdf
    {
        sourceFile := A_LoopFileFullPath

        ; Generate filename using current date/time
        FormatTime, currentDate,, yyyy-MM-dd
        FormatTime, currentTime,, HH mm ss
        baseName := currentDate . "_IScour_" . currentTime
        destFile := destFolder . "\" . baseName . ".pdf"

        ; Make filename unique if needed
        counter := 1
        while FileExist(destFile)
        {
            destFile := destFolder . "\" . baseName . "_" . counter . ".pdf"
            counter++
        }

        ; Try to move the file (skip if locked/in use)
        FileMove, %sourceFile%, %destFile%, 1
        if (ErrorLevel = 0)
        {
            filesToOpen .= destFile . "`n"
        }
    }

    ; Open all successfully moved files
    Loop, Parse, filesToOpen, `n
    {
        if (A_LoopField != "")
            Run, % A_LoopField
    }
return