# Set-ExecutionPolicy RemoteSigned -Scope Process
Param(
    [String]$src = "src.pptx",
    [String]$dst = "dest.pptx",
    [Array]$words = {"word"}
)

function Extract-Slide {
    $srcPPTX    = $args[0]
    $dstPPTX    = $args[1]
    $words      = $args[2]
    $numSlide = $srcPPTX.Slides.Count

    # Uncomment code($dstPPTX_other) if you want to extract non-matching slides as well.
    # $dstPPTX_other = $powerPoint.Presentations.Add()

    # Slide Loop
    for ($i=1; $i -le $numSlide; $i++){
        $slide = $srcPPTX.Slides.Item($i)
        $found = $false

        # Words Loop
        foreach ($word in $words){
            if($slide.Shapes.Title.TextFrame.TextRange.Text.Contains($word) ){
                Write-Host "Found!!! Slide: $i" 
                $found = $true
                break
            }
        }

        # Extract slide
        if ($found){
            [void]$dstPPTX.Slides.InsertFromFile($srcPath, $dstPPTX.Slides.Count, $i, $i)
        }else{
            # [void]$dstPPTX_other.Slides.InsertFromFile($srcPath, $dstPPTX_other.Slides.Count, $i, $i)
        }
    } 

    $dstPPTX.Saveas($dstPath)
    # $dstPPTX_other.Saveas($dstPath + "_other")
}

function Main {
    try{
        $powerPoint = New-Object -ComObject powerpoint.application
        $powerPoint.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue

        $srcPPTX = $powerPoint.presentations.open($srcPath, $True, $False)
        $dstPPTX = $powerPoint.Presentations.Add()

        # Copy title(Slide 1) from srcPPTX
        [void]$dstPPTX.Slides.InsertFromFile($srcPath, 0, 1, 1)
    }
    catch [System.IO.FileNotFoundException]{
        Write-Output "Could not find $src or $dst"
        exit
    }
    catch [System.IO.IOException]{
        Write-Output "IO error with the file: $src or $dst"
        exit
    }

    Extract-Slide $srcPPTX $dstPPTX $words

    # Cleanup
    $srcPPTX.Close()
    $srcPPTX = $null
    $dstPPTX.Close()
    $dstPPTX = $null

    $powerPoint.Quit()
    $powerPoint = $null
    [GC]::Collect()
}


[string]$path = Split-Path $MyInvocation.MyCommand.Path -Parent
Set-Location $path
[Boolean]$first = $true

$srcPath = $path+"\"+$src
$dstPath = $path+"\"+$dst

Write-Host "SrcFile: $src SrcFilePath: $srcPath"
Write-Host "DstFile: $dst DstFilePath: $dstPath"
Write-Host "Word List: $words"

Main