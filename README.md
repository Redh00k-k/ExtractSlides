# ExtractSlides
PowerShell script to extract pptx slides with titles containing specified words.

# Usage
```
.\ExtractSlides.ps1 <Source> <Destination> <Words>
```

# Example
```powershell
> .\ExtractSlides.ps1 src.pptx dst.pptx B,C
SrcFile: .\src.pptx SrcFilePath: \path\to\src.pptx
DstFile: .\dst.pptx DstFilePath: \path\to\dst.pptx
Word List: B C
Found!!! Slide: 4
Found!!! Slide: 7
Found!!! Slide: 8

> dir
-a----        2023/07/17     15:50          37730 dst.pptx
-a----        2023/07/17     15:49           2416 ExtractSlides.ps1
-a----        2023/07/17     15:48          42757 src.pptx
```