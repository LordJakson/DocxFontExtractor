# DocxFontExtractor
Extracts embedded Fonts from a docx file
```
usage: DocxFontExtractor [-h] 
                         [-d [fonts]] 
                         [-l] 
                         [-n 'Font Name' ['Font Name' ...]]
                         [-i fontX [fontX ...]] [-v]
                         filename
  
Extracts embedded font from a docx file and converts it to usable TTF files

positional arguments:
  filename              DocxFile to extract the Fonts from.
  
options:
  -h, --help            show this help message and exit.
  -d [fonts], --TargetDir [fonts]
                        Directory to extract the Fonts to.
  -l, --List            List all the embedded Fonts instead of extracting them.
  -n 'Font Name' ['Font Name' ...], --FontNames 'Font Name' ['Font Name' ...]
                        List of font names to extract.
                        Regular expression can be used.
  -i fontX [fontX ...], --FontIds fontX [fontX ...]
                        List of font ids to extract.
  -v, --Verbose         Display logging information's.
                        -v  = INFO
                        -vv = DEBUG
```    
                        
