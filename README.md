## ppt_translator:
A python script for translating the content in a .ppt or .pptx file

## Setup:
1) Run file.reg to enable drag-and-drop conversion.
2) Install necessary Python 3.8 dependencies (i.e. pptx, googletrans, pandas,).
3) Drag and drop .ppt or .pptx file on top of ppt_translator.py.
4) Specify target language.
5) Open translated document!

## To do:
preserve page numbers
access hidden text frames
give glossary lang labels
highlight translated content for dubious translations
if font is larger check if font of next run matches to concatenate
replace "& \n" and "&\n" with "& "

## Troubleshooting:
1) JSONDecodeError: Expecting value: line 1 column 1 (char 0)
Is usually caused because your IP address has been blocked by google.
Try changing your IP address with a vpn or wait awhile.
2) TypeError: 'NoneType' object is not iterable
Could be caused by extra spaces or weird characters within the powerpoint.
