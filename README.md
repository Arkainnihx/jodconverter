# jodconverter
JODConverter automates document conversions using LibreOffice/OpenOffice.org
Hack to turn on line numbering call converter with boolean enableLineNumbering:

converter.convert(inputFile, outputFile, enableLineNumbering);

This is passed down to StandardConversionTask where LineNumberingProperties are set and Headers & Footers are disabled.
It would be nicer to just pass PropertySet to apply to document.
These changes are not saved. They are only applied to the document before conversion.
