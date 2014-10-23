Funtech.EpdmWordToPdf
=====================

Enterprise PDM Task add-in for saving Word documents to PDF

- Uses System.Reflection namespace to avoid requiring a reference to Microsoft.Office.Interop.Word assembly (i.e., one less assembly to deploy when the add-in is added to the vault).
- Targets .NET 3.5
- Microsoft Word must be installed on the host machine where the task ultimately runs and the host machine.
- Currently the task will simply save the PDF to the same location/name as the original doc/docx file(s) and takes no further action and therefore it is up to the user to check the files in manually after the task has run, etc.