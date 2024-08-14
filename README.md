# similar-file-organizer

Utility to organize all .doc, .docx, and .odt files in a provided directory and group them together according to the Levenshtein distance of their text contents. All documents that are below a specified maximum dissimilarity with each other are placed in their own directory. 

## Notes 
For .odt support on Linux, project must reference forked project https://github.com/wodeeken/AODL. If running on Windows, you can remove the project reference from the .csproj file and use the standard package distributed through Nuget.
