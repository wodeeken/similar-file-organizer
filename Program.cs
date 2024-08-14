using Fastenshtein;
using System.IO;
using System.Text;
using AODL.Document.TextDocuments;
using AODL.Document.Content;
using System.Xml.Linq;
using NPOI.POIFS.FileSystem;
using NPOI.HWPF;
using NPOI.HWPF.UserModel;
using NPOI.HPSF;
using NPOI.XWPF.UserModel;
using System.Linq.Expressions;
using System.Globalization;
using System.ComponentModel.DataAnnotations;
using System.Reflection;
using NPOI.SS.Formula.Functions;
namespace similar_file_organizer;

class Program
{
    static void WriteHeader(){
        Console.WriteLine("");
        Console.WriteLine($"V{Assembly.GetEntryAssembly().GetName().Version}");
    }
    static List<string> persistentLines = new List<string>();
    // All lines written here will persist in console for the lifetime of the application.
    static void WritePersistent(string lineToWrite){
        persistentLines.Add(lineToWrite);
        Console.WriteLine(lineToWrite);
    }
    // All lines written here will persist only until the console is cleared next.
    static void WriteTemporary(string lineToWrite){
        Console.WriteLine(lineToWrite);
    }
    // Clear console of all temporary lines.
    static void RewritePersistentConsole(){
        Console.Clear();
        foreach(var line in persistentLines)
            Console.WriteLine(line);
    }
    static void Main(string[] args)
    {
        string filepath = "";
        double maxDistance = 0;
        WritePersistent("^^Similar File Organizer^^");
        WritePersistent($"V{Assembly.GetEntryAssembly().GetName().Version}");
        WritePersistent($"Last Updated: June 15, 2024");
        if(args.Count() == 1){
            if(args[0] == "-h"){
                Console.WriteLine("This program interates through all .doc, .docx, and .odt files in the provided directory and finds the levenshtein distance of the text of each pair.");
                Console.WriteLine("Each distance is normalized by dividing by the maximum of the two's text length to a number between 0 and 1.");
                Console.WriteLine("All documents with distances below the maximum dissimilarity are organized into their own folders.");
                Console.WriteLine("Usage: similar_file_organizer path/to/directory [0 - 1]");
                Console.WriteLine("path/to/directory - the directory to organize by similarity");
                Console.WriteLine("[0 - 1]: a decimal between 0 and 1 specifying the maximum dissimilarity of similar files (Optional, default is .2).");
                return;
            }else{
                if(Directory.Exists(args[0])){
                    filepath = args[0];
                    maxDistance = .2;
                }else{
                    Console.WriteLine($"Provided directory path {args[0]} does not exist.");
                    return;
                }
            }
        }else if(args.Count() == 2){
            if(Directory.Exists(args[0])){
                filepath = args[0];
                if(!Double.TryParse(args[1], out maxDistance)){
                    Console.WriteLine($"Provided Max Dissimilarity {args[1]} is invalid.");
                    return;
                }
            }else{
                Console.WriteLine($"Provided directory path {args[0]} does not exist.");
                return;
            }
        }else{
            Console.WriteLine("Incorrect number of parameters specified. Use similar_file_organizer -h for usage instructions.");
            return;
        }
        
        Dictionary<Tuple<string,string>,float> filePair = new Dictionary<Tuple<string, string>, float>();
        Dictionary<string,string> fileText = new Dictionary<string, string>();
        bool usedStep0 = false;
        // First arg - file path. Second Arg - max distance (0-1) for similar files.
        //string filepath = "/home/will/Desktop/TestingDocs/" ;//"/mnt/shares/Documents";
        //double maxDistance = 0.2;
        WritePersistent($"Running comparison in directory {filepath} with maximum dissimilarity of {maxDistance}");
        string[] directory = Directory.GetFiles(filepath);
        WritePersistent($"Files to process: {directory.Count()}");
        WriteTemporary("Determining size of directory...");
        // Step 0 - If the size of the directory is 1 GB or less, store document text in-memory for effeciency.
        long size = DirSize(new DirectoryInfo(filepath));
        if(size <= 1000000000){
            WritePersistent($"Directory {filepath} has size ~{size/1000000000} gb, will read in text and store in-memory for efficiency.");
            usedStep0 = true;
            // Step 0 - read in text from all documents and store in-memory.
            int currentPairCount1 = 0;
            DateTime beginTime1 = DateTime.Now;
            for(int i = 0;i < directory.Count();i++){
                RewritePersistentConsole();
                WriteTemporary($"Step 0 - read all document text: {(((double)i / (double) directory.Count()) * 100).ToString("0")}% - ({DateTime.Now - beginTime1})");
                if(ReadDocumentText(directory[i], out string fileText1)){
                    fileText.Add(directory[i], fileText1);
                }else{
                    WritePersistent($"Could not read in file {directory[i]}. Skipping similarity calculation.");
                }
                
                
            }
            RewritePersistentConsole();
            WritePersistent($"Step 0 Completed - {DateTime.Now - beginTime1}");
        }else{
            WritePersistent($"Directory {filepath} has size ~{size/1000000000} gb, too large to cache text in memory. Skipping step 0.");
        }
        // Total number of pairs.
        int currentPairCount = 0;
        DateTime beginTime = DateTime.Now;
        // Step 1 - calculate distances between file contents.
        for(int i = 0;i < directory.Count();i++){
            RewritePersistentConsole();
            WriteTemporary($"Step 1 - similarity calculations: {(((double)i /(double) directory.Count()) * 100).ToString("0")}% - ({DateTime.Now - beginTime})");
            // Get text from doc depending on extension.
            int lastIndex1 = directory[i].LastIndexOf(".");
            string ext1 = directory[i].Substring(lastIndex1);
            string file1 = ""; 
            if(!usedStep0){
                if(!ReadDocumentText(directory[i], out file1)){
                    currentPairCount += directory.Count() - i;
                    RewritePersistentConsole();
                    WritePersistent($"Could not read in file {directory[i]}. Skipping similarity calculation.");
                    continue;
                }
            }else{
                if(fileText.ContainsKey(directory[i]))
                    file1 = fileText[directory[i]];
                else
                    continue;
            }
            if(file1 == ""){
                currentPairCount += directory.Count() - i;
                RewritePersistentConsole();
                WritePersistent($"File {directory[i]} has no text. Skipping similarity calculation.");
                continue;
            }
                
            Fastenshtein.Levenshtein lev = new Fastenshtein.Levenshtein(file1);
        
            for(int j = i + 1; j < directory.Count();j++){
                RewritePersistentConsole();
                WriteTemporary($"Step 1 - similarity calculations: {(((double)i / (double) directory.Count()) * 100).ToString("0")}% - ({DateTime.Now - beginTime})");
                int lastIndex2 = directory[j].LastIndexOf(".");
                string ext2 = directory[j].Substring(lastIndex2);
                string file2 = ""; 
                if(!usedStep0){
                    if(!ReadDocumentText(directory[j], out file2)){
                        currentPairCount++;
                        RewritePersistentConsole();
                        WritePersistent($"Could not read in file {directory[j]}. Skipping similarity calculation.");
                        continue;
                    }
                }else{
                    if(fileText.ContainsKey(directory[j]))
                        file2 = fileText[directory[j]];
                    else
                        continue;
                }
                if(file2 == ""){
                    currentPairCount++;
                    RewritePersistentConsole();
                    WritePersistent($"File {directory[j]} has no text. Skipping similarity calculation.");
                    continue;
                }
                int distance = lev.DistanceFrom(file2);
                System.Diagnostics.Debug.WriteLine($"Distance: {distance}, {directory[i]} Length: {file1.Length}, {directory[j]}  Length: {file2.Length}");
                filePair.Add(new Tuple<string,string>(directory[i],directory[j]), (float)  lev.DistanceFrom(file2) / (Math.Max(file1.Length,file2.Length)));
                currentPairCount++;
            }
        }
        RewritePersistentConsole();
        WritePersistent($"Step 1 Completed - {DateTime.Now - beginTime}");
        // Step 2 - organize similar files into folders.
        // Delimiter
        string delimiter = "";
        if(Environment.OSVersion.Platform == PlatformID.Unix){
            delimiter = "/";
        }else if(Environment.OSVersion.Platform == PlatformID.Win32NT){
            delimiter = "\\";
        }
        // first, find all files that are within specified distance of each other.
        var similarFiles = filePair.Where(x => x.Value <= maxDistance ).ToList();
        int initialCount = similarFiles.Count();
        DateTime step2BeginTime = DateTime.Now;
        while(similarFiles.Count > 0){
            RewritePersistentConsole();
            WriteTemporary($"Step 2 - similar file move: {((1 - ((double)similarFiles.Count / (double) initialCount )) * 100).ToString("0")}% - ({DateTime.Now - step2BeginTime})");
            var currentSimilarFile = similarFiles.ElementAt(0);
            // For each simliar pair, find all other pairs with the same files.
            var similar_File1_1 = similarFiles.Where(x => x.Key != currentSimilarFile.Key &&
                                                             currentSimilarFile.Key.Item1 == x.Key.Item1).ToList();
            var similar_File1_2 = similarFiles.Where(x => x.Key != currentSimilarFile.Key &&
                                                             currentSimilarFile.Key.Item1 == x.Key.Item2).ToList();
            var similar_File2_1 = similarFiles.Where(x => x.Key != currentSimilarFile.Key &&
                                                             currentSimilarFile.Key.Item2 == x.Key.Item1).ToList();
            var similar_File2_2 = similarFiles.Where(x => x.Key != currentSimilarFile.Key &&
                                                             currentSimilarFile.Key.Item2 == x.Key.Item2).ToList();
            // Create a folder for this set of similar files.
            int indexOfNameBegin = currentSimilarFile.Key.Item1.LastIndexOf(delimiter);
            int indexOfExtension = currentSimilarFile.Key.Item1.Substring(indexOfNameBegin+ 1).LastIndexOf(".");
            string folderName = currentSimilarFile.Key.Item1.Substring( indexOfNameBegin + 1);
            folderName = folderName.Substring(0, indexOfExtension);
            bool addedFile = false;
            DirectoryInfo dr = Directory.CreateDirectory(Path.Combine(filepath, folderName));
            while(similar_File1_1.Count() > 0){
                var simFile = similar_File1_1.ElementAt(0);
                //  Get name.
                int indexOfNameBegin1 = simFile.Key.Item2.LastIndexOf(delimiter);
                if(File.Exists(simFile.Key.Item2)){
                    File.Move(simFile.Key.Item2, Path.Combine(dr.FullName,simFile.Key.Item2.Substring(indexOfNameBegin1 + 1)));
                    addedFile = true;
                }
                similarFiles.Remove(simFile);
                similar_File1_1.Remove(simFile);
                similar_File1_2.Remove(simFile);
                similar_File2_1.Remove(simFile);
                similar_File2_2.Remove(simFile);
            }
            
            while(similar_File1_2.Count() > 0){
                var simFile = similar_File1_2.ElementAt(0);
                int indexOfNameBegin1 = simFile.Key.Item1.LastIndexOf(delimiter);
                if(File.Exists(simFile.Key.Item1)){
                    File.Move(simFile.Key.Item1, Path.Combine(dr.FullName,simFile.Key.Item1.Substring(indexOfNameBegin1 + 1)));
                    addedFile = true;
                }
                similarFiles.Remove(simFile);
                similar_File1_1.Remove(simFile);
                similar_File1_2.Remove(simFile);
                similar_File2_1.Remove(simFile);
                similar_File2_2.Remove(simFile);
            }
            while(similar_File2_1.Count() > 0){
                var simFile = similar_File2_1.ElementAt(0);
                int indexOfNameBegin1 = simFile.Key.Item2.LastIndexOf(delimiter);
                if(File.Exists( simFile.Key.Item2)){
                    File.Move(simFile.Key.Item2, Path.Combine(dr.FullName,simFile.Key.Item2.Substring(indexOfNameBegin1 + 1)));
                    addedFile = true;
                }
                similarFiles.Remove(simFile);
                similar_File1_1.Remove(simFile);
                similar_File1_2.Remove(simFile);
                similar_File2_1.Remove(simFile);
                similar_File2_2.Remove(simFile);
            }
            while(similar_File2_2.Count() > 0){
                var simFile = similar_File2_2.ElementAt(0);
                int indexOfNameBegin1 = simFile.Key.Item1.LastIndexOf(delimiter);
                if(File.Exists(simFile.Key.Item1)){
                    File.Move(simFile.Key.Item1, Path.Combine(dr.FullName,simFile.Key.Item1.Substring(indexOfNameBegin1 + 1)));
                    addedFile = true;
                }
                similarFiles.Remove(simFile);
                similar_File1_1.Remove(simFile);
                similar_File1_2.Remove(simFile);
                similar_File2_1.Remove(simFile);
                similar_File2_2.Remove(simFile);
            }
            if(File.Exists(currentSimilarFile.Key.Item1)){
                int indexOfNameBegin3 = similarFiles.ElementAt(0).Key.Item1.LastIndexOf(delimiter);
                File.Move(currentSimilarFile.Key.Item1, Path.Combine(dr.FullName, currentSimilarFile.Key.Item1.Substring(indexOfNameBegin3 + 1)));
                addedFile = true;
            }
            if(File.Exists(currentSimilarFile.Key.Item2)){
                int indexOfNameBegin3 = similarFiles.ElementAt(0).Key.Item2.LastIndexOf(delimiter);
                File.Move(currentSimilarFile.Key.Item2, Path.Combine(dr.FullName, currentSimilarFile.Key.Item2.Substring(indexOfNameBegin3 + 1)));
                addedFile = true;
            }

            similarFiles.Remove(currentSimilarFile);
            // Remove folder if no files were ever added to it.
            if(!addedFile){
                Directory.Delete(dr.FullName);
            }
        }
        RewritePersistentConsole();
        WritePersistent($"Step 2 Completed - {DateTime.Now - step2BeginTime}");
        WritePersistent("Quitting");

    }
    public static List<string> SimilarFiles_Helper(List<KeyValuePair<Tuple<string,string>, float>> similarFiles, KeyValuePair<Tuple<string,string>,float> currentItem){
        List<string> listToReturn = new List<string>();
        List<KeyValuePair<Tuple<string,string>, float>> file1_1Similar = similarFiles.Where(x => (x.Key.Item1 == currentItem.Key.Item1 ) && x.Key != currentItem.Key).ToList();
        List<KeyValuePair<Tuple<string,string>, float>> file1_2Similar = similarFiles.Where(x => (x.Key.Item1 == currentItem.Key.Item2 ) && x.Key != currentItem.Key).ToList();
        List<KeyValuePair<Tuple<string,string>, float>> file2_1Similar = similarFiles.Where(x => (x.Key.Item2 == currentItem.Key.Item1) && x.Key != currentItem.Key).ToList();
        List<KeyValuePair<Tuple<string,string>, float>> file2_2Similar = similarFiles.Where(x => (x.Key.Item2 == currentItem.Key.Item2) && x.Key != currentItem.Key).ToList();
        for(int i = 0; i < file1_1Similar.Count; i++){
            KeyValuePair<Tuple<string,string>,float> curPair = file1_1Similar.ElementAt(i);
            listToReturn.Add(curPair.Key.Item2);
            listToReturn.Concat(SimilarFiles_Helper(similarFiles, curPair));
        }
        for(int i = 0; i < file1_2Similar.Count; i++){
            KeyValuePair<Tuple<string,string>,float> curPair = file1_2Similar.ElementAt(i);
            listToReturn.Add(curPair.Key.Item2);
            listToReturn.Concat(SimilarFiles_Helper(similarFiles, curPair));
        }
        for(int i = 0; i < file2_1Similar.Count; i++){
            KeyValuePair<Tuple<string,string>,float> curPair = file2_1Similar.ElementAt(i);
            listToReturn.Add(curPair.Key.Item2);
            listToReturn.Concat(SimilarFiles_Helper(similarFiles, curPair));
        }
        for(int i = 0; i < file2_2Similar.Count; i++){
            KeyValuePair<Tuple<string,string>,float> curPair = file2_2Similar.ElementAt(i);
            listToReturn.Add(curPair.Key.Item2);
            listToReturn.Concat(SimilarFiles_Helper(similarFiles, curPair));
        }
        
        return listToReturn;
    }
    public static bool GetTextFromOdt(string path, out string fileText)
    {
        try{
            var sb = new StringBuilder();
            using (var doc = new TextDocument())
            {
                doc.Load(path);

                //The header and footer are in the DocumentStyles part. Grab the XML of this part
                XElement stylesPart = XElement.Parse(doc.DocumentStyles.Styles.OuterXml);
                //Take all headers and footers text, concatenated with return carriage
                string stylesText = string.Join("\r\n", stylesPart.Descendants().Where(x => x.Name.LocalName == "header" || x.Name.LocalName == "footer").Select(y => y.Value));

                //Main content
                var mainPart = doc.Content.Cast<IContent>();
                var mainText = String.Join("\r\n", mainPart.Select(x => x.Node.InnerText));

                //Append both text variables
                sb.Append(stylesText + "\r\n");
                sb.Append(mainText);
                fileText = sb.ToString();
                return true;
            }

        }catch(Exception e){
            fileText = "";
            return false;
        }

    }
    public static bool GetTextFromDoc(string path, out string fileText){
        System.Text.EncodingProvider ppp = System.Text.CodePagesEncodingProvider.Instance;
        Encoding.RegisterProvider(ppp);
        try{
            NPOI.POIFS.FileSystem.POIFSFileSystem fs = new NPOI.POIFS.FileSystem.POIFSFileSystem(File.OpenRead(path));
            HWPFDocument doc = new HWPFDocument(fs);
            NPOI.HWPF.UserModel.Range docRange = doc.GetRange();
            fileText = docRange.Text;
            return true;
        }catch(Exception e){
            fileText = "";
           return false;
        }
    }
    public static bool GetTextFromDocx(string path, out string fileText){
        try{
            string returnString = "";
            using(FileStream fileStream = File.OpenRead(path)){
                XWPFDocument doc = new XWPFDocument(fileStream);
                foreach(var paragraph in doc.Paragraphs){
                    returnString += paragraph.ParagraphText;
                }
            }
            fileText = returnString;
            return true;
        }catch(Exception e){
            fileText = "";
            return false;
        }
    }
    public static long DirSize(DirectoryInfo d) 
    {    
        long size = 0;    
        // Add file sizes.
        FileInfo[] fis = d.GetFiles();
        foreach (FileInfo fi in fis) 
        {      
            size += fi.Length;    
        }
        // Add subdirectory sizes.
        DirectoryInfo[] dis = d.GetDirectories();
        foreach (DirectoryInfo di in dis) 
        {
            size += DirSize(di);   
        }
        return size;  
    }
    public static bool ReadDocumentText(string path, out string fileText){
        int lastIndex = path.LastIndexOf(".");
        string ext = path.Substring(lastIndex);
        switch(ext){
            case ".doc":
                return GetTextFromDoc(path, out fileText);
                
            case ".docx":
                return GetTextFromDocx(path, out fileText);
            case ".odt":
                return GetTextFromOdt(path, out fileText);
            case ".txt":
                try{
                    fileText = File.ReadAllText(path);
                    return true;
                }catch(Exception e){
                    fileText = "";
                    return false;
                }
            default:
                fileText = "";
                return false;
            
        }
    }
    
   
}
