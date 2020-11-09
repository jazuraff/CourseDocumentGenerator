namespace MIL.RTI.CourseDocumentGenerator.FileHandlers.File
{
    public static class FileHandler
    {
        public static void Copy(string sourceFileName, string sourcePath, string targetPath, string targetFileName = null)
        {
            var sourceFile = System.IO.Path.Combine(sourcePath, sourceFileName);
            var targetFile =
                System.IO.Path.Combine(targetPath, targetFileName ?? sourceFileName);

            System.IO.Directory.CreateDirectory(targetPath);

            System.IO.File.Copy(sourceFile, targetFile, true);
        }
    }
}