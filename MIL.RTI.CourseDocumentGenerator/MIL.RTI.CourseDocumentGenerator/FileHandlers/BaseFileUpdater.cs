using MIL.RTI.CourseDocumentGenerator.Constants;
using MIL.RTI.CourseDocumentGenerator.FileHandlers.File;

namespace MIL.RTI.CourseDocumentGenerator.FileHandlers
{
    public abstract class BaseFileUpdater
    {
        protected readonly string TargetPath;
        protected readonly string BaseTargetFileName;
        protected readonly string SourcePath;
        protected readonly int PhaseCount;

        protected BaseFileUpdater(string sourcePath, string targetPath, string baseFileName, ClassType classType)
        {
            string newFileName;

            switch (classType)
            {
                case ClassType.Mosq:
                    newFileName = $"{baseFileName}_13M10";
                    PhaseCount = 1;
                    break;
                case ClassType.Alc:
                    newFileName = $"{baseFileName}_13M30";
                    PhaseCount = 2;
                    break;
                default:
                    newFileName = $"{baseFileName}_13M40";
                    PhaseCount = 2;
                    break;
            }

            SourcePath = sourcePath;
            TargetPath = targetPath;
            BaseTargetFileName = newFileName;
        }

        protected string GetFileName(int phase)
        {
            var fileName = $"{BaseTargetFileName}_Phase{phase}.xlsx";
            return fileName;
        }

        protected string GetFullPath(int phase)
        {
            var fullPath = System.IO.Path.Combine(TargetPath, GetFileName(phase));

            return fullPath;
        }

        protected void CopyFile(string sourcePath, string sourceFileName, string targetPath, string targetFileName = null)
        {
            FileHandler.Copy(sourceFileName, sourcePath, targetPath, targetFileName);
        }
    }
}