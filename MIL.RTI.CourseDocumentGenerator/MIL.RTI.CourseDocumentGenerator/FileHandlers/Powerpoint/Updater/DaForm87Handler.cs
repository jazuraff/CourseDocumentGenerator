using MIL.RTI.CourseDocumentGenerator.Constants;
using MIL.RTI.CourseDocumentGenerator.FileHandlers.Abstracts;
using MIL.RTI.CourseDocumentGenerator.FileHandlers.Interfaces;
using MIL.RTI.CourseDocumentGenerator.Requests;

namespace MIL.RTI.CourseDocumentGenerator.FileHandlers.Powerpoint.Updater
{
    class DaForm87Handler : BaseFileUpdater, IUpdateFile
    {
        private const string BaseFileName = "DA_Form_87_Certificates";

        public DaForm87Handler(string sourcePath, string targetPath, ClassType classType)
        : base(sourcePath, targetPath, BaseFileName, classType, FileTypes.Powerpoint) { }

        public void UpdateFile(CourseCounselingRequest request)
        {
            var fileName = $"{BaseTargetFileName}_Phase{request.Phase}{FileTypes.Powerpoint}";

            CopyFile(SourcePath, fileName, TargetPath);
        }
    }
}
