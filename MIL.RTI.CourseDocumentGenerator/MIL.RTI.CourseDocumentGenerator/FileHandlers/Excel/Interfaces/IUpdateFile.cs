using MIL.RTI.CourseDocumentGenerator.Requests;

namespace MIL.RTI.CourseDocumentGenerator.FileHandlers.Excel.Interfaces
{
    public interface IUpdateFile
    {
        void UpdateFile(CourseCounselingRequest request);
    }
}