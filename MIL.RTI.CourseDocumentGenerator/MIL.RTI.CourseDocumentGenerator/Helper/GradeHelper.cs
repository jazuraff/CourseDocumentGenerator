namespace MIL.RTI.CourseDocumentGenerator.Helper
{
    public static class GradeHelper
    {
        public static string ToRank(this string grade)
        {
            switch (grade.ToLower())
            {
                case "e1":
                    return "PVT";
                case "e2":
                    return "PV2";
                case "e3":
                    return "PFC";
                case "e4":
                    return "SPC";
                case "e5":
                    return "SGT";
                case "e6":
                    return "SSG";
                case "e7":
                    return "SFC";
                case "e8":
                    return "1SG";
                case "e9":
                    return "CSM";
                default:
                    return "";
            }
        }
    }
}
