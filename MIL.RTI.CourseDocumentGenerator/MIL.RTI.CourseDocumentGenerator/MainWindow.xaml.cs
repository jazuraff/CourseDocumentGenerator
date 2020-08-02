using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using MIL.RTI.IText.PdfManipulator;
using MIL.RTI.PdfDocuments.Constants;
using MIL.RTI.PdfDocuments.Models;
using MIL.RTI.PdfDocuments.Requests;

namespace MIL.RTI.CourseDocumentGenerator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private bool _handle = true;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void BtnBrowseSoliderData_Click(object sender, RoutedEventArgs e)
        {

        }

        private void BtnBrowseDestination_Click(object sender, RoutedEventArgs e)
        {

        }

        private void BtnGenerate_Click(object sender, RoutedEventArgs e)
        {
            var course = CboCourseSelection.Text;
            CourseCounselingRequest request;

            switch (course)
            {
                case "13m10 -- MOSQ":
                    request = BuildRequest();

                    //Things I want this app to do:
                    // - Take in excel spreadsheet of soldier data:
                    //    -- Name(last, first, MI), Rank/Grade, MOS
                    // - Input boxes for initial, midcourse, end of course counseling dates
                    // - Input for Name and Title of Counselor
                    // - Dropdown of possible choices for organization
                    // - Each Soldier should have an initial, midcourse, end of course counseling statement
                    // - Tab for each (initial, mid, end) file that allows edits to the current statements
                    string Source = "/myfiles/Counsel.pdf";
                    string Destination = "/myfiles/Counsel_Edit.pdf";

                    var manip = new Da4856Pdf(Source, Destination);

                    manip.GeneratePdf(request);
                    break;
            }
        }

        private CourseCounselingRequest BuildRequest()
        {
            var request = new CourseCounselingRequest
            {
                CounselorName = TxtCounselorName.Text,
                Destination = TxtDestination.Text,
                SoldierDataFileLocation = TxtSoldierData.Text,
                InitialCounseling = new CounselingData
                {
                    Assessment = TxtAssessmentInitial.Text,
                    DateOfCounseling = DtDateOfCounselingInitial.SelectedDate,
                    KeyPoints = TxtKeyPointsInitial.Text,
                    LeaderResponsibilities = TxtLeaderResponsibilitiesInitial.Text,
                    PlanOfAction = TxtPlanOfActionInitial.Text,
                    PurposeOfCounseling = TxtPurposeInitial.Text
                }
            };

            return request;
        }

        private void CboCourseSelection_DropdownClosed(object sender, System.EventArgs e)
        {
            if (_handle) Handle();
            _handle = true;
        }

        private void CboCourseSelection_SelectionChanged(object sender,
            System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (sender is ComboBox cmb) _handle = !cmb.IsDropDownOpen;
            Handle();
        }

        private void Handle()
        {
            ClearData();

            switch (CboCourseSelection.SelectedItem.ToString().Split(new[] {": "}, StringSplitOptions.None).Last())
            {
                case "13M10 -- MOSQ":
                    PopulateMosQData();
                    break;
                case "13M30 -- ALC":
                    PopulateAlcData();
                    break;
                case "13M40 -- SLC":
                    PopulateSlcData();
                    break;
            }
        }

        private void PopulateMosQData()
        {
            TxtPurposeInitial.Text = MosQualificationDefault.Purpose;
        }

        private void PopulateAlcData()
        {
            TxtPurposeInitial.Text = MosQualificationDefault.Purpose;
        }

        private void PopulateSlcData()
        {
            TxtPurposeInitial.Text = MosQualificationDefault.Purpose;
        }

        private void ClearData()
        {
            TxtPurposeInitial.Text = "";
            TxtAssessmentInitial.Text = "";
            TxtKeyPointsInitial.Text = "";
            TxtLeaderResponsibilitiesInitial.Text = "";
            TxtPlanOfActionInitial.Text = "";

            TxtPurposeMidCourse.Text = "";
            TxtAssessmentMidCourse.Text = "";
            TxtKeyPointsMidCourse.Text = "";
            TxtLeaderResponsibilitiesMidCourse.Text = "";
            TxtPlanOfActionMidCourse.Text = "";

            TxtPurposeEndCourse.Text = "";
            TxtAssessmentEndCourse.Text = "";
            TxtKeyPointsEndCourse.Text = "";
            TxtLeaderResponsibilitiesEndCourse.Text = "";
            TxtPlanOfActionEndCourse.Text = "";
        }
    }
}
