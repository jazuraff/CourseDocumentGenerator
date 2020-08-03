using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using MIL.RTI.CourseDocumentGenerator.Constants.CourseDefaults;
using MIL.RTI.CourseDocumentGenerator.FileHandlers;
using MIL.RTI.CourseDocumentGenerator.FileHandlers.Excel;
using MIL.RTI.CourseDocumentGenerator.Helper;
using MIL.RTI.CourseDocumentGenerator.Models;
using MIL.RTI.CourseDocumentGenerator.Requests;
using ComboBox = System.Windows.Controls.ComboBox;
using MessageBox = System.Windows.MessageBox;

namespace MIL.RTI.CourseDocumentGenerator
{
    /// <summary>
    ///     Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private bool _handle = true;
        private const string Source = "./Files/Da4856July2014.pdf";

        public MainWindow()
        {
            InitializeComponent();
        }

        private void BtnBrowseSoliderData_Click(object sender, RoutedEventArgs e)
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Filter = @"Excel files (*.xls or .xlsx)|.xls;*.xlsx";
                dlg.ValidateNames = true;
                var result = dlg.ShowDialog();

                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    TxtSoldierData.Text = dlg.FileName;
                }
            }
        }

        private void BtnBrowseDestination_Click(object sender, RoutedEventArgs e)
        {
            using (var dlg = new FolderBrowserDialog())
            {
                dlg.Description = @"Select the directory you want your files to be in";
                dlg.ShowNewFolderButton = true;

                var result = dlg.ShowDialog();

                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    TxtDestination.Text = dlg.SelectedPath;
                }
            }
        }

        private void BtnGenerate_Click(object sender, RoutedEventArgs e)
        {
            //Things I want this app to do:
            // - Take in excel spreadsheet of soldier data:
            //    -- Name(last, first, MI), Rank/Grade, MOS
            // - Input boxes for initial, midcourse, end of course counseling dates
            // - Input for Name and Title of Counselor
            // - Dropdown of possible choices for organization
            // - Each Soldier should have an initial, midcourse, end of course counseling statement
            // - Tab for each (initial, mid, end) file that allows edits to the current statements
            using (new WaitCursor())
            {
                var excelSheet = new SoldierDataFile(TxtSoldierData.Text);
                List<SoldierData> soldierData;

                try
                {
                    soldierData = excelSheet.GetSoldierData();
                }
                catch (InvalidDataException ide)
                {
                    MessageBox.Show(ide.Message);
                    return;
                }

                var request = BuildRequest(soldierData);

                var errors = request.Validate();

                if (errors.Count > 0)
                {
                    var formattedErrors = "";

                    errors.ForEach(er => { formattedErrors += $"- {er}\r\n"; });

                    MessageBox.Show(formattedErrors, "Please Enter Correct Data");
                }

                var generator = new CourseFileGenerator(request);
                generator.Execute();
            }
        }

        private CourseCounselingRequest BuildRequest(List<SoldierData> soldierData)
        {
            var request = new CourseCounselingRequest
            {
                CounselorName = TxtCounselorName.Text,
                Destination = TxtDestination.Text,
                SoldierData = soldierData,
                InitialCounseling = new CounselingData
                {
                    Assessment = TxtAssessmentInitial.Text,
                    DateOfCounseling = DtDateOfCounselingInitial.SelectedDate,
                    KeyPoints = TxtKeyPointsInitial.Text,
                    LeaderResponsibilities = TxtLeaderResponsibilitiesInitial.Text,
                    PlanOfAction = TxtPlanOfActionInitial.Text,
                    PurposeOfCounseling = TxtPurposeInitial.Text
                },
                MidCourseCounseling = new CounselingData
                {
                    Assessment = TxtAssessmentMidCourse.Text,
                    DateOfCounseling = DtDateOfCounselingMidCourse.SelectedDate,
                    KeyPoints = TxtKeyPointsMidCourse.Text,
                    LeaderResponsibilities = TxtLeaderResponsibilitiesMidCourse.Text,
                    PlanOfAction = TxtPlanOfActionMidCourse.Text,
                    PurposeOfCounseling = TxtPurposeMidCourse.Text
                },
                EndOfCourseCounseling = new CounselingData
                {
                    Assessment = TxtAssessmentEndCourse.Text,
                    DateOfCounseling = DtDateOfCounselingEnd.SelectedDate,
                    KeyPoints = TxtKeyPointsEndCourse.Text,
                    LeaderResponsibilities = TxtLeaderResponsibilitiesEndCourse.Text,
                    PlanOfAction = TxtPlanOfActionEndCourse.Text,
                    PurposeOfCounseling = TxtPurposeEndCourse.Text
                }
            };

            return request;
        }

        private void CboCourseSelection_DropdownClosed(object sender, EventArgs e)
        {
            if (_handle) Handle();
            _handle = true;
        }

        private void CboCourseSelection_SelectionChanged(object sender,
            SelectionChangedEventArgs e)
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
            TxtPurposeInitial.Text = MosQualificationDefault.InitialPurpose;
            
            TxtPurposeMidCourse.Text = MosQualificationDefault.MidCoursePurpose;
            
            TxtPurposeEndCourse.Text = MosQualificationDefault.EndCoursePurpose;
        }

        private void PopulateAlcData()
        {

        }

        private void PopulateSlcData()
        {

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