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
            // - Dropdown of possible choices for organization
            using (new WaitCursor())
            {
                var request = BuildRequest();

                var errors = request.Validate();

                var excelSheet = new SoldierDataFile(TxtSoldierData.Text);

                if (errors.Count > 0)
                {
                    var formattedErrors = "";

                    errors.ForEach(er => { formattedErrors += $"- {er}\r\n"; });

                    MessageBox.Show(formattedErrors, "Please Enter Correct Data");

                    return;
                }

                try
                {
                    var soldierData = excelSheet.GetSoldierData();
                    request.SoldierData = soldierData;
                }
                catch (InvalidDataException ide)
                {
                    MessageBox.Show(ide.Message);
                    return;
                }

                var generator = new CourseFileGenerator(request);
                generator.Execute();
            }
        }

        private CourseCounselingRequest BuildRequest()
        { 

            var request = new CourseCounselingRequest
            {
                CounselorName = TxtCounselorName.Text,
                Destination = TxtDestination.Text,
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
            TxtKeyPointsInitial.Text = MosQualificationDefault.InitialKeyPoints;
            TxtPlanOfActionInitial.Text = MosQualificationDefault.InitialPlanOfAction;
            TxtLeaderResponsibilitiesInitial.Text = MosQualificationDefault.InitialLeaderResponsibilities;
            TxtAssessmentInitial.Text = MosQualificationDefault.InitialAssessment;

            TxtPurposeMidCourse.Text = MosQualificationDefault.MidCoursePurpose;
            TxtKeyPointsMidCourse.Text = MosQualificationDefault.MidCourseKeyPoints;
            TxtPlanOfActionMidCourse.Text = MosQualificationDefault.MidCoursePlanOfAction;
            TxtLeaderResponsibilitiesMidCourse.Text = MosQualificationDefault.MidCourseLeaderResponsibilities;
            TxtAssessmentMidCourse.Text = MosQualificationDefault.MidCourseAssessment;

            TxtPurposeEndCourse.Text = MosQualificationDefault.EndCoursePurpose;
            TxtKeyPointsEndCourse.Text = MosQualificationDefault.EndCourseKeyPoints;
            TxtPlanOfActionEndCourse.Text = MosQualificationDefault.EndCoursePlanOfAction;
            TxtLeaderResponsibilitiesEndCourse.Text = MosQualificationDefault.EndCourseLeaderResponsibilities;
            TxtAssessmentEndCourse.Text = MosQualificationDefault.EndCourseAssessment;
        }

        private void PopulateAlcData()
        {
            TxtPurposeInitial.Text = AlcDefault.InitialPurpose;
            TxtKeyPointsInitial.Text = AlcDefault.InitialKeyPoints;
            TxtPlanOfActionInitial.Text = AlcDefault.InitialPlanOfAction;
            TxtLeaderResponsibilitiesInitial.Text = AlcDefault.InitialLeaderResponsibilities;
            TxtAssessmentInitial.Text = AlcDefault.InitialAssessment;

            TxtPurposeMidCourse.Text = AlcDefault.MidCoursePurpose;
            TxtKeyPointsMidCourse.Text = AlcDefault.MidCourseKeyPoints;
            TxtPlanOfActionMidCourse.Text = AlcDefault.MidCoursePlanOfAction;
            TxtLeaderResponsibilitiesMidCourse.Text = AlcDefault.MidCourseLeaderResponsibilities;
            TxtAssessmentMidCourse.Text = AlcDefault.MidCourseAssessment;

            TxtPurposeEndCourse.Text = AlcDefault.EndCoursePurpose;
            TxtKeyPointsEndCourse.Text = AlcDefault.EndCourseKeyPoints;
            TxtPlanOfActionEndCourse.Text = AlcDefault.EndCoursePlanOfAction;
            TxtLeaderResponsibilitiesEndCourse.Text = AlcDefault.EndCourseLeaderResponsibilities;
            TxtAssessmentEndCourse.Text = AlcDefault.EndCourseAssessment;
        }

        private void PopulateSlcData()
        {
            TxtPurposeInitial.Text = SlcDefault.InitialPurpose;
            TxtKeyPointsInitial.Text = SlcDefault.InitialKeyPoints;
            TxtPlanOfActionInitial.Text = SlcDefault.InitialPlanOfAction;
            TxtLeaderResponsibilitiesInitial.Text = SlcDefault.InitialLeaderResponsibilities;
            TxtAssessmentInitial.Text = SlcDefault.InitialAssessment;

            TxtPurposeMidCourse.Text = SlcDefault.MidCoursePurpose;
            TxtKeyPointsMidCourse.Text = SlcDefault.MidCourseKeyPoints;
            TxtPlanOfActionMidCourse.Text = SlcDefault.MidCoursePlanOfAction;
            TxtLeaderResponsibilitiesMidCourse.Text = SlcDefault.MidCourseLeaderResponsibilities;
            TxtAssessmentMidCourse.Text = SlcDefault.MidCourseAssessment;

            TxtPurposeEndCourse.Text = SlcDefault.EndCoursePurpose;
            TxtKeyPointsEndCourse.Text = SlcDefault.EndCourseKeyPoints;
            TxtPlanOfActionEndCourse.Text = SlcDefault.EndCoursePlanOfAction;
            TxtLeaderResponsibilitiesEndCourse.Text = SlcDefault.EndCourseLeaderResponsibilities;
            TxtAssessmentEndCourse.Text = SlcDefault.EndCourseAssessment;
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