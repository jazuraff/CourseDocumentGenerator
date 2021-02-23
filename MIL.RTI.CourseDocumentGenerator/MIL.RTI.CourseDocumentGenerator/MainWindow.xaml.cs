using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using MIL.RTI.CourseDocumentGenerator.Constants;
using MIL.RTI.CourseDocumentGenerator.Constants.CourseDefaults;
using MIL.RTI.CourseDocumentGenerator.FileHandlers;
using MIL.RTI.CourseDocumentGenerator.FileHandlers.Excel;
using MIL.RTI.CourseDocumentGenerator.Helper;
using MIL.RTI.CourseDocumentGenerator.Models;
using MIL.RTI.CourseDocumentGenerator.Requests;
using MessageBox = System.Windows.MessageBox;

namespace MIL.RTI.CourseDocumentGenerator
{
    /// <summary>
    ///     Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //TODO: the following processes have not yet been built
        //1. Certificates - powerpoint - one slide per soldier - Form 87
            // - Rank, First Name, Last Name, date range, CourseManager, Instructor
        private bool _handle = true;
        private ClassType _class;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void BtnBrowseSoliderData_Click(object sender, RoutedEventArgs e)
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Filter = @"Excel files (*.xls or .xlsx)|*.xls;*.xlsx";
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
            using (new WaitCursor())
            {
                var request = BuildRequest();

                var errors = request.Validate();

                var excelSheet = new SoldierDataFileHandler(TxtSoldierData.Text);

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
                finally
                {
                    excelSheet.Dispose();
                }

                var generator = new CourseFileGenerator(request);
                generator.Execute();
            }
        }

        private CourseCounselingRequest BuildRequest()
        { 

            var request = new CourseCounselingRequest
            {
                InstructorName = TxtInstructorName.Text,
                InstructorTitle = TxtInstructorTitle.Text,
                Destination = TxtDestination.Text,
                CourseStartDate = DtStartDate.SelectedDate,
                CourseEndDate = DtEndDate.SelectedDate,
                ClassNumber = txtClassNumber.Text,
                FiscalYear = txtFiscalYear.Text,
                Phase = int.Parse(CboPhase.Text),
                Class = _class,
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
            if (_handle)
            {
                ClearData();
                HandleCourseSelection();
            }

            _handle = true;
        }

        private void HandleCourseSelection()
        {
            switch (CboCourseSelection.SelectedItem?.ToString().Split(new[] {": "}, StringSplitOptions.None).Last())
            {
                case "13M10 -- MOSQ":
                    _class = ClassType.Mosq;
                    PopulateMosQData();
                    CboPhase.IsEnabled = false;
                    DtDateOfCounselingMidCourse.Visibility = Visibility.Visible;
                    LblMidCourseDate.Visibility = Visibility.Visible;
                    TabMidCourse.Visibility = Visibility.Visible;
                    break;
                case "13M30 -- ALC":
                    _class = ClassType.Alc;
                    DtDateOfCounselingMidCourse.Visibility = Visibility.Hidden;
                    LblMidCourseDate.Visibility = Visibility.Hidden;
                    TabMidCourse.Visibility = Visibility.Hidden;
                        
                    PopulateAlcData();
                    break;
                case "13M40 -- SLC":
                    _class = ClassType.Slc;
                    DtDateOfCounselingMidCourse.Visibility = Visibility.Hidden;
                    LblMidCourseDate.Visibility = Visibility.Hidden;
                    TabMidCourse.Visibility = Visibility.Hidden;
                    PopulateSlcData();
                    break;
            }
        }

        public void PopulateMosQData()
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

            ComboBoxItem typeItem = (ComboBoxItem)CboPhase.SelectedItem;
            var value = typeItem.Content.ToString();

            TxtPurposeInitial.Text = TxtPurposeInitial.Text.Replace("%PHASE%", value);
            TxtLeaderResponsibilitiesInitial.Text = TxtLeaderResponsibilitiesInitial.Text.Replace("%PHASE%", value);
            TxtAssessmentInitial.Text = TxtAssessmentInitial.Text.Replace("%PHASE%", value);

            switch (value)
            {
                case "1":
                    TxtPurposeEndCourse.Text = SlcDefaultPhase1.EndCoursePurpose;
                    TxtKeyPointsEndCourse.Text = SlcDefaultPhase1.EndCourseKeyPoints;
                    TxtPlanOfActionEndCourse.Text = SlcDefaultPhase1.EndCoursePlanOfAction;
                    TxtLeaderResponsibilitiesEndCourse.Text = SlcDefaultPhase1.EndCourseLeaderResponsibilities;
                    TxtAssessmentEndCourse.Text = SlcDefaultPhase1.EndCourseAssessment;
                    break;
                case "2":
                    TxtPurposeEndCourse.Text = SlcDefaultPhase2.EndCoursePurpose;
                    TxtKeyPointsEndCourse.Text = SlcDefaultPhase2.EndCourseKeyPoints;
                    TxtPlanOfActionEndCourse.Text = SlcDefaultPhase2.EndCoursePlanOfAction;
                    TxtLeaderResponsibilitiesEndCourse.Text = SlcDefaultPhase2.EndCourseLeaderResponsibilities;
                    TxtAssessmentEndCourse.Text = SlcDefaultPhase2.EndCourseAssessment;
                    break;
            }
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

            CboPhase.SelectedIndex = 0;
            CboPhase.IsEnabled = true;
        }

        private void CboPhase_DropDownClosed(object sender, EventArgs e)
        {
            HandleCourseSelection();
        }
    }
}