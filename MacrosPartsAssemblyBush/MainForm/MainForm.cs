using System;
using System.Windows.Forms;
//using SolidWorks.Interop.sldworks;
//using SolidWorks.Interop.swcommands;
//using SolidWorks.Interop.swconst;

namespace MacrosPartsAssemblyBush
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        // forbidding anything but digits
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(textBox1.Text, "[^0-9]"))
            {
                textBox1.Text = textBox1.Text.Remove(textBox1.Text.Length - 1);
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(textBox2.Text, "[^0-9]"))
            {
                textBox2.Text = textBox2.Text.Remove(textBox2.Text.Length - 1);
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(textBox3.Text, "[^0-9]"))
            {
                textBox3.Text = textBox3.Text.Remove(textBox3.Text.Length - 1);
            }
        }

        // reference to SW window
        SldWorks swApp;

        // model parts
        private void button1_Click(object sender, EventArgs e)
        {
            // preventing empty fields
            if (string.IsNullOrWhiteSpace(textBox1.Text) || string.IsNullOrWhiteSpace(textBox2.Text)
                || string.IsNullOrWhiteSpace(textBox3.Text))
            {
                MessageBox.Show("Введіть всі значення!", "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            double x1 = Convert.ToDouble(textBox1.Text);
            double x2 = Convert.ToDouble(textBox2.Text);
            double x3 = Convert.ToDouble(textBox3.Text);

            // checking detail size condition
            if (x2 >= 42 - 5)
            {
                MessageBox.Show("Width of base can not exceed or equal the diameter of the cylinder",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (x2 < 10)
            {
                MessageBox.Show("Width of base can not be less than the width of the rib",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (x1 <= 42)
            {
                MessageBox.Show("Length of base can not be less than or equal " +
                    "to the diameter of the cylinder",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (x3 <= 19)
            {
                MessageBox.Show("Height of detail can not be less than or equal " +
                    "to the length of the inner section of the cylinder",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // m to mm conversion
            x1 = -x1 / 1000;
            x2 = -x2 / 1000;
            x3 = -x3 / 1000;

            // Opening SW 2016
            Guid Guid1 = new Guid("F16137AD-8EE8-4D2A-8CAC-DFF5D1F67522");
            object processSW = System.Activator.CreateInstance(System.Type.GetTypeFromCLSID(Guid1));

            swApp = (SldWorks)processSW;
            swApp.Visible = true;

            // selection of detail modeling mode
            swApp.NewPart();
            ModelDoc2 swDoc = null;
            bool boolstatus = false;
            int longstatus = 0;

            // first detail build
            swDoc = ((ModelDoc2)(swApp.ActiveDoc));
            SketchSegment skSegment = null;
            boolstatus = swDoc.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swDoc.SketchManager.InsertSketch(true);
            swDoc.ClearSelection2(true);
            skSegment = ((SketchSegment)(swDoc.SketchManager.CreateCircle(0, 0, 0, -0.021, 0, 0)));
            swDoc.ShowNamedView2("*Trimetric", 8);
            swDoc.ClearSelection2(true);
            boolstatus = swDoc.Extension.SelectByID2("Arc1", "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
            Feature myFeature = null;
            myFeature = ((Feature)(swDoc.FeatureManager.FeatureExtrusion2(true, false, false, 0, 0, 0.01, 0.01, false, false, false, false, 0.017453292519943334, 0.017453292519943334, false, false, false, false, true, true, true, 0, 0, false)));
            swDoc.ISelectionManager.EnableContourSelection = false;
            boolstatus = swDoc.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swDoc.SketchManager.InsertSketch(true);
            swDoc.ClearSelection2(true);
            boolstatus = swDoc.Extension.SetUserPreferenceToggle(((int)(swUserPreferenceToggle_e.swSketchAddConstToRectEntity)), ((int)(swUserPreferenceOption_e.swDetailingNoOptionSpecified)), true);
            boolstatus = swDoc.Extension.SetUserPreferenceToggle(((int)(swUserPreferenceToggle_e.swSketchAddConstLineDiagonalType)), ((int)(swUserPreferenceOption_e.swDetailingNoOptionSpecified)), true);
            Array vSkLines = null;
            vSkLines = ((Array)(swDoc.SketchManager.CreateCenterRectangle(0, 0, 0, x1 / 2, x2 / 2, 0)));
            swDoc.ViewZoomtofit2();
            swDoc.ShowNamedView2("*Top", 5);
            swDoc.ClearSelection2(true);
            skSegment = ((SketchSegment)(swDoc.SketchManager.CreateLine(x1 / 2, Abs(x2 / 2), 0.000000, 0.000000, 0.021000, 0.000000)));
            skSegment = ((SketchSegment)(swDoc.SketchManager.CreateLine(0.000000, 0.021000, 0.000000, Abs(x1 / 2), Abs(x2 / 2), 0.000000)));
            skSegment = ((SketchSegment)(swDoc.SketchManager.CreateLine(x1 / 2, x2 / 2, 0.000000, 0.000000, -0.021000, 0.000000)));
            skSegment = ((SketchSegment)(swDoc.SketchManager.CreateLine(0.000000, -0.021000, 0.000000, Abs(x1 / 2), x2 / 2, 0.000000)));
            swDoc.ClearSelection2(true);
            boolstatus = swDoc.Extension.SelectByID2("Line1", "SKETCHSEGMENT", -0.0044024507535198876, 0, x2 / 2, false, 2, null, 0);
            boolstatus = swDoc.SketchManager.SketchTrim(4, 0, 0, 0);
            boolstatus = swDoc.Extension.SelectByID2("Line6", "SKETCHSEGMENT", -0.012921376009237012, 0, -0.0054405793723103189, false, 2, null, 0);
            boolstatus = swDoc.SketchManager.SketchTrim(4, 0, 0, 0);
            boolstatus = swDoc.Extension.SelectByID2("Line5", "SKETCHSEGMENT", -0.014135271591155435, 0, 0.005951693301539128, false, 2, null, 0);
            boolstatus = swDoc.SketchManager.SketchTrim(4, 0, 0, 0);
            boolstatus = swDoc.Extension.SelectByID2("Line5", "SKETCHSEGMENT", 0.01393190144612956, 0, -0.0058660637667913938, false, 2, null, 0);
            boolstatus = swDoc.SketchManager.SketchTrim(4, 0, 0, 0);
            boolstatus = swDoc.Extension.SelectByID2("Line6", "SKETCHSEGMENT", 0.011650517863344173, 0, 0.0049054812056186, false, 2, null, 0);
            boolstatus = swDoc.SketchManager.SketchTrim(4, 0, 0, 0);
            boolstatus = swDoc.Extension.SelectByID2("Line3", "SKETCHSEGMENT", -0.016022866314760208, 0, Abs(x2 / 2), false, 2, null, 0);
            boolstatus = swDoc.SketchManager.SketchTrim(4, 0, 0, 0);
            swDoc.ClearSelection2(true);
            myFeature = ((Feature)(swDoc.FeatureManager.FeatureExtrusion2(true, false, false, 0, 0, 0.01, 0.01, false, false, false, false, 0.017453292519943334, 0.017453292519943334, false, false, false, false, true, true, true, 0, 0, false)));
            swDoc.ISelectionManager.EnableContourSelection = false;
            boolstatus = swDoc.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swDoc.SketchManager.InsertSketch(true);
            swDoc.ClearSelection2(true);
            skSegment = ((SketchSegment)(swDoc.SketchManager.CreateCircle(0.000000, 0.000000, 0.000000, -0.009, 0, 0.000000)));
            swDoc.ClearSelection2(true);
            boolstatus = swDoc.Extension.SelectByID2("Arc1", "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
            myFeature = ((Feature)(swDoc.FeatureManager.FeatureCut3(true, false, true, 0, 0, 0.01, 0.01, false, false, false, false, 0.017453292519943334, 0.017453292519943334, false, false, false, false, false, true, true, true, true, false, 0, 0, false)));
            swDoc.ISelectionManager.EnableContourSelection = false;

            // saving the detail
            swDoc.ClearSelection2(true);
            longstatus = swDoc.SaveAs3(@"G:\1.SLDPRT", 0, 2);

            // second detail build
            swApp.NewPart();
            swDoc = ((ModelDoc2)(swApp.ActiveDoc));
            boolstatus = swDoc.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swDoc.SketchManager.InsertSketch(true);
            swDoc.ClearSelection2(true);
            skSegment = null;
            skSegment = ((SketchSegment)(swDoc.SketchManager.CreateCircle(0.000000, 0.000000, 0.000000, -0.021, 0, 0.000000)));
            swDoc.ShowNamedView2("*Trimetric", 8);
            swDoc.ClearSelection2(true);
            boolstatus = swDoc.Extension.SelectByID2("Arc1", "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
            myFeature = null;
            myFeature = ((Feature)(swDoc.FeatureManager.FeatureExtrusion2(true, false, false, 0, 0, 0.035000000000000003, 0.01, false, false, false, false, 0.017453292519943334, 0.017453292519943334, false, false, false, false, true, true, true, 0, 0, false)));
            swDoc.ISelectionManager.EnableContourSelection = false;
            boolstatus = swDoc.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swDoc.SketchManager.InsertSketch(true);
            swDoc.ClearSelection2(true);
            skSegment = ((SketchSegment)(swDoc.SketchManager.CreateCircle(0.000000, 0.000000, 0.000000, -0.016, 0, 0.000000)));
            swDoc.ClearSelection2(true);
            boolstatus = swDoc.Extension.SelectByID2("Arc1", "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
            myFeature = ((Feature)(swDoc.FeatureManager.FeatureCut3(true, false, true, 0, 0, 0.035000000000000003, 0.035000000000000003, false, false, false, false, 0.017453292519943334, 0.017453292519943334, false, false, false, false, false, true, true, true, true, false, 0, 0, false)));
            swDoc.ISelectionManager.EnableContourSelection = false;
            boolstatus = swDoc.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swDoc.SketchManager.InsertSketch(true);
            swDoc.ClearSelection2(true);
            skSegment = ((SketchSegment)(swDoc.SketchManager.CreateCircle(0.000000, 0.000000, 0.000000, -0.016, 0, 0.000000)));
            swDoc.ClearSelection2(true);
            boolstatus = swDoc.Extension.SelectByID2("Arc1", "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
            myFeature = ((Feature)(swDoc.FeatureManager.FeatureExtrusion2(true, false, false, 0, 0, 0.017000000000000001, 0.035000000000000003, false, false, false, false, 0.017453292519943334, 0.017453292519943334, false, false, false, false, true, true, true, 0, 0, false)));
            swDoc.ISelectionManager.EnableContourSelection = false;
            boolstatus = swDoc.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swDoc.SketchManager.InsertSketch(true);
            swDoc.ClearSelection2(true);
            skSegment = ((SketchSegment)(swDoc.SketchManager.CreateCircle(0.000000, 0.000000, 0.000000, -0.009, 0, 0.000000)));
            swDoc.ClearSelection2(true);
            boolstatus = swDoc.Extension.SelectByID2("Arc1", "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
            myFeature = ((Feature)(swDoc.FeatureManager.FeatureCut3(true, false, true, 0, 0, 0.017999999999999999, 0.017000000000000001, false, false, false, false, 0.017453292519943334, 0.017453292519943334, false, false, false, false, false, true, true, true, true, false, 0, 0, false)));
            swDoc.ISelectionManager.EnableContourSelection = false;

            // saving the detail
            swDoc.ClearSelection2(true);
            longstatus = swDoc.SaveAs3(@"G:\2.SLDPRT", 0, 2);

            // third detail build
            swApp.NewPart();
            swDoc = ((ModelDoc2)(swApp.ActiveDoc));
            boolstatus = swDoc.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swDoc.SketchManager.InsertSketch(true);
            swDoc.ClearSelection2(true);
            boolstatus = swDoc.Extension.SetUserPreferenceToggle(((int)(swUserPreferenceToggle_e.swSketchAddConstToRectEntity)), ((int)(swUserPreferenceOption_e.swDetailingNoOptionSpecified)), true);
            boolstatus = swDoc.Extension.SetUserPreferenceToggle(((int)(swUserPreferenceToggle_e.swSketchAddConstLineDiagonalType)), ((int)(swUserPreferenceOption_e.swDetailingNoOptionSpecified)), true);
            vSkLines = null;
            vSkLines = ((Array)(swDoc.SketchManager.CreateCenterRectangle(0, 0, 0, x1 / 2, -0.005, 0)));
            swDoc.ShowNamedView2("*Trimetric", 8);
            swDoc.ClearSelection2(true);
            boolstatus = swDoc.Extension.SelectByID2("Line5", "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
            boolstatus = swDoc.Extension.SelectByID2("Line6", "SKETCHSEGMENT", 0, 0, 0, true, 0, null, 0);
            boolstatus = swDoc.Extension.SelectByID2("Point1", "SKETCHPOINT", 0, 0, 0, true, 0, null, 0);
            boolstatus = swDoc.Extension.SelectByID2("Line2", "SKETCHSEGMENT", 0, 0, 0, true, 0, null, 0);
            boolstatus = swDoc.Extension.SelectByID2("Line1", "SKETCHSEGMENT", 0, 0, 0, true, 0, null, 0);
            boolstatus = swDoc.Extension.SelectByID2("Line4", "SKETCHSEGMENT", 0, 0, 0, true, 0, null, 0);
            boolstatus = swDoc.Extension.SelectByID2("Line3", "SKETCHSEGMENT", 0, 0, 0, true, 0, null, 0);
            myFeature = null;
            myFeature = ((Feature)(swDoc.FeatureManager.FeatureExtrusion2(true, false, false, 0, 0, 0.035000000000000003, 0.01, false, false, false, false, 0.017453292519943334, 0.017453292519943334, false, false, false, false, true, true, true, 0, 0, false)));
            swDoc.ISelectionManager.EnableContourSelection = false;
            swDoc.ShowNamedView2("*Front", 1);
            boolstatus = swDoc.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swDoc.SketchManager.InsertSketch(true);
            swDoc.ClearSelection2(true);
            skSegment = null;
            skSegment = ((SketchSegment)(swDoc.SketchManager.CreateLine(x1 / 2, 0.000000, 0.000000, x1 / 2, 0.010000, 0.000000)));
            skSegment = ((SketchSegment)(swDoc.SketchManager.CreateLine(x1 / 2, 0.010000, 0.000000, -0.020908, 0.035000, 0.000000)));
            swDoc.ClearSelection2(true);
            boolstatus = swDoc.Extension.SelectByID2("Line3", "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
            myFeature = ((Feature)(swDoc.FeatureManager.FeatureCut3(false, false, false, 1, 1, 0.035000000000000003, 0.035000000000000003, false, false, false, false, 0.017453292519943334, 0.017453292519943334, false, false, false, false, false, true, true, true, true, false, 0, 0, false)));
            swDoc.ISelectionManager.EnableContourSelection = false;
            boolstatus = swDoc.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swDoc.SketchManager.InsertSketch(true);
            swDoc.ClearSelection2(true);
            skSegment = ((SketchSegment)(swDoc.SketchManager.CreateLine(Abs(x1 / 2), 0.000000, 0.000000, Abs(x1 / 2), 0.010000, 0.000000)));
            skSegment = ((SketchSegment)(swDoc.SketchManager.CreateLine(Abs(x1 / 2), 0.010000, 0.000000, 0.020908, 0.035000, 0.000000)));
            swDoc.ClearSelection2(true);
            boolstatus = swDoc.Extension.SelectByID2("Line3", "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
            myFeature = ((Feature)(swDoc.FeatureManager.FeatureCut3(false, true, false, 1, 1, 0.035000000000000003, 0.035000000000000003, false, false, false, false, 0.017453292519943334, 0.017453292519943334, false, false, false, false, false, true, true, true, true, false, 0, 0, false)));
            swDoc.ISelectionManager.EnableContourSelection = false;
            boolstatus = swDoc.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swDoc.SketchManager.InsertSketch(true);
            swDoc.ShowNamedView2("*Top", 5);
            swDoc.ClearSelection2(true);
            skSegment = ((SketchSegment)(swDoc.SketchManager.CreateCircle(0.000000, 0.000000, 0.000000, -0.018, 0, 0.000000)));
            swDoc.ClearSelection2(true);
            boolstatus = swDoc.Extension.SelectByID2("Arc1", "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
            swDoc.ClearSelection2(true);
            boolstatus = swDoc.Extension.SelectByID2("Arc1", "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
            swDoc.ClearSelection2(true);
            boolstatus = swDoc.Extension.SelectByID2("Arc1", "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
            swDoc.ClearSelection2(true);
            boolstatus = swDoc.Extension.SelectByID2("Arc1", "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
            myFeature = ((Feature)(swDoc.FeatureManager.FeatureCut3(true, false, true, 0, 0, 0.035000000000000003, 0.035000000000000003, false, false, false, false, 0.017453292519943334, 0.017453292519943334, false, false, false, false, false, true, true, true, true, false, 0, 0, false)));
            swDoc.ISelectionManager.EnableContourSelection = false;

            // saving the detail
            swDoc.ClearSelection2(true);
            longstatus = swDoc.SaveAs3(@"G:\3.SLDPRT", 0, 2);
        }

        // creating an assembly
        private void button2_Click(object sender, EventArgs e)
        {
            if (swApp == null)
            {
                MessageBox.Show("You have to build all three details beforehand",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else
            {
                AssemblyDoc swAssembly = swApp.INewAssembly();
                ModelDoc2 swDoc = (ModelDoc2)swApp.ActiveDoc;
                int longwarnings = 0;
                int longstatus = 0;

                Component2 com = default(Component2);
                swApp.OpenDoc6(@"G:\1.SLDPRT", 1, 0, "", longstatus, longwarnings);
                com = swAssembly.AddComponent4(@"G:\1.SLDPRT", "Default", 0, -0.0125, 0);
                swApp.CloseDoc(@"G:\1.SLDPRT");

                swApp.OpenDoc6(@"G:\2.SLDPRT", 1, 0, "", longstatus, longwarnings);
                com = swAssembly.AddComponent4(@"G:\2.SLDPRT", "Default", 0, 0, 0);
                swApp.CloseDoc(@"G:\2.SLDPRT");


                swApp.OpenDoc6(@"G:\3.SLDPRT", 1, 0, "", longstatus, longwarnings);
                com = swAssembly.AddComponent4(@"G:\3.SLDPRT", "Default", 0, 0, 0);
                swApp.CloseDoc(@"G:\3.SLDPRT");
            }
        }

    }
}
