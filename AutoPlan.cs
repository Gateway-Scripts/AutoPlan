using System;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using System.Reflection;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

// TODO: Replace the following version attributes by creating AssemblyInfo.cs. You can do this in the properties of the Visual Studio project.
[assembly: AssemblyVersion("1.0.0.2")]
[assembly: AssemblyFileVersion("1.0.0.1")]
[assembly: AssemblyInformationalVersion("1.0")]

// TODO: Uncomment the following line if the script requires write access.
[assembly: ESAPIScript(IsWriteable = true)]

namespace AutoPlan
{
    class Program
    {
        private static string _patientId;

        [STAThread]
        static void Main(string[] args)
        {
            try
            {

                using (Application app = Application.CreateApplication())
                {
                    Execute(app);
                }//dispose
            }
            catch (Exception e)
            {
                Console.Error.WriteLine(e.ToString());
            }
        }
        static void Execute(Application app)
        {
            // TODO: Add your code here.
            //Open the patient.
            //In the future, we may want to inject the patient ID as an argument in the script.
            //create a _patientId field for storing the patientID.
            if (String.IsNullOrEmpty(_patientId))
            {
                Console.WriteLine("Input patient ID:");
                _patientId = Console.ReadLine();
            }
            Patient patient = app.OpenPatientById(_patientId);
            //Check patient ID was input correctly. End script if incorrect.
            if (patient == null) { Console.WriteLine("Invalid patient ID"); return; }
            //prior to any automation, must call patient.BeginModifications()
            patient.BeginModifications();

            //create a new course each time the script is run.
            //Course course = patient.AddCourse();//Create a new course and index the course ID.

            //Put each plan in the same course.
            string courseId = "AutoCourse";
            Course course;//course is null at this point.
            //LINQ to check if the course "AutoCourse" already exists. 
            //LINQ => Language Integrated Query. (=> Lambda expression)
            //collection.LinqExpression(placeholder => placeholder[.Property] [condition])
            if(patient.Courses.Any(c => c.Id.Equals(courseId, StringComparison.OrdinalIgnoreCase)))
            {
                course = patient.Courses.First(c => c.Id.Equals(courseId, StringComparison.OrdinalIgnoreCase));
            }
            else
            {
                course = patient.AddCourse();
                course.Id = courseId;
            }
            //have the user select the structure set. 
            int ssIndex = 0;
            foreach(StructureSet ss in patient.StructureSets)
            {
                Console.WriteLine($"[{ssIndex}]. {ss.Id} - ({ss.Image.Id} - {ss.Image.CreationDateTime.Value.ToShortDateString()})");
                ssIndex++;
            }
            Console.WriteLine("Select a structure set.");
            int ssSelect = Convert.ToInt32(Console.ReadLine());
            StructureSet structureSet = patient.StructureSets.ElementAt(ssSelect);//ElementAt is a LINQ query as well!
            //Create a new plansetup. 
            ExternalPlanSetup plan = course.AddExternalPlanSetup(structureSet);
            Console.WriteLine($"Plan {plan.Id} generated on course {course.Id}");
            //set prescription.
            //the VM Eclipse is in cGy, must use the same unit as the TPS.
            plan.SetPrescription(28, new DoseValue(250, DoseValue.DoseUnit.cGy), 1.0);

            //set the target. Method came in 16.1.
            StringBuilder errString = new StringBuilder();
            //the code below assumes only 1 PTV structure exists.
            Structure target = plan.StructureSet.Structures.FirstOrDefault(st => st.DicomType == "PTV");
            plan.SetTargetStructureIfNoDose(target, errString);
            Console.WriteLine($"Rx set to {plan.TotalDose} in {plan.NumberOfFractions}fx on target volume {target.Id}");

            double[] gantryAngles = new double[] { 180, 130, 75, 25, 335, 285, 230 };
            ExternalBeamMachineParameters parameters = new ExternalBeamMachineParameters("GS_TrueBeam",
                "10X",//do not put FFF on the energy mode.
                600,
                "STATIC",
                null);//primary fluence mode can be "FFF"
            foreach(double gantry in gantryAngles)
            {
                plan.AddStaticBeam(parameters,
                    new VRect<double>(-50, -50, 50, 50),//10x10 field to start.
                    0,
                    gantry,
                    0,
                    target.CenterPoint);//isocenter to the center of target (warning shifts may be precise).
                //alternatively for isocenter can use structureSet.Image.UserOrigin if user origin is set.
            }
            Console.WriteLine($"{plan.Beams.Count()} beams added to the plan");

            Console.WriteLine("Calculating Dose...");
            plan.CalculateDose();
            //this should always be the last line in a stand-alone automation script.


            app.SaveModifications();
            Console.WriteLine("Application saved. Reload patient in Eclipse.");
            
        }
    }
}
