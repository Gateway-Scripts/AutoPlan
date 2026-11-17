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
                if (args.Length > 0)
                {
                    _patientId = args.First();
                }
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
                "ARC",//"STATIC"
                null);//primary fluence mode can be "FFF"
            //foreach(double gantry in gantryAngles)
            //{
            //    plan.AddStaticBeam(parameters,
            //        new VRect<double>(-50, -50, 50, 50),//10x10 field to start.
            //        0,
            //        gantry,
            //        0,
            //        target.CenterPoint);//isocenter to the center of target (warning shifts may be precise).
            //    //alternatively for isocenter can use structureSet.Image.UserOrigin if user origin is set.
            //}
            //add 2 arcs
            for(int i = 0; i < 2; i++)
            {
                Beam beam = plan.AddConformalArcBeam(parameters,
                    i % 2 == 0 ? 30 : 330,//if its arc 1, set collimator to 30, for arc 2 set to 330.
                    20,
                    i % 2 == 0 ? 195 : 165,//ternary operator
                    i % 2 == 0 ? 165 : 195,
                    i % 2 == 0 ? GantryDirection.Clockwise : GantryDirection.CounterClockwise,
                    0,
                    target.CenterPoint);
                beam.FitCollimatorToStructure(new FitToStructureMargins(5), target, true, true, false);
                //check X-jaw size and reduce size if > 150mm.
            }
            Console.WriteLine($"{plan.Beams.Count()} beams added to the plan");

            //allow the user to select a RapidPlan Model.
            int rpIndex = 0;
            foreach(var rp in app.Calculation.GetDvhEstimationModelSummaries())//.Where(r=>r.IsPublished)//.Where(r=>r.IsTrained)//OrderBy(r=>r.Name)
            {
                Console.WriteLine($"[{rpIndex}]. {rp.Name} - {rp.TreatmentSite}");
                rpIndex++;
            }
            Console.WriteLine("Select a RapidPlan Model");
            int rpSelect = Convert.ToInt32(Console.ReadLine());
            var rpModel = app.Calculation.GetDvhEstimationModelSummaries().ElementAt(rpSelect);

            Dictionary<string, DoseValue> targetMatches = new Dictionary<string, DoseValue>();
            Dictionary<string, string> structureMatches = new Dictionary<string, string>();

            var rpModelStructures = app.Calculation.GetDvhEstimationModelStructures(rpModel.ModelUID);
            //find a match to all plan structures based on the ID.
            foreach(Structure structure in structureSet.Structures)
            {
                if(structure.DicomType == "PTV")
                {
                    //add to structure matches and to target matches.
                    if (rpModelStructures.Any(s => s.Id.Equals(structure.Id, StringComparison.OrdinalIgnoreCase)))
                    {
                        var rpStructureMatch = rpModelStructures.First(s => s.Id.Equals(structure.Id, StringComparison.OrdinalIgnoreCase));
                        structureMatches.Add(structure.Id, rpStructureMatch.Id);
                        targetMatches.Add(structure.Id, plan.TotalDose);                        
                    }
                    else
                    {
                        //TODO: handle if your PTV does not find a match on ID.
                        //Perhaps match on rpStructureMatch.StructureType == DVHEstimationStructureType.PTV;
                    }
                }
                else
                {
                    //normal structures
                    if (rpModelStructures.Any(s => s.Id.Equals(structure.Id, StringComparison.OrdinalIgnoreCase)))
                    {
                        var rpStructureMatch = rpModelStructures.First(s => s.Id.Equals(structure.Id, StringComparison.OrdinalIgnoreCase));
                        structureMatches.Add(structure.Id, rpStructureMatch.Id);
                    }
                }
            }
            Console.WriteLine($"Structures matched: \n\t{String.Join("\n\t",structureMatches.Select(s=>$"{s.Key}=>{s.Value}"))}");
            Console.WriteLine("Calculating DVH Estimates...");
            plan.CalculateDVHEstimates(rpModel.Name, targetMatches, structureMatches);
            //add a ring if one does not exist. And Add the normal tissue objective.
            string ringId = "NS_Ring05";
            Structure ring;
            if(structureSet.Structures.Any(st=>st.Id.Equals(ringId, StringComparison.OrdinalIgnoreCase)))
            {
                ring = structureSet.Structures.First(st => st.Id.Equals(ringId, StringComparison.OrdinalIgnoreCase));
            }
            else
            {
                ring = structureSet.AddStructure("CONTROL", ringId);
                ring.SegmentVolume = target.Margin(5).Sub(target);
            }
            //to manually add objectives to the optimization.
            plan.OptimizationSetup.AddPointObjective(ring, 
                OptimizationObjectiveOperator.Upper,
                plan.TotalDose * 1.05,
                0,
                100);
            //add the NTO
            plan.OptimizationSetup.AddAutomaticNormalTissueObjective(100);

            Console.WriteLine("Optimizing...");
            plan.OptimizationSetup.UseJawTracking = true;
            plan.OptimizeVMAT(new OptimizationOptionsVMAT(OptimizationIntermediateDoseOption.UseIntermediateDose,
                plan.Beams.First().MLC.Id));
            //IMRT only.
            //Console.WriteLine("Calculating leaf motions...");
            //plan.CalculateLeafMotions();
            Console.WriteLine("Calculating Dose...");
            plan.CalculateDose();
            //this should always be the last line in a stand-alone automation script.


            app.SaveModifications();
            Console.WriteLine("Application saved. Reload patient in Eclipse.");
            
        }
    }
}
