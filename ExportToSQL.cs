
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using MSProject = Microsoft.Office.Interop.MSProject;
using System.Data.Entity;
using Office = Microsoft.Office.Core;
using DQIDB;

namespace GetProjectData
{
    public partial class ExportToSQL
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
   //         this.Application.NewProject += new Microsoft.Office.Interop.MSProject._EProjectApp2_NewProjectEventHandler(Application_NewProject);

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        void Application_NewProject(Microsoft.Office.Interop.MSProject.Project pj)

        {
            object missingValue = System.Reflection.Missing.Value;
     

           this.Application.FileOpenEx("https://intranet.floridahousing.org/DMS/DataRemediation/DQM%20-%20Batch%205/SiteAssets/DQM%20Archive%20Batch%205%20-Tasks.mpp",
                     missingValue, missingValue, missingValue, missingValue,
                     missingValue, missingValue, missingValue, missingValue,
                     missingValue, missingValue, MSProject.PjPoolOpen.pjPoolReadOnly,
                     missingValue, missingValue, missingValue, missingValue,
                     missingValue);

       

            MSProject.Project project = this.Application.ActiveProject;

            using (var context = new DQIEntities())
            context.Database.ExecuteSqlCommand("Exec dbo.usp_PRJ_003_TruncGraphTables");

            using (var context = new DQIEntities())

     
                    foreach (MSProject.Task task in project.Tasks)
                    {


                        var Grh_Tasks = new Grh_Tasks
                        {
                            ID = task.ID,
                            Name = task.Name,
                                 Status = task.Status.ToString(),
                                 ResourceNames = task.ResourceNames.ToString(),
                                Duration    = Convert.ToInt32(task.Duration),
                               Finish_Date =  Convert.ToDateTime(task.Finish),
                            OutlineLevel = Convert.ToInt16(task.OutlineLevel),
                            C__Complete  = Convert.ToInt16(task.PercentComplete)
                        };
                        context.Grh_Tasks.Add(Grh_Tasks);
                        context.SaveChanges();
                    }

                    this.Application.FileExit(MSProject.PjSaveType.pjDoNotSave);
                

        }



        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
