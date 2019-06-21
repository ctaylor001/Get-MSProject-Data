
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using MSProject = Microsoft.Office.Interop.MSProject;
using System.Data.Entity;
using Office = Microsoft.Office.Core;
using DQIDB;
using DQILogger.Core;

namespace GetProjectData
{
    public partial class ExportToSQL
    {

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            var fd = GetFlogDetail("starting application", null);
            Logger.WriteDiagnostic(fd);


            var tracker = new PerfTracker("GetProjectData_Execution", "", fd.UserName,
              fd.Location, fd.TaskName, fd.Layer);

            var filelocate = new List<String>();
            using (var context = new DQIEntities())

                foreach (var MsProjLocation in context.DMSBatchControlLogs.ToList())
                {
                    string filelocation = (MsProjLocation.MSProjLocation);
                    filelocate.Add(filelocation);
                }



            using (var context = new DQIEntities())
                context.Database.ExecuteSqlCommand("Exec dbo.usp_PRJ_003_TruncGraphTables");

            foreach (var filelocation in filelocate)
            {
                object missingValue = System.Reflection.Missing.Value;
                this.Application.FileOpenEx(filelocation,
                          missingValue, missingValue, missingValue, missingValue,
                          missingValue, missingValue, missingValue, missingValue,
                          missingValue, missingValue, MSProject.PjPoolOpen.pjPoolReadOnly,
                          missingValue, missingValue, missingValue, missingValue,
                          missingValue);

                //    this.Application.NewProject += new Microsoft.Office.Interop.MSProject._EProjectApp2_NewProjectEventHandler(Application_NewProject);



                /*
                var filelocate = new List<String>();
                using (var context = new DQIEntities())

                    foreach (var MsProjLocation in context.DMSBatchControlLogs.ToList())
                    {
                        string filelocation = (MsProjLocation.MSProjLocation);
                        filelocate.Add(filelocation);
                    }
                    using (var context = new DQIEntities())
                    context.Database.ExecuteSqlCommand("Exec dbo.usp_PRJ_003_TruncGraphTables");
                */
                try
                {

                    //using (var context = new DQIEntities())
                    // foreach (var filelocation in filelocate)
                    //Not bad works - foreach (var MsProjLocation in context.DMSBatchControlLogs.ToList())
                    // bad - 1 foreach (var MsProjLocation in context.DMSBatchControlLogs)

                    // String filelocation = MsProjLocation.MSProjLocation;


                    MSProject.Project project = this.Application.ActiveProject;


           
                    using (var context = new DQIEntities())
                        foreach (MSProject.Task task in project.Tasks)
                        {
                            var Grh_Tasks = new Grh_Tasks
                            {
                                ID = task.ID,
                                Name = task.Name,
                                Status = task.Status.ToString(),
                                ResourceNames = task.ResourceNames.ToString(),
                                Duration = Convert.ToInt32(task.Duration),
                                Finish_Date = Convert.ToDateTime(task.Finish),
                                OutlineLevel = Convert.ToInt16(task.OutlineLevel),
                                C__Complete = Convert.ToInt16(task.PercentComplete)
                               
                            };
                            context.Grh_Tasks.Add(Grh_Tasks);
                            context.SaveChanges();
                        };
                    //    };

                    Application.FileExit(MSProject.PjSaveType.pjSave);
                    Application.FileCloseEx(MSProject.PjSaveType.pjDoNotSave);

                }

                catch (Exception ex)
                {
                    fd = GetFlogDetail("", ex);
                    Logger.WriteError(fd);
                }

                using (var context = new DQIEntities())
                    context.Database.ExecuteSqlCommand("Exec dbo.usp_PRJ_004_InsertGrh_TableReport");



            }

        }
        
    

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

       // void Application_NewProject(Microsoft.Office.Interop.MSProject.Project pj)

        //{
        //    var fd = GetFlogDetail("starting application", null);
        //    Logger.WriteDiagnostic(fd);


        //    var tracker = new PerfTracker("GetProjectData_Execution", "", fd.UserName,
        //      fd.Location, fd.TaskName, fd.Layer);

        //    /*
        //    var filelocate = new List<String>();
        //    using (var context = new DQIEntities())

        //        foreach (var MsProjLocation in context.DMSBatchControlLogs.ToList())
        //        {
        //            string filelocation = (MsProjLocation.MSProjLocation);
        //            filelocate.Add(filelocation);
        //        }
        //        using (var context = new DQIEntities())
        //        context.Database.ExecuteSqlCommand("Exec dbo.usp_PRJ_003_TruncGraphTables");
        //    */
        //    try
        //    {
                    
        //            //using (var context = new DQIEntities())
        //           // foreach (var filelocation in filelocate)
        //            //Not bad works - foreach (var MsProjLocation in context.DMSBatchControlLogs.ToList())
        //           // bad - 1 foreach (var MsProjLocation in context.DMSBatchControlLogs)
                    
        //               // String filelocation = MsProjLocation.MSProjLocation;

          
        //                MSProject.Project project = this.Application.ActiveProject;
        //                using (var context = new DQIEntities())
        //                foreach (MSProject.Task task in project.Tasks)
        //                {
        //                    var Grh_Tasks = new Grh_Tasks
        //                    {
        //                        ID = task.ID,
        //                        Name = task.Name,
        //                        Status = task.Status.ToString(),
        //                        ResourceNames = task.ResourceNames.ToString(),
        //                        Duration = Convert.ToInt32(task.Duration),
        //                        Finish_Date = Convert.ToDateTime(task.Finish),
        //                        OutlineLevel = Convert.ToInt16(task.OutlineLevel),
        //                        C__Complete = Convert.ToInt16(task.PercentComplete)
        //                    };
        //                    context.Grh_Tasks.Add(Grh_Tasks);
        //                    context.SaveChanges();
        //                };
        //        //    };

        //        Application.FileExit(MSProject.PjSaveType.pjSave);

        //    }

        //    catch (Exception ex)
        //    {
        //        fd = GetFlogDetail("", ex);
        //        Logger.WriteError(fd);
        //    }

        //    //

        //}
        ////Utility Method  Centrally setting details 


        private static LogDetail GetFlogDetail(string message, Exception ex)
        {
            return new LogDetail
            {
                TaskName = "Logger",
                Location = "GetProjectData",    // this application
                Layer = "Job",                  // unattended executable invoked somehow
                UserName = Environment.UserName,
                Hostname = Environment.MachineName,
                Message = message,
                Exception = ex
            };

        }


       /*
        private static List<string> GetmdsBatchlog(List<string> filelocate)
        {

            var filelocate = new List<String>();

            using (var context = new DQIEntities())

                foreach (var MsProjLocation in context.DMSBatchControlLogs.ToList())
                {
                    string filelocation = (MsProjLocation.MSProjLocation);
                    filelocate.Add(filelocation);
                    return (filelocate);
                }
         }
       */






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

