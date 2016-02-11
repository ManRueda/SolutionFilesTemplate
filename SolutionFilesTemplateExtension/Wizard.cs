using System;
using System.Collections.Generic;
using System.Linq;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.TemplateWizard;
using System.IO;

namespace SolutionFilesTemplateExtension
{
    public class Wizard : IWizard
    {
        private DTE _dte;

        public void BeforeOpeningFile(ProjectItem project) { }
        public void ProjectItemFinishedGenerating(ProjectItem projectItem) { }
        public void RunFinished() { }

        public void RunStarted(object automationObject, Dictionary<string, string> replacementsDictionary, WizardRunKind runKind, object[] customParams)
        {
            _dte = (DTE)automationObject;

            // add an entry to the dictionary to specify the string used for the $greeting$ token 
            //replacementsDictionary.Add("$greeting$", "Hello Custom Library");
        }

        public void ProjectFinishedGenerating(Project project)
        {
            //Get the DTE related to the solution
            EnvDTE.DTE dte = project.DTE;

            //Get the first project of all the solution project
            Project firstProject = dte.Solution.Projects.Cast<Project>().FirstOrDefault();

            //Cast the project items as IEnumerable<ProjectItem>
            IEnumerable<ProjectItem> projectItems = firstProject.ProjectItems.Cast<ProjectItem>();

            //Cast the solution as s Solution2 (has the AddSolutionFolder method)
            Solution2 solution = (Solution2)dte.Solution;

            //Get the 
            ProjectItem solutionFilesItem = projectItems.Single(i => i.Name == "SolutionFiles");

            if (solutionFilesItem != null)
            {
                // Get the LocalPath property
                Property prop = solutionFilesItem.Properties.Cast<Property>().Single(i => i.Name == "LocalPath");

                //Get the attribute and check if it a folder or not
                if (File.GetAttributes(prop.Value.ToString()).HasFlag(FileAttributes.Directory))
                {
                    //Create a new solution folder and return a Project object
                    Project solutionFolderProject = solution.AddSolutionFolder("Solution Files");

                    //Cast the Project object to a SolutionFolder to be able to create more Solution Folders children
                    SolutionFolder solutionFolder = (SolutionFolder)solutionFolderProject.Object;

                    //Call the function to add the original files to the solution folder
                    addFilesToSolution(solutionFolderProject, solutionFolder, prop.Value.ToString());

                    //Delete the original files and remove the project reference
                    Directory.Delete(prop.Value.ToString(), true);
                    solutionFilesItem.Remove();
                }
                else
                {

                }

            }
        }

        public void addFilesToSolution(Project solutionFolderProject, SolutionFolder solutionFolder, string basePath)
        {
            //For each file, add to the folder
            foreach (var file in Directory.EnumerateFiles(basePath))
            {
                //Add the file as a copy into the solution folder
                solutionFolderProject.ProjectItems.AddFromFileCopy(file);
            }
            //for each directory, create a new solution folder and add the files
            foreach (var dir in Directory.EnumerateDirectories(basePath))
            {
                //Create a new solution folder and return a Project object
                Project folderProject = solutionFolder.AddSolutionFolder(Path.GetFileName(dir));
                //Cast the Project object to a SolutionFolder to be able to create more Solution Folders children
                SolutionFolder solFolder = (SolutionFolder)folderProject.Object;
                //Call the function to add the original files to the solution folder
                addFilesToSolution(folderProject, solFolder, dir);
            }
        }

        public bool ShouldAddProjectItem(string filePath) { return true; }
    }
}
