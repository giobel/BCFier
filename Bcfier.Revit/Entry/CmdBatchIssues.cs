using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Bcfier.Bcf;
using Bcfier.Bcf.Bcf2;
using Bcfier.Data.Utils;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Media;

namespace Bcfier.Revit.Entry
{

  [Transaction(TransactionMode.Manual)]
  [Regeneration(RegenerationOption.Manual)]
  public class CmdBatchIssues : IExternalCommand
  {

    /// <summary>
    /// Main Command Entry Point
    /// </summary>
    /// <param name="commandData"></param>
    /// <param name="message"></param>
    /// <param name="elements"></param>
    /// <returns></returns>
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
      try
      {
                UIApplication uiapp = commandData.Application;
                UIDocument uidoc = uiapp.ActiveUIDocument;

                BcfContainer bcfContainer = new BcfContainer();
                BcfFile bcfFile = new BcfFile();

                Markup issue = new Markup(DateTime.Now);

                Topic tp = new Topic();
                tp.Title = "My first Issue";
                tp.Description = "a box and nothing else";
                issue.Topic = tp;

                

                var view = new ViewPoint(issue.Viewpoints.Any());

                if (!Directory.Exists(Path.Combine(bcfFile.TempPath, issue.Topic.Guid)))
                    Directory.CreateDirectory(Path.Combine(bcfFile.TempPath, issue.Topic.Guid));



                var c = new Comment
                {
                    Comment1 = "Comment something",
                    Author = Utils.GetUsername(),
                    //Status = comboStatuses.SelectedValue.ToString(),
                    //VerbalStatus = VerbalStatus.Text,
                    Date = DateTime.Now,
                    Viewpoint = new CommentViewpoint { Guid = view.Guid }
                    };
                    issue.Comment.Add(c);
                




                //first save the image, then update path
                string path = Path.Combine(bcfFile.TempPath, issue.Topic.Guid, view.Snapshot);
                SaveRevitSnapshot(uidoc.Document, path);
                view.SnapshotPath = path;

                //neede for UI binding
                issue.Viewpoints.Add(view);

                issue.Viewpoints.Last().VisInfo = Data.RevitView.GenerateViewpoint(uiapp.ActiveUIDocument);

                if (uidoc.Document.Title != null)
                    issue.Header[0].Filename = uidoc.Document.Title;
                else
                    issue.Header[0].Filename = "Unknown";


                bcfFile.Issues.Add(issue);
                bcfContainer.SaveFile(bcfFile);


                return Result.Succeeded;

      }
      catch (Exception e)
      {
        message = e.Message;
        return Result.Failed;
      }

    }
        private string SaveRevitSnapshot(Document doc, string path)
        {
            try
            {
                //string tempImg = Path.Combine(Path.GetTempPath(), "BCFier", Path.GetTempFileName() + ".png");
                string tempImg = path;
                var options = new ImageExportOptions
                {
                    FilePath = tempImg,
                    HLRandWFViewsFileType = ImageFileType.PNG,
                    ShadowViewsFileType = ImageFileType.PNG,
                    ExportRange = ExportRange.VisibleRegionOfCurrentView,
                    ZoomType = ZoomFitType.FitToPage,
                    ImageResolution = ImageResolution.DPI_72,
                    PixelSize = 1000
                };
                doc.ExportImage(options);

                //File.Delete(tempImg);

                return tempImg;
            }
            catch (System.Exception ex1)
            {
                TaskDialog.Show("Error!", "exception: " + ex1);
                return null;
            }
        }
    }
}