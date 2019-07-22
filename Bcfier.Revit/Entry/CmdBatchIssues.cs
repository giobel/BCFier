using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Bcfier.Bcf;
using Bcfier.Bcf.Bcf2;
using Bcfier.Data.Utils;
using System;
using System.Collections.Generic;
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
                Document doc = uidoc.Document;

                BcfContainer bcfContainer = new BcfContainer();
                BcfFile bcfFile = new BcfFile();

                IList<Reference> selectedIssues = uidoc.Selection.PickObjects(ObjectType.Element, "Select issues");

                //Get the UI view
                UIView uiview = null;

                IList<UIView> uiviews = uidoc.GetOpenUIViews();

                foreach (UIView uv in uiviews)
                {
                    if (uv.ViewId.Equals(doc.ActiveView.Id))
                    {
                        uiview = uv;
                        break;
                    }
                }

                using (Transaction t = new Transaction(doc, "Export selected issues"))
                {
                    t.Start();
                    //Isolate the objects to select them easily. Reset the temporary view mode before exporting the images
                    uidoc.ActiveView.DisableTemporaryViewMode(TemporaryViewMode.TemporaryHideIsolate);

                    foreach (Reference item in selectedIssues)
                    {
                        Element issueElement = doc.GetElement(item);

                        string checkStatus = issueElement.LookupParameter("Status").AsValueString();

                        string checkPosition = issueElement.LookupParameter("Position Check").AsString();

                        Markup issue = new Markup(DateTime.Now);

                        Topic tp = new Topic();
                        tp.Title = issueElement.Name;
                        tp.Description = checkStatus;
                        issue.Topic = tp;



                        var view = new ViewPoint(issue.Viewpoints.Any());

                        if (!Directory.Exists(Path.Combine(bcfFile.TempPath, issue.Topic.Guid)))
                            Directory.CreateDirectory(Path.Combine(bcfFile.TempPath, issue.Topic.Guid));



                        var c = new Comment
                        {
                            Comment1 = checkPosition,
                            Author = Utils.GetUsername(),
                            //Status = comboStatuses.SelectedValue.ToString(),
                            //VerbalStatus = VerbalStatus.Text,
                            Date = DateTime.Now,
                            Viewpoint = new CommentViewpoint { Guid = view.Guid }
                        };
                        issue.Comment.Add(c);

                        // Zoom to the selected element
                        uidoc.ShowElements(item.ElementId);

                        SetViewZoomExtent(uiview, 1.2);

                        string path = Path.Combine(bcfFile.TempPath, issue.Topic.Guid, view.Snapshot);
                        // Export image
                        SaveRevitSnapshot(uidoc.Document, path);

                        view.SnapshotPath = path;

                        issue.Viewpoints.Add(view);

                        issue.Viewpoints.Last().VisInfo = Data.RevitView.GenerateViewpoint(uiapp.ActiveUIDocument);

                        if (uidoc.Document.Title != null)
                            issue.Header[0].Filename = uidoc.Document.Title;
                        else
                            issue.Header[0].Filename = "Unknown";


                        bcfFile.Issues.Add(issue);


                    }

                    t.Commit();
                
                 
                }

                bcfContainer.SaveFile(bcfFile);


                return Result.Succeeded;

      }
      catch (Exception e)
      {
        message = e.Message;
        return Result.Failed;
      }

    }
        private void SetViewZoomExtent(UIView uiview, double scaleFactor)
        {
            Rectangle rect = uiview.GetWindowRectangle();
            IList<XYZ> corners = uiview.GetZoomCorners();
            XYZ p = corners[0];
            XYZ q = corners[1];

            XYZ v = q - p;
            XYZ center = p + 0.5 * v;
            v *= scaleFactor;
            p = center - v;
            q = center + v;

            uiview.ZoomAndCenterRectangle(p, q);
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