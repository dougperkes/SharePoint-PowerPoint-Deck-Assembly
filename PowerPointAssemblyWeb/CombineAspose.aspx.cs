﻿using Aspose.Slides;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace PowerPointAssemblyWeb
{
    public partial class CombineAspose : System.Web.UI.Page
    {
        protected void Page_PreInit(object sender, EventArgs e)
        {
            Uri redirectUrl;
            switch (SharePointContextProvider.CheckRedirectionStatus(Context, out redirectUrl))
            {
                case RedirectionStatus.Ok:
                    return;
                case RedirectionStatus.ShouldRedirect:
                    Response.Redirect(redirectUrl.AbsoluteUri, endResponse: true);
                    break;
                case RedirectionStatus.CanNotRedirect:
                    Response.Write("An error occurred while processing your request.");
                    Response.End();
                    break;
            }
        }

        private List<ShptFile> files = new List<ShptFile>();
        private List<string> order = new List<string>();

        protected void Page_Load(object sender, EventArgs e)
        {
            //only bind first time in
            if (!this.IsPostBack)
            {
                //Get the url parameters for SPListId and SPListItemId
                var listID = new Guid(Request["SPListId"]);
                var listItemIDs = Request["SPListItemId"].Split(',');

                // The following code gets the client context and Title property by using TokenHelper.
                // To access other properties, the app may need to request permissions on the host web.
                var spContext = SharePointContextProvider.Current.GetSharePointContext(Context);
                using (var clientContext = spContext.CreateUserClientContextForSPHost())
                {
                    if (listItemIDs.Count() > 0 && listItemIDs[0] != "null")
                    {
                        var list = clientContext.Web.Lists.GetById(listID);
                        clientContext.Load(list);
                        clientContext.ExecuteQuery();

                        //get each item
                        int index = 0;
                        foreach (var id in listItemIDs)
                        {
                            var file = list.GetItemById(Convert.ToInt32(id)).File;
                            clientContext.Load(file);
                            clientContext.ExecuteQuery();

                            //only add presentations
                            if (file.Name.EndsWith(".pptx", StringComparison.CurrentCultureIgnoreCase))
                            {
                                files.Add(new ShptFile()
                                {
                                    Id = Convert.ToInt32(id),
                                    Name = file.Name,
                                    Path = file.ServerRelativeUrl,
                                    Index = index++
                                });
                            }
                        }

                        //get presentation count for ordering
                        for (int i = 1; i <= files.Count; i++)
                            order.Add(i.ToString());
                    }
                    else
                    {

                        //btnOk.Enabled = false;
                        //txtFileName.Enabled = false;
                        var list = clientContext.Web.Lists.GetById(listID);

                        var allFiles = list.RootFolder.Files;
                        clientContext.Load(allFiles,
                        files => files.Include(file => file.ListItemAllFields, file => file.Name, file => file.ServerRelativeUrl)
                        .Take(2000));
                        clientContext.ExecuteQuery();


                        var allIds = "";
                        var tempFiles = new List<ShptFile>();
                        foreach (var file in allFiles)
                        {
                            if (file.Name.EndsWith(".pptx", StringComparison.CurrentCultureIgnoreCase)
                                && file.ListItemAllFields["Order0"] != null)
                            {
                                var order = file.ListItemAllFields["Order0"].ToString();
                                int idx;
                                if (int.TryParse(order, out idx))
                                {
                                    tempFiles.Add(new ShptFile()
                                    {
                                        Id = file.ListItemAllFields.Id,
                                        Name = file.Name,
                                        Path = file.ServerRelativeUrl,
                                        Index = idx
                                    });
                                }
                            }
                        }
                        tempFiles = tempFiles.OrderBy(f => f.Index).ToList();
                        for (var index = 0; index < tempFiles.Count; index++)
                        {
                            var file = tempFiles[index];
                            file.Index = index + 1;
                            allIds += file.Id + ",";
                            files.Add(file);
                            order.Add((index + 1).ToString());
                        }

                        allIdsHidden.Value = allIds;


                    }

                    //bind the files to the UI
                    gridViewSelectedFiles.DataSource = files;
                    gridViewSelectedFiles.DataBind();

                }
            }
        }

        protected void btnOk_Click(object sender, EventArgs e)
        {
            //Get the url parameters for SPListId and SPListItemId
            var listID = new Guid(Request["SPListId"]);

            String[] listItemIDs;
            if (Request["SPListItemId"] != null && Request["SPListItemId"] != "null")
            {
                listItemIDs = Request["SPListItemId"].Split(',');
            }
            else
            {
                listItemIDs = allIdsHidden.Value.Split(',');
            }

            // The following code gets the client context and Title property by using TokenHelper.
            // To access other properties, the app may need to request permissions on the host web.
            var spContext = SharePointContextProvider.Current.GetSharePointContext(Context);
            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                var list = clientContext.Web.Lists.GetById(listID);
                clientContext.Load(list, l => l.RootFolder, l => l.RootFolder.ServerRelativeUrl);
                clientContext.ExecuteQuery();

                //setup sources for the combined file
                //List<OrderedSlideSource> sources = new List<OrderedSlideSource>();
                var files = new List<ShptFile>();

                //check if row is selected
                foreach (GridViewRow rowItem in gridViewSelectedFiles.Rows)
                {
                    CheckBox chk = (CheckBox)rowItem.Cells[0].FindControl("chkItem");
                    CheckBox chkFormat = (CheckBox)rowItem.Cells[0].FindControl("chkFormat");
                    DropDownList cbo = (DropDownList)rowItem.Cells[2].FindControl("cboOrder");
                    if (chk.Checked)
                    {
                        files.Add(new ShptFile()
                        {
                            Id = Convert.ToInt32(listItemIDs[rowItem.RowIndex]),
                            Index = Convert.ToInt32(cbo.SelectedValue)
                        });
                    }

                }

                if (files.Count > 1)
                {
                    //reorder
                    using (Presentation combined = new Presentation())
                    {
                        
                        foreach (var ss in files.OrderBy(i => i.Index))
                        {
                            //read the binary and prepare to combine
                            var file = list.GetItemById(ss.Id).File;
                            clientContext.Load(file);
                            clientContext.ExecuteQuery();

                            var stream = file.OpenBinaryStream();
                            clientContext.ExecuteQuery();
                            using (var ms = stream.Value)
                            {
                                try {
                                    using (Presentation sourceSlide = new Presentation(ms))
                                    {
                                        combined.SlideSize.Type = sourceSlide.SlideSize.Type;
                                        combined.SlideSize.Size = sourceSlide.SlideSize.Size;

                                        ISlide slide = sourceSlide.Slides[0];
                                        var bitmap = slide.GetThumbnail(1.0f, 1.0f); // new System.Drawing.Size(16680, 11760));
                                        bitmap.Save(Server.MapPath("~/App_Data/Slide_" + ss.Index + "_0.png"), System.Drawing.Imaging.ImageFormat.Png);
                                        combined.Slides.AddClone(slide);
                                    }
                                } catch (Exception ex)
                                {
                                    //TODO: better handle exceptions here. Not sure why we are getting Out of memory exceptions.
                                    System.Diagnostics.Debug.WriteLine(ex.ToString());
                                }
                            }
                        }

                        using (var combinedStream = new MemoryStream())
                        {
                            //combined.Save(combinedStream, Aspose.Slides.Export.SaveFormat.Pptx);
                            combined.Save(Server.MapPath("~/App_Data/temp.pptx"), Aspose.Slides.Export.SaveFormat.Pptx);

                            string fileName = ((txtFileName.Text.EndsWith(".pptx", StringComparison.CurrentCultureIgnoreCase)) ? txtFileName.Text : txtFileName.Text + ".pptx");
                            //list.RootFolder.UploadFile(fileName, combinedStream, true);
                            list.RootFolder.UploadFile(fileName, Server.MapPath("~/App_Data/temp.pptx"), true);
                        }

                    }
                }

                //add script to page to close dialog and refresh page
                ScriptManager.RegisterClientScriptBlock(this, typeof(Combine), "closeDialog", "closeParentDialog(true);", true);
            }
        }
        protected void gridViewSelectedFiles_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            //ignore hearder
            if (e.Row.RowIndex != -1)
            {
                //find and bind the order dropdown
                DropDownList cboOrder = (DropDownList)e.Row.FindControl("cboOrder");
                var o = (e.Row.RowIndex + 1).ToString();
                cboOrder.Attributes.Add("data-prev", o);
                cboOrder.DataSource = order;
                cboOrder.DataBind();
                cboOrder.SelectedValue = o;
            }
        }
    }

    public class ShptFileAspose
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Path { get; set; }
        public int Index { get; set; }
    }

}