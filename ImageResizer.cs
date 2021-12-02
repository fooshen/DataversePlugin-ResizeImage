using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System.IO;
using System.Drawing;
using Microsoft.Crm.Sdk.Messages;

namespace DVDemoPlugins
{
    public class ImageResizer : IPlugin
    {        
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            ITracingService tracingSvc = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            
            if(context.InputParameters.Contains("Target"))
            {
                if(context.InputParameters["Target"] is Entity)
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    if(entity.LogicalName == "ms_demoimagequeue")
                    {
                        tracingSvc.Trace("Demo Image Queue entity updated");
                        Guid id = Guid.Parse(entity.Attributes["ms_demoimagequeueid"].ToString());
   
                        IOrganizationServiceFactory svcFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                        IOrganizationService svc = svcFactory.CreateOrganizationService(context.UserId);
                       
                        //get image information
                        Entity imageQueueRecord = svc.Retrieve("ms_demoimagequeue", id, new ColumnSet("ms_image"));

                        //checks if image is uploaded
                        if (imageQueueRecord.Contains("ms_image"))
                        {
                            //get the full image file
                            InitializeFileBlocksDownloadRequest fileRequest = new InitializeFileBlocksDownloadRequest();
                            fileRequest.Target = new EntityReference("ms_demoimagequeue", id);
                            fileRequest.FileAttributeName = "ms_image";
                            InitializeFileBlocksDownloadResponse fileResponse = (InitializeFileBlocksDownloadResponse)svc.Execute(fileRequest);
                            DownloadBlockRequest imageRequest = new DownloadBlockRequest();
                            imageRequest.FileContinuationToken = fileResponse.FileContinuationToken;
                            DownloadBlockResponse imageResponse = (DownloadBlockResponse)svc.Execute(imageRequest);
                            byte[] imageBytes = imageResponse.Data;
                            if (imageBytes.Length > 0)
                            {
                                using (MemoryStream ms = new MemoryStream(imageBytes))
                                {
                                    Image originalImage = Image.FromStream(ms);

                                    int originalHeight = originalImage.Height;
                                    int originalWidth = originalImage.Width;

                                    float resizeRatio = 0;
                                    if (originalHeight > 500)
                                    {
                                        resizeRatio = (float)(originalHeight / 500f);
                                    }
                                    else if (originalWidth > 500)
                                    {
                                        resizeRatio = (float)(originalWidth / 500f);
                                    }
                                    if (resizeRatio > 0)
                                    {
                                        int newHeight = (int)(originalHeight / resizeRatio);
                                        int newWidth = (int)(originalWidth / resizeRatio);
                                        Bitmap newImage = new Bitmap(newWidth, newHeight);
                                        using (Graphics g = Graphics.FromImage((Image)newImage))
                                        {
                                            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                            g.DrawImage(originalImage, 0, 0, newWidth, newHeight);
                                        }

                                        using (MemoryStream tempMs = new MemoryStream())
                                        {
                                            newImage.Save(tempMs, originalImage.RawFormat);
                                            byte[] newImageContent = tempMs.ToArray();
                                            Entity resizedImageRecord = new Entity("ms_demoimage");
                                            resizedImageRecord["ms_demoimageid"] = svc.Create(resizedImageRecord);
                                            resizedImageRecord.Attributes.Add("ms_originalimage", imageBytes);
                                            resizedImageRecord.Attributes.Add("ms_resizedimage", newImageContent);
                                            resizedImageRecord.Attributes.Add("ms_originalimageheight", originalHeight);
                                            resizedImageRecord.Attributes.Add("ms_originalimagewidth", originalWidth);
                                            resizedImageRecord.Attributes.Add("ms_resizedimageheight", newHeight);
                                            resizedImageRecord.Attributes.Add("ms_resizedimagewidth", newWidth);
                                            svc.Update(resizedImageRecord);                                            
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

        }
    }
}
