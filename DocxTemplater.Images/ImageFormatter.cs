using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplater.Formatter;
using SixLabors.ImageSharp;
using System;
using System.Linq;
using DocxTemplater.ImageBase;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace DocxTemplater.Images
{
    public class ImageFormatter : IFormatter, IImageServiceProvider
    {
        public bool CanHandle(Type type, string prefix)
        {
            var prefixUpper = prefix.ToUpper();
            return prefixUpper is "IMAGE" or "IMG" && type == typeof(byte[]);
        }

        public void ApplyFormat(ITemplateProcessingContext templateContext, FormatterContext formatterContext, Text target)
        {
            // TODO: handle other ppi values than default 96
            // see https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.pixelsperinch?view=openxml-2.8.1#remarks
            if (formatterContext.Value is not byte[] imageBytes)
            {
                return;
            }
            if (imageBytes.Length == 0)
            {
                target.Text = string.Empty;
                return;
            }
            try
            {
                var root = target.GetRoot();
                var maxPropertyId = templateContext.ImageService.GetImage(root, imageBytes, out ImageInformation imageInfo);
                // Image ist a child element of a <wps:wsp> (TextBox)
                if (!TryHandleImageInWordprocessingShape(target, imageInfo, formatterContext.Args.FirstOrDefault() ?? string.Empty, maxPropertyId, templateContext.ImageService))
                {// Image is not a child element of a <wps:wsp> (TextBox) - rotation and scale is determined by the arguments
                    ImageUtils.AddInlineGraphicToRun(target, imageInfo, maxPropertyId, formatterContext.Args, templateContext.ImageService);
                }

            }
            catch (Exception e) when (e is InvalidImageContentException or UnknownImageFormatException)
            {
                throw new OpenXmlTemplateException("Could not detect image format", e);
            }
        }

        private static bool TryHandleImageInWordprocessingShape(Text target, ImageInformation imageInfo,
            string firstArgument, uint maxPropertyId, IImageService imageService)
        {
            var drawing = target.GetFirstAncestor<Drawing>();
            if (drawing == null)
            {
                return false;
            }

            // get extent of the drawing either from the anchor or inline
            var targetExtent = target.GetFirstAncestor<DW.Anchor>()?.GetFirstChild<DW.Extent>() ?? target.GetFirstAncestor<DW.Inline>()?.GetFirstChild<DW.Extent>();
            if (targetExtent != null)
            {
                double scale = 0;
                var imageCx = OpenXmlHelper.PixelsToEmu(imageInfo.PixelWidth);
                var imageCy = OpenXmlHelper.PixelsToEmu(imageInfo.PixelHeight);
                if (firstArgument.Equals("KEEPRATIO", StringComparison.CurrentCultureIgnoreCase))
                {
                    scale = Math.Min(targetExtent.Cx / (double)imageCx, targetExtent.Cy / (double)imageCy);
                }
                else if (firstArgument.Equals("STRETCHW", StringComparison.CurrentCultureIgnoreCase))
                {
                    scale = targetExtent.Cx / (double)imageCx;
                }
                else if (firstArgument.Equals("STRETCHH", StringComparison.CurrentCultureIgnoreCase))
                {
                    scale = targetExtent.Cy / (double)imageCy;
                }

                if (scale > 0)
                {
                    targetExtent.Cx = (long)(imageCx * scale);
                    targetExtent.Cy = (long)(imageCy * scale);
                }

                ReplaceAnchorContentWithPicture(imageInfo.ImagePartRelationId, maxPropertyId, drawing, imageInfo.ExifRotation, imageService);
            }

            target.Remove();
            return true;
        }

        private static void ReplaceAnchorContentWithPicture(string impagepartRelationShipId, uint maxPropertyId, Drawing original, ImageRotation imageInfoExifRotation, IImageService imageService)
        {
            var propertyId = maxPropertyId + 1;
            var inlineOrAnchor = (OpenXmlElement)original.GetFirstChild<DW.Anchor>() ??
                                 (OpenXmlElement)original.GetFirstChild<DW.Inline>();
            var originaleExtent = inlineOrAnchor.GetFirstChild<DW.Extent>();
            var transform = inlineOrAnchor.Descendants<A.Transform2D>().FirstOrDefault();
            var rotation = imageInfoExifRotation.AddUnits(transform?.Rotation ?? 0);
            var clonedInlineOrAnchor = inlineOrAnchor.CloneNode(false);

            if (inlineOrAnchor is DW.Anchor anchor)
            {
                clonedInlineOrAnchor.Append(new DW.SimplePosition { X = 0L, Y = 0L });
                var horzPosition = anchor.GetFirstChild<DW.HorizontalPosition>().CloneNode(true);
                var vertPosition = inlineOrAnchor.GetFirstChild<DW.VerticalPosition>().CloneNode(true);
                clonedInlineOrAnchor.Append(horzPosition);
                clonedInlineOrAnchor.Append(vertPosition);
                clonedInlineOrAnchor.Append(new DW.Extent { Cx = originaleExtent.Cx, Cy = originaleExtent.Cy });
                clonedInlineOrAnchor.Append(new DW.EffectExtent
                {
                    LeftEdge = 0L,
                    TopEdge = 0L,
                    RightEdge = 0L,
                    BottomEdge = 0L
                });
                clonedInlineOrAnchor.Append(new DW.WrapNone());
            }
            else if (inlineOrAnchor is DW.Inline)
            {
                clonedInlineOrAnchor.Append(new DW.Extent { Cx = originaleExtent.Cx, Cy = originaleExtent.Cy });
                clonedInlineOrAnchor.Append(new DW.EffectExtent
                {
                    LeftEdge = 0L,
                    TopEdge = 0L,
                    RightEdge = 0L,
                    BottomEdge = 0L
                });
            }

#pragma warning disable IDE0300
            clonedInlineOrAnchor.Append(new OpenXmlElement[]
            {
                new DW.DocProperties
                {
                    Id = propertyId,
                    Name = $"Picture {propertyId}"
                },
                new DW.NonVisualGraphicFrameDrawingProperties(
                    new A.GraphicFrameLocks {NoChangeAspect = true}),
                new A.Graphic(
                    new A.GraphicData(
                            imageService.CreatePicture(impagepartRelationShipId, propertyId, originaleExtent.Cx, originaleExtent.Cy, rotation)
                        )
                        {Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"})
            });
            var dw = new Drawing(clonedInlineOrAnchor);
            original.InsertAfterSelf(dw);
            original.Remove();
        }

        public IImageService CreateImageService()
        {
            return new ImageService();
        }
    }
}
