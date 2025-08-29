using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplater.Formatter;
using SixLabors.ImageSharp;
using System;
using System.Text.RegularExpressions;
using DocxTemplater.ImageBase;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using System.Linq;

namespace DocxTemplater
{
    public static class ImageUtils
    {
        private static readonly Regex ArgumentRegex = new(@"(?<key>[whr]):(?<value>\d+)(?<unit>px|cm|in|pt|mm)?", RegexOptions.Compiled, TimeSpan.FromMilliseconds(500));

        public static Drawing CreateDrawing(ImageInformation imageInfo, uint maxDocumentPropertyId, string[] arguments, IImageService imageService)
        {
            var propertyId = maxDocumentPropertyId + 1;

            TransformSize(imageInfo.PixelWidth, imageInfo.PixelHeight, arguments, out var cx, out var cy, out var rotation);
            rotation = rotation.AddUnits(imageInfo.ExifRotation.Units);

            // Define the reference of the image.
            var drawing =
                new Drawing(
                    new DW.Inline(
                        new DW.Extent { Cx = cx, Cy = cy },
                        new DW.EffectExtent
                        {
                            LeftEdge = 0L,
                            TopEdge = 0L,
                            RightEdge = 0L,
                            BottomEdge = 0L
                        },
                        new DW.DocProperties
                        {
                            Id = propertyId,
                            Name = $"Picture {propertyId}"
                        },
                        new DW.NonVisualGraphicFrameDrawingProperties(
                            new A.GraphicFrameLocks { NoChangeAspect = true }),
                        new A.Graphic(
                            new A.GraphicData(
                                    imageService.CreatePicture(imageInfo.ImagePartRelationId, propertyId, cx, cy, rotation)
                                )
                            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                    )
                    {
                        DistanceFromTop = (UInt32Value)0U,
                        DistanceFromBottom = (UInt32Value)0U,
                        DistanceFromLeft = (UInt32Value)0U,
                        DistanceFromRight = (UInt32Value)0U
                    });
            return drawing;
        }

        public static void AddInlineGraphicToRun(OpenXmlElement target, ImageInformation imageInfo, uint maxDocumentPropertyId, string[] arguments, IImageService imageService)
        {
            var drawing = CreateDrawing(imageInfo, maxDocumentPropertyId, arguments, imageService);
            target.InsertAfterSelf(drawing);
            target.Remove();
        }

        public static void TransformSize(int pixelWidth, int pixelHeight, string[] arguments, out int outCxEmu, out int outCyEmu, out ImageRotation rotationInDegree)
        {
            var cxEmu = -1;
            var cyEmu = -1;
            rotationInDegree = ImageRotation.CreateFromDegree(0);

            if (arguments == null || arguments.Length == 0)
            {
                outCxEmu = OpenXmlHelper.PixelsToEmu(pixelWidth);
                outCyEmu = OpenXmlHelper.PixelsToEmu(pixelHeight);
                return;
            }

            foreach (var argument in arguments)
            {
                try
                {
                    var matches = ArgumentRegex.Matches(argument);
                    if (matches.Count == 0)
                    {
                        outCxEmu = OpenXmlHelper.PixelsToEmu(pixelWidth);
                        outCyEmu = OpenXmlHelper.PixelsToEmu(pixelHeight);
                        return;
                    }

                    foreach (Match match in matches)
                    {
                        var key = match.Groups["key"].Value;
                        var value = int.Parse(match.Groups["value"].Value);
                        var unit = match.Groups["unit"].Value;
                        switch (key)
                        {
                            case "w":
                                cxEmu = OpenXmlHelper.LengthToEmu(value, unit);
                                break;
                            case "h":
                                cyEmu = OpenXmlHelper.LengthToEmu(value, unit);
                                break;
                            case "r":
                                rotationInDegree = ImageRotation.CreateFromDegree(value);
                                break;
                        }
                    }
                }
                catch (RegexMatchTimeoutException)
                {
                    throw new OpenXmlTemplateException($"Invalid image formatter argument '{argument}'");
                }
            }

            if (cxEmu == -1 && cyEmu == -1)
            {
                outCxEmu = OpenXmlHelper.PixelsToEmu(pixelWidth);
                outCyEmu = OpenXmlHelper.PixelsToEmu(pixelHeight);
                return;
            }

            if (cxEmu == -1)
            {
                cxEmu = (int)(cyEmu * ((double)pixelWidth / pixelHeight));
            }
            else if (cyEmu == -1)
            {
                cyEmu = (int)(cxEmu * ((double)pixelHeight / pixelWidth));
            }
            else
            {
                // if both are set, the aspect ratio is kept
                var aspectRatio = (double)pixelWidth / pixelHeight;
                var newAspectRatio = (double)cxEmu / cyEmu;
                if (aspectRatio > newAspectRatio)
                {
                    cyEmu = (int)(cxEmu / aspectRatio);
                }
                else
                {
                    cxEmu = (int)(cyEmu * aspectRatio);
                }
            }
            outCxEmu = cxEmu;
            outCyEmu = cyEmu;
        }
    }
}
