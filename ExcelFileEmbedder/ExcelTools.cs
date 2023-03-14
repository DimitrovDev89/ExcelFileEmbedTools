using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;
using OpenMcdf;
using System.Text;
using ExcelFileTools;
using System.Xml.Linq;
using DocumentFormat.OpenXml.VariantTypes;
using System.IO;

namespace ExcelFileEmbedder
{
    public class ExcelTools
    {
        private readonly EmbedFileOptions _options;

        public ExcelTools(EmbedFileOptions options)
        {
            _options = options;
        }
        public void EmbedFile()
        {
            var tempPath = $"{_options.ExcelFilePath}.tmp";
            File.Copy(_options.ExcelFilePath, tempPath, true);
            using SpreadsheetDocument package = SpreadsheetDocument.Open(tempPath,
                isEditable: true, new OpenSettings { AutoSave = false });
            EmbedFileInDocument(package);
            package.SaveAs(_options.OutputExcelPath);
        }

        private void EmbedFileInDocument(SpreadsheetDocument document)
        {
            var workbookPart1 = document.WorkbookPart;
            var worksheetPart1 = workbookPart1.WorksheetParts.First();
            GenerateWorksheetPart1Content(worksheetPart1);

            VmlDrawingPart vmlDrawingPart1 = worksheetPart1.AddNewPart<VmlDrawingPart>("rId3");
            GenerateVmlDrawingPart1Content(vmlDrawingPart1);

            ImagePart imagePart1 = vmlDrawingPart1.AddNewPart<ImagePart>("image/x-emf", "rId1");
            GenerateImagePart1Content(imagePart1);

            DrawingsPart drawingsPart1 = worksheetPart1.AddNewPart<DrawingsPart>("rId2");
            GenerateDrawingsPart1Content(drawingsPart1);

            SpreadsheetPrinterSettingsPart spreadsheetPrinterSettingsPart1 = worksheetPart1.AddNewPart<SpreadsheetPrinterSettingsPart>("rId1");
            GenerateSpreadsheetPrinterSettingsPart1Content(spreadsheetPrinterSettingsPart1);

            worksheetPart1.AddPart(imagePart1, "rId5");

            EmbeddedObjectPart embeddedObjectPart1 = worksheetPart1.AddNewPart<EmbeddedObjectPart>("application/vnd.openxmlformats-officedocument.oleObject", "rId4");
            GenerateEmbeddedObjectPart1Content(embeddedObjectPart1);

            SetPackageProperties(document);
        }
        private void SetPackageProperties(OpenXmlPackage document)
        {
            document.PackageProperties.Creator = "Ivan Dimitrov";
            document.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2023-03-10T17:03:51Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2023-03-10T17:09:24Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.LastModifiedBy = "Ivan Dimitrov";
        }

        private void GenerateEmbeddedObjectPart1Content(EmbeddedObjectPart embeddedObjectPart1)
        {
            System.IO.Stream dataStream = _options.EmbedFileStream;
            var dataSize = dataStream.Length;

            OpenMcdf.CompoundFile compoundFile = new OpenMcdf.CompoundFile();
            CFStream compObj = compoundFile.RootStorage.AddStream("\x0001CompObj");
            CFStream ole10Stream = compoundFile.RootStorage.AddStream("\x0001Ole10Native");
            compoundFile.RootStorage.CLSID = Guid.Parse("{0003000c-0000-0000-c000-000000000046}");

            using MemoryStream compObjStream = new MemoryStream();
            var binaryWriter = new BinaryWriter(compObjStream);

            //28 bytes, this is the CompObjHeader
            binaryWriter.Write(Enumerable.Repeat((byte)0, 28).ToArray());

            "OLE Package".WriteAsNullTerminatedAsciiWithLenPrefixString(binaryWriter);
            "Package".WriteAsNullTerminatedAsciiWithLenPrefixString(binaryWriter);
         
            // reserved field - empty
            binaryWriter.Write((UInt32)0);

            var bufferCompObj = new byte[compObjStream.Length];
            compObjStream.Seek(0, SeekOrigin.Begin);
            compObjStream.Read(bufferCompObj, 0, bufferCompObj.Length);
            compObj.SetData(bufferCompObj);

            // ole10stream

            using CompoundStream ole10MemStream = new CompoundStream(ole10Stream);
            var ole10BinaryWriter = new BinaryWriter(ole10MemStream);
            ole10BinaryWriter.Write(Enumerable.Repeat((byte)0, 4).ToArray());

            ole10BinaryWriter.Write((UInt16)0x0002);


            _options.FileName.WriteAsNullTerminatedAsciiString(ole10BinaryWriter);
            //temp path
            _options.FileName.WriteAsNullTerminatedAsciiString(ole10BinaryWriter);

            // Skip 2 unused bytes
            ole10BinaryWriter.Write((byte)0);
            ole10BinaryWriter.Write((byte)0);

            // Read format
            ole10BinaryWriter.Write((UInt16)0x00000003);


            // Read temporary path
            _options.FileName.WriteAsNullTerminatedAsciiWithLenPrefixString(ole10BinaryWriter);
            ole10BinaryWriter.Write((UInt32)dataSize);//check if correct

            var buffer = new byte[1024 * 1024];
            while(true)
            {
                var read = dataStream.Read(buffer, 0, buffer.Length);
                if(read == 0)
                {
                    break;
                }
                ole10BinaryWriter.Write(buffer, 0, read);
            }

            using MemoryStream myStream2 = new MemoryStream();
            compoundFile.Save(myStream2);
            compoundFile.Close();
            myStream2.Seek(0, SeekOrigin.Begin);
            embeddedObjectPart1.FeedData(myStream2);
        }


        private void GenerateWorksheetPart1Content(WorksheetPart worksheetPart1)
        {
            Worksheet worksheet1 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac xr xr2 xr3" } };
            worksheet1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet1.AddNamespaceDeclaration("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
            worksheet1.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            worksheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            worksheet1.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");
            worksheet1.AddNamespaceDeclaration("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2");
            worksheet1.AddNamespaceDeclaration("xr3", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3");
            worksheet1.SetAttribute(new OpenXmlAttribute("xr", "uid", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision", "{896C9882-1CD0-4A60-8BCB-B7818816E234}"));

            SheetViews sheetViews1 = new SheetViews();

            SheetView sheetView1 = new SheetView() { TabSelected = true, WorkbookViewId = (UInt32Value)0U };
            Selection selection1 = new Selection() { ActiveCell = "B4", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "B4" } };

            sheetView1.Append(selection1);

            sheetViews1.Append(sheetView1);
            SheetFormatProperties sheetFormatProperties1 = new SheetFormatProperties() { DefaultRowHeight = 14.5D, DyDescent = 0.35D };
            SheetData sheetData1 = new SheetData();
            PageMargins pageMargins1 = new PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            PageSetup pageSetup1 = new PageSetup()
            {
                Orientation = OrientationValues.Portrait,
                HorizontalDpi = (UInt32Value)4294967293U,
                VerticalDpi = (UInt32Value)0U,
                Id = "rId1"
            };
            Drawing drawing1 = new Drawing() { Id = "rId2" };
            LegacyDrawing legacyDrawing1 = new LegacyDrawing() { Id = "rId3" };

            OleObjects oleObjects1 = new OleObjects();

            AlternateContent alternateContent2 = new AlternateContent();
            alternateContent2.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

            AlternateContentChoice alternateContentChoice2 = new AlternateContentChoice() { Requires = "x14" };// slicer style add if not exists

            OleObject oleObject1 = new OleObject() { ProgId = "Package2", ShapeId = (UInt32Value)1025U, Id = "rId4" };

            EmbeddedObjectProperties embeddedObjectProperties1 = new EmbeddedObjectProperties() { DefaultSize = false, Id = "rId5" };

            ObjectAnchor objectAnchor1 = new ObjectAnchor() { MoveWithCells = true };

            FromMarker fromMarker1 = new FromMarker();
            Xdr.ColumnId columnId1 = new Xdr.ColumnId();
            columnId1.Text = _options.PositionFrom.Column;
            Xdr.ColumnOffset columnOffset1 = new Xdr.ColumnOffset();
            columnOffset1.Text = "0";
            Xdr.RowId rowId1 = new Xdr.RowId();
            rowId1.Text = _options.PositionFrom.Row;
            Xdr.RowOffset rowOffset1 = new Xdr.RowOffset();
            rowOffset1.Text = "0";

            fromMarker1.Append(columnId1);
            fromMarker1.Append(columnOffset1);
            fromMarker1.Append(rowId1);
            fromMarker1.Append(rowOffset1);

            ToMarker toMarker1 = new ToMarker();
            Xdr.ColumnId columnId2 = new Xdr.ColumnId();
            columnId2.Text = _options.PositionTo.Column;
            Xdr.ColumnOffset columnOffset2 = new Xdr.ColumnOffset();
            columnOffset2.Text = "431800";
            Xdr.RowId rowId2 = new Xdr.RowId();
            rowId2.Text = _options.PositionTo.Row;
            Xdr.RowOffset rowOffset2 = new Xdr.RowOffset();
            rowOffset2.Text = "158750";

            toMarker1.Append(columnId2);
            toMarker1.Append(columnOffset2);
            toMarker1.Append(rowId2);
            toMarker1.Append(rowOffset2);

            objectAnchor1.Append(fromMarker1);
            objectAnchor1.Append(toMarker1);

            embeddedObjectProperties1.Append(objectAnchor1);

            oleObject1.Append(embeddedObjectProperties1);

            alternateContentChoice2.Append(oleObject1);

            AlternateContentFallback alternateContentFallback1 = new AlternateContentFallback();
            OleObject oleObject2 = new OleObject() { ProgId = "Package2", ShapeId = (UInt32Value)1025U, Id = "rId4" };

            alternateContentFallback1.Append(oleObject2);

            alternateContent2.Append(alternateContentChoice2);
            alternateContent2.Append(alternateContentFallback1);

            oleObjects1.Append(alternateContent2);

            worksheet1.Append(sheetViews1);
            worksheet1.Append(sheetFormatProperties1);
            worksheet1.Append(sheetData1);
            worksheet1.Append(pageMargins1);
            worksheet1.Append(pageSetup1);
            worksheet1.Append(drawing1);
            worksheet1.Append(legacyDrawing1);
            worksheet1.Append(oleObjects1);

            worksheetPart1.Worksheet = worksheet1;
        }

        private void GenerateVmlDrawingPart1Content(VmlDrawingPart vmlDrawingPart1)
        {
            System.Xml.XmlTextWriter writer = new System.Xml.XmlTextWriter(vmlDrawingPart1.GetStream(System.IO.FileMode.Create), System.Text.Encoding.UTF8);
            writer.WriteRaw("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"\r\n xmlns:o=\"urn:schemas-microsoft-com:office:office\"\r\n xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\r\n <o:shapelayout v:ext=\"edit\">\r\n  <o:idmap v:ext=\"edit\" data=\"1\"/>\r\n </o:shapelayout><v:shapetype id=\"_x0000_t75\" coordsize=\"21600,21600\" o:spt=\"75\"\r\n  o:preferrelative=\"t\" path=\"m@4@5l@4@11@9@11@9@5xe\" filled=\"f\" stroked=\"f\">\r\n  <v:stroke joinstyle=\"miter\"/>\r\n  <v:formulas>\r\n   <v:f eqn=\"if lineDrawn pixelLineWidth 0\"/>\r\n   <v:f eqn=\"sum @0 1 0\"/>\r\n   <v:f eqn=\"sum 0 0 @1\"/>\r\n   <v:f eqn=\"prod @2 1 2\"/>\r\n   <v:f eqn=\"prod @3 21600 pixelWidth\"/>\r\n   <v:f eqn=\"prod @3 21600 pixelHeight\"/>\r\n   <v:f eqn=\"sum @0 0 1\"/>\r\n   <v:f eqn=\"prod @6 1 2\"/>\r\n   <v:f eqn=\"prod @7 21600 pixelWidth\"/>\r\n   <v:f eqn=\"sum @8 21600 0\"/>\r\n   <v:f eqn=\"prod @7 21600 pixelHeight\"/>\r\n   <v:f eqn=\"sum @10 21600 0\"/>\r\n  </v:formulas>\r\n  <v:path o:extrusionok=\"f\" gradientshapeok=\"t\" o:connecttype=\"rect\"/>\r\n  <o:lock v:ext=\"edit\" aspectratio=\"t\"/>\r\n </v:shapetype><v:shape id=\"Object_x0020_1\" o:spid=\"_x0000_s1025\" type=\"#_x0000_t75\"\r\n  style=\'position:absolute;margin-left:0;margin-top:0;width:34pt;height:41.5pt;\r\n  z-index:1;visibility:visible;mso-wrap-style:square\' filled=\"t\" fillcolor=\"window [65]\"\r\n  stroked=\"t\" strokecolor=\"windowText [64]\" o:insetmode=\"auto\">\r\n  <v:imagedata o:relid=\"rId1\" o:title=\"\"/>\r\n  <x:ClientData ObjectType=\"Pict\">\r\n   <x:SizeWithCells/>\r\n   <x:Anchor>\r\n    0, 0, 0, 0, 0, 68, 2, 25</x:Anchor>\r\n   <x:CF>Pict</x:CF>\r\n   <x:AutoPict/>\r\n  </x:ClientData>\r\n </v:shape></xml>");
            writer.Flush();
            writer.Close();
        }

        private void GenerateImagePart1Content(ImagePart imagePart1)
        {
            System.IO.Stream data = GetBinaryDataStream(imagePart1Data);
            imagePart1.FeedData(data);
            data.Close();
        }




        private string imagePart1Data = "AQAAAGwAAAAHAAAAAAAAADwAAABJAAAAAAAAAAAAAACSAwAAWwQAACBFTUYAAAEAcCcAAA0AAAACAAAAAAAAAAAAAAAAAAAAAAoAAEAGAABYAQAA1wAAAAAAAAAAAAAAAAAAAMA/BQDYRwMACgAAABAAAAAAAAAAAAAAAAkAAAAQAAAAQwAAAFIAAABSAAAAcAEAAAEAAADu////AAAAAAAAAAAAAAAAkAEAAAAAAAEAAAAAUwBlAGcAbwBlACAAVQBJAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABkdgAIAAAAACUAAAAMAAAAAQAAABgAAAAMAAAAAAAAABkAAAAMAAAA////AHIAAACgJAAACgAAAAAAAAA5AAAALwAAAAoAAAAAAAAAMAAAADAAAAAAgP8BAAAAAAAAAAAAAIA/AAAAAAAAAAAAAIA/AAAAAAAAAAD///8AAAAAAGwAAAA0AAAAoAAAAAAkAAAwAAAAMAAAACgAAAAwAAAAMAAAAAEAIAADAAAAACQAAAAAAAAAAAAAAAAAAAAAAAAAAP8AAP8AAP8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJGQj/+QkI7/j4+O/4+Ojf+Ojo3/jo2M/42Ni/+NjIv/jYyL/4yLiv+Li4n/i4qJ/4qKiP+KiYj/iYmH/4mIhv+IiIb/iIeG/4eGhf+HhoT/hoWE/4aFg/+FhIP/hYSC/4SDgv+Eg4H/g4KB/4OCgP+CgX//goF//4GAfv+BgH7/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJGRj//7+/r/+/v6//v7+v/7+/r/+/v6//v7+v/7+/r/+/v6//v7+v/7+/r/+/v6//v7+v/7+/r/+/v6//v7+v/7+/r/+/v6//v7+v/7+/r/+/v6//v7+v/7+/r/+/v6//v7+v/7+/r/+/v6//v7+v/7+/r/+/v6//v7+v+BgH7/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJKRkP/7+/r/9/b1//f29f/39vX/9/b1//f29f/39vX/9/b1//f29f/39vX/9/b1//f29f/39vX/9/b1//f29f/39vX/9/b1//f29f/39vX/9/b1//f29f/39vX/9/b1//f29f/39vX/9/b1//f29f/39vX/9/b1//v7+v+CgX//AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJKSkP/7+/r/9/b1//f29f/39vX/9/b1//f29f/39vX/9/b1//f29f/39vX/9/b1//f29f/39vX/9/b1//f29f/39vX/9/b1//f29f/39vX/9/b1//f29f/39vX/9/b1//f29f/39vX/9/b1//f29f/39vX/9/b1//v7+v+CgX//AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJOSkf/7+/v/9/b2/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/9/b1//v7+v+DgoD/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJOTkv/8+/v/+Pf2//j39v/49/b/+Pf2//j39v/39/b/9/b2//f29v/39vb/9/b1//f29f/39vX/9/b1//f29f/39vX/9/b1//f29f/39vX/9/b1//f29f/39vX/9/b1//f29f/39vX/9/b1//f29f/39vX/9/b1//v7+v+DgoH/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJSTkv/8+/v/+Pf2//j39v/49/b/+Pf2//j39v/49/b/+Pf2//j39v/49/b/9/f2//f29v/39vb/9/b1//f29f/39vX/9/b1//f29f/39vX/9/b1//f29f/39vX/9/b1//f29f/39vX/9/b1//f29f/39vX/9/b1//v7+v+Eg4H/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJSUk//8+/v/+Pf2/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/9/b1//v7+v+Eg4L/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJWUk//8+/v/+Pf2//j39v/49/b/+Pf2//j39v/49/b/+Pf2//j39v/49/b/+Pf2//j39v/49/b/+Pf2//j39v/39/b/9/b2//f29f/39vX/9/b1//f29f/39vX/9/b1//f29f/39vX/9/b1//f29f/39vX/9/b1//v7+v+FhIL/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJWVlP/8+/v/+Pf2//j39v/49/b/+Pf2//j39v/49/b/+Pf2//j39v/49/b/+Pf2//j39v/49/b/+Pf2//j39v/49/b/+Pf2//f39v/39vb/9/b1//f29f/39vX/9/b1//f29f/39vX/9/b1//f29f/39vX/9/b1//v7+v+FhIP/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJaVlP/8+/v/+Pf3/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/9/b1//v7+v+GhYP/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJaWlf/8/Pv/+Pj3//j49//4+Pf/+Pj3//j39//49/f/+Pf3//j39//49/b/+Pf2//j39v/49/b/+Pf2//j39v/49/b/+Pf2//j39v/49/b/+Pf2//f39v/39vb/9/b1//f29f/39vX/9/b1//f29f/39vX/9/b1//v7+v+GhYT/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJeWlf/8/Pv/+fj3//n49//5+Pf/+fj3//j49//4+Pf/+Pj3//j49//49/f/+Pf3//j39//49/b/+Pf2//j39v/49/b/+Pf2//j39v/49/b/+Pf2//j39v/49/b/9/b2//f29f/39vX/9/b1//f29f/39vX/9/b1//v7+v+HhoT/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJeXlv/8/Pv/+fj3/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/9/b1//v7+v+HhoX/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJiXl//8/Pv/+fj3//n49//5+Pf/+fj3//n49//5+Pf/+fj3//n49//5+Pf/+fj3//j49//4+Pf/+Pj3//j39//49/f/+Pf2//j39v/49/b/+Pf2//j39v/49/b/+Pf2//j39v/39vb/9/b1//f29f/39vX/9/b1//v7+v+Ih4b/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJiYl//8/Pv/+fj3//n49//5+Pf/+fj3//n49//5+Pf/+fj3//n49//5+Pf/+fj3//n49//5+Pf/+Pj3//j49//49/f/+Pf3//j39v/49/b/+Pf2//j39v/49/b/+Pf2//j39v/49/b/9/b2//f29f/39vX/9/b1//v7+v+IiIb/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJmYmP/8/Pz/+fj4/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/9/b1//v7+v+JiIb/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJmZmP/8/Pz/+fj4//n4+P/5+Pj/+fj4//n4+P/5+Pj/+fj4//n4+P/5+Pf/+fj3//n49//5+Pf/+fj3//n49//5+Pf/+Pj3//j49//49/f/+Pf3//j39v/49/b/+Pf2//j39v/49/b/+Pf2//j39v/39vX/9/b1//v7+v+JiYf/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJqZmf/8/Pz/+fn4//n5+P/5+fj/+fn4//n5+P/5+fj/+fj4//n4+P/5+Pj/+fj4//n49//5+Pf/+fj3//n49//5+Pf/+fj3//n49//4+Pf/+Pf3//j39//49/b/+Pf2//j39v/49/b/+Pf2//j39v/39/b/9/b1//v7+v+KiYj/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJqamf/8/Pv/+fj3/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/9/b2//v7+v+Kioj/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJubmv/9/Pz/+vn4//r5+P/6+fj/+vn4//r5+P/5+Pf/+fj3//n5+P/5+fj/+fn4//n4+P/5+Pj/+fj4//n49//5+Pf/+fj3//n49//5+Pf/+fj3//j49//49/f/+Pf2//j39v/49/b/+Pf2//j39v/49/b/+Pf2//v7+/+Lion/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJubmv/8/Pz/+fj4//n4+P/5+Pj/+vn4//r5+P/6+fj/+vn4//r5+P/5+Pf/+fn4//n5+P/5+fj/+fj4//n4+P/5+Pf/+fj3//n49//5+Pf/+fj3//n49//4+Pf/+Pf3//j39v/49/b/+Pf2//j39v/49/b/+Pf2//v7+/+Li4n/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJycm//9/Pz/+vn5/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/+Pf2//z7+/+Mi4r/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJycnP/9/Pz/+vn4//r5+P/6+fj/+vn4//r5+P/6+fn/+fn5//n4+P/6+fj/+vn4//r5+P/5+Pf/+fn4//n5+P/5+Pj/+fj4//n49//5+Pf/+fj3//n49//5+Pf/+Pj3//j39//49/f/+Pf2//j39v/49/b/+Pf2//z7+/+NjIv/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJ2dnP/9/fz/+vr5//r6+f/6+vn/+vn4//r5+P/6+fj/+vn4//r5+f/5+Pj/+fj4//r5+P/6+fj/+fj3//n5+P/5+fj/+fj4//n4+P/5+Pf/+fj3//n49//5+Pf/+fj3//j49//49/f/+Pf2//j39v/49/b/+Pf2//z7+/+NjIv/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJ2dnf/9/Pz/+vn5/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/+Pf2//z7+/+NjYv/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJ6enf/9/Pz/+vn5//r5+f/6+fn/+vn5//r5+f/6+fn/+vr5//r6+f/6+fj/+vn4//r5+f/5+Pj/+vn4//r5+P/5+Pf/+fn4//n5+P/5+Pj/+fj4//n49//5+Pf/+fj3//n49//4+Pf/+Pf3//j39v/49/b/+Pf2//z7+/+OjYz/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJ6env/9/fz/+/r5//v6+f/7+vn/+vn5//r5+f/6+fn/+vn5//r5+f/6+vn/+vn4//r5+P/6+fn/+fj4//r5+P/6+fj/+fn4//n5+P/5+Pj/+fj4//n49//5+Pf/+fj3//n49//4+Pf/+Pf3//j39//49/b/+Pf2//z7+/+Ojo3/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJ+fnv/9/fz/+vr5/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/ra2t/62trf+tra3/+Pf2//z7+/+Pjo3/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJ+fn//9/fz/+vr5//r6+f/6+vn/+vr5//r6+f/6+vn/+/r5//r5+f/6+fn/+vn5//r6+f/6+fj/+vn4//n4+P/6+fj/+vn4//n49//5+fj/+fj4//n4+P/5+Pf/+fj3//n49//5+Pf/+Pj3//j39//49/b/+Pf2//z7+/+Pj47/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAKCgoP/9/f3/+vr6//r6+v/6+vn/+vr5//r6+f/6+vn/+vr5//v6+f/6+fn/+vn5//r5+f/6+vn/+vn4//r5+f/5+Pj/+vn4//r5+P/5+fj/+fn4//n4+P/5+Pj/+fj3//n49//5+Pf/+Pj3//j49//49/f/+Pf2//z7+/+QkI7/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";

        private string spreadsheetPrinterSettingsPart1Data = "QwBhAG4AbwBuACAAVABTADUAMAAwADAAIABzAGUAcgBpAGUAcwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEECgzcANQOA9+BAwEAAQDqCm8IZAABAAcA/f8CAAEAAAABAAEATABlAHQAdABlAHIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAAAAAgAAAAEAAAABAAAAAAAAAAAAAAAAAAAAAAAAANQOAABCSkRNCgwAAAAAAACQCgAA2QAAANkAAAAAAAAAAAAAAAEAAABWVAAAJG0AACwBAACAAgAAYE8AAARqAAAsAQAAgAIAAGBPAAAEagAAVlQAACRtAAAsAQAAgAIAAHYCAAD0AQAAYE8AAARqAAAsAQAAgAIAAHYCAAD0AQAALAEAAIACAAB2AgAA9AEAAGBPAAAEagAAWAJYAhgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAnAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAAAADAAAAAwAAAAAAAAACAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAAwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAAEAAAADAAAAAwAAAAIAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAAEAAAADAAAABwAAAAMAAAAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAwAAAAIAAAABAAAAAAAAAAEAAAAAAAAAAAAAAGQAAAABAAAAVlQAACRtAAABAAAAAQAAAFZUAAAkbQAAAQAAAAIAAAAAAAAAAQAAAAEAAAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJMAAAAAAAAAAAAAAEAKAAABAAAAAQAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEAAAD//wAAAAAAAAAAAAAAAAAACgAAAAAAAAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAAAAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAAAAAAAAAP//AAAAAAAAAAAAAAAAAAACAAAAAAAAAAAAAAAAAAAAAQAAAAAAAAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABWVAAAJG0AAAAAAAABAAAAfwAAAH8AAAB/AAAAfwAAAAAAAAAAAAAAAAAAAAAAAADnAwAA/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAAKAAAAAAAAAAAAAAAAAAAAAAAAAOcDAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAUAAAABAAAAAAAAAAIAAAAAAAAAAgAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAOgDAAAAAAAAAQAAAAAAAAABAAAAAAAAAAAAAAACAAAAAAAAAAAAAAAAAAAAAQAAAAAAAAAEAAAAAAAAAAAAAAABAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAAAAAQAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAEAAAAZAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQwBhAG4AbwBuACAAVABTADUAMAAwADAAIABzAGUAcgBpAGUAcwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEECgzcANQOA9+BAwEAAQDqCm8IZAABAAcA/f8CAAEAAAABAAEATABlAHQAdABlAHIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAAAAAgAAAAEAAAABAAAAAAAAAAAAAAAAAAAAAAAAAD0zvkg=";

        private System.IO.Stream GetBinaryDataStream(string base64String)
        {
            return new System.IO.MemoryStream(System.Convert.FromBase64String(base64String));
        }

        private void GenerateDrawingsPart1Content(DrawingsPart drawingsPart1)
        {
            Xdr.WorksheetDrawing worksheetDrawing1 = new Xdr.WorksheetDrawing();
            worksheetDrawing1.AddNamespaceDeclaration("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
            worksheetDrawing1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            AlternateContent alternateContent3 = new AlternateContent();
            alternateContent3.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

            AlternateContentChoice alternateContentChoice3 = new AlternateContentChoice() { Requires = "a14" };
            alternateContentChoice3.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            Xdr.TwoCellAnchor twoCellAnchor1 = new Xdr.TwoCellAnchor() { EditAs = Xdr.EditAsValues.OneCell };

            Xdr.FromMarker fromMarker2 = new Xdr.FromMarker();
            Xdr.ColumnId columnId3 = new Xdr.ColumnId();
            columnId3.Text = _options.PositionFrom.Column;
            Xdr.ColumnOffset columnOffset3 = new Xdr.ColumnOffset();
            columnOffset3.Text = "0";
            Xdr.RowId rowId3 = new Xdr.RowId();
            rowId3.Text = _options.PositionFrom.Row;
            Xdr.RowOffset rowOffset3 = new Xdr.RowOffset();
            rowOffset3.Text = "0";

            fromMarker2.Append(columnId3);
            fromMarker2.Append(columnOffset3);
            fromMarker2.Append(rowId3);
            fromMarker2.Append(rowOffset3);

            Xdr.ToMarker toMarker2 = new Xdr.ToMarker();
            Xdr.ColumnId columnId4 = new Xdr.ColumnId();
            columnId4.Text = _options.PositionTo.Column;
            Xdr.ColumnOffset columnOffset4 = new Xdr.ColumnOffset();
            columnOffset4.Text = "431800";
            Xdr.RowId rowId4 = new Xdr.RowId();
            rowId4.Text = _options.PositionTo.Row;
            Xdr.RowOffset rowOffset4 = new Xdr.RowOffset();
            rowOffset4.Text = "158750";

            toMarker2.Append(columnId4);
            toMarker2.Append(columnOffset4);
            toMarker2.Append(rowId4);
            toMarker2.Append(rowOffset4);

            Xdr.Shape shape1 = new Xdr.Shape() { Macro = "", TextLink = "" };

            Xdr.NonVisualShapeProperties nonVisualShapeProperties1 = new Xdr.NonVisualShapeProperties();

            Xdr.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Xdr.NonVisualDrawingProperties() { Id = (UInt32Value)1025U, Name = "Object 1", Hidden = true };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList1 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension1 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{63B3BB69-23CF-44E3-9099-C40C66FF867C}" };
            A14.CompatExtension compatExtension1 = new A14.CompatExtension() { ShapeId = "_x0000_s1025" };

            nonVisualDrawingPropertiesExtension1.Append(compatExtension1);

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension2 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement3 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{8FDDA969-3543-ED29-9D74-16B006238762}\" />");

            nonVisualDrawingPropertiesExtension2.Append(openXmlUnknownElement3);

            nonVisualDrawingPropertiesExtensionList1.Append(nonVisualDrawingPropertiesExtension1);
            nonVisualDrawingPropertiesExtensionList1.Append(nonVisualDrawingPropertiesExtension2);

            nonVisualDrawingProperties1.Append(nonVisualDrawingPropertiesExtensionList1);
            Xdr.NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties1 = new Xdr.NonVisualShapeDrawingProperties();

            nonVisualShapeProperties1.Append(nonVisualDrawingProperties1);
            nonVisualShapeProperties1.Append(nonVisualShapeDrawingProperties1);

            Xdr.ShapeProperties shapeProperties1 = new Xdr.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents1 = new A.Extents() { Cx = 0L, Cy = 0L };

            transform2D1.Append(offset1);
            transform2D1.Append(extents1);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            A.SolidFill solidFill7 = new A.SolidFill();

            A.RgbColorModelHex rgbColorModelHex12 = new A.RgbColorModelHex() { Val = "FFFFFF", LegacySpreadsheetColorIndex = 65, MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "a14" } };

            solidFill7.Append(rgbColorModelHex12);

            A.Outline outline4 = new A.Outline() { Width = 9525 };

            A.SolidFill solidFill8 = new A.SolidFill();

            A.RgbColorModelHex rgbColorModelHex13 = new A.RgbColorModelHex() { Val = "000000", LegacySpreadsheetColorIndex = 64, MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "a14" } };

            solidFill8.Append(rgbColorModelHex13);
            A.PresetDash presetDash4 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter4 = new A.Miter() { Limit = 800000 };
            A.HeadEnd headEnd1 = new A.HeadEnd();
            A.TailEnd tailEnd1 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            outline4.Append(solidFill8);
            outline4.Append(presetDash4);
            outline4.Append(miter4);
            outline4.Append(headEnd1);
            outline4.Append(tailEnd1);
            A.EffectList effectList4 = new A.EffectList();

            A.ShapePropertiesExtensionList shapePropertiesExtensionList1 = new A.ShapePropertiesExtensionList();

            A.ShapePropertiesExtension shapePropertiesExtension1 = new A.ShapePropertiesExtension() { Uri = "{AF507438-7753-43E0-B8FC-AC1667EBCBE1}" };

            A14.HiddenEffectsProperties hiddenEffectsProperties1 = new A14.HiddenEffectsProperties();

            A.EffectList effectList5 = new A.EffectList();

            A.OuterShadow outerShadow2 = new A.OuterShadow() { Distance = 35921L, Direction = 2700000, Alignment = A.RectangleAlignmentValues.Center, RotateWithShape = false };
            A.RgbColorModelHex rgbColorModelHex14 = new A.RgbColorModelHex() { Val = "808080" };

            outerShadow2.Append(rgbColorModelHex14);

            effectList5.Append(outerShadow2);

            hiddenEffectsProperties1.Append(effectList5);

            shapePropertiesExtension1.Append(hiddenEffectsProperties1);

            shapePropertiesExtensionList1.Append(shapePropertiesExtension1);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);
            shapeProperties1.Append(solidFill7);
            shapeProperties1.Append(outline4);
            shapeProperties1.Append(effectList4);
            shapeProperties1.Append(shapePropertiesExtensionList1);

            shape1.Append(nonVisualShapeProperties1);
            shape1.Append(shapeProperties1);
            Xdr.ClientData clientData1 = new Xdr.ClientData();

            twoCellAnchor1.Append(fromMarker2);
            twoCellAnchor1.Append(toMarker2);
            twoCellAnchor1.Append(shape1);
            twoCellAnchor1.Append(clientData1);

            alternateContentChoice3.Append(twoCellAnchor1);
            AlternateContentFallback alternateContentFallback2 = new AlternateContentFallback();

            alternateContent3.Append(alternateContentChoice3);
            alternateContent3.Append(alternateContentFallback2);

            worksheetDrawing1.Append(alternateContent3);

            drawingsPart1.WorksheetDrawing = worksheetDrawing1;
        }

        private void GenerateSpreadsheetPrinterSettingsPart1Content(SpreadsheetPrinterSettingsPart spreadsheetPrinterSettingsPart1)
        {
            System.IO.Stream data = GetBinaryDataStream(spreadsheetPrinterSettingsPart1Data);
            spreadsheetPrinterSettingsPart1.FeedData(data);
            data.Close();
        }
    }
}