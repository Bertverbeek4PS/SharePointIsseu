dotnet
{
    //Base and Third Party
    assembly("mscorlib")
    {
        type("System.Collections.Generic.Dictionary`2"; "Dictionary_Of_T_U")
        {
        }

        type("Microsoft.Win32.RegistryKey"; "RegistryKey")
        {
        }

        type("Microsoft.Win32.RegistryHive"; "RegistryHive")
        {
        }

        type("Microsoft.Win32.RegistryView"; "RegistryView")
        {
        }

        type("System.Reflection.AssemblyName"; "AssemblyName")
        {
        }

        type("System.Collections.Generic.List`1"; "List_Of_T")
        {
        }

        type("System.Reflection.MethodInfo"; "MethodInfo")
        {
        }

        type("System.Security.Cryptography.HMACSHA256"; "HMACSHA256")
        {
        }

        type("System.Environment+SpecialFolder"; "Environment_SpecialFolder")
        {
        }

        type("System.Security.Cryptography.SHA1CryptoServiceProvider"; "SHA1CryptoServiceProvider")
        {
        }

        type("System.Security.Cryptography.SHA256Managed"; "SHA256Managed")
        {
        }

        type("System.Security.Cryptography.MD5CryptoServiceProvider"; "MD5CryptoServiceProvider")
        {
        }

        type("System.Security.Cryptography.Rfc2898DeriveBytes"; "Rfc2898DeriveBytes")
        {
        }

        type("System.Security.Cryptography.HashAlgorithmName"; "HashAlgorithmName")
        {
        }

        type("System.Security.Cryptography.DeriveBytes"; "DeriveBytes")
        {
        }

        type("System.Reflection.ParameterInfo"; "ParameterInfo")
        {
        }

        type("System.TimeZoneInfo+AdjustmentRule"; "TimeZoneInfo_AdjustmentRule")
        {
        }

        type("System.Runtime.Serialization.Formatters.Binary.BinaryFormatter"; "BinaryFormatter")
        {
        }

        type("System.Collections.Generic.IEnumerable`1"; "IEnumerable_Of_T")
        {
        }

        type("System.Collections.Generic.IEnumerator`1"; "IEnumerator_Of_T")
        {
        }

        type("System.IO.FileAccess"; "FileAccess")
        {
        }

        type("System.Reflection.Assembly"; "Assembly")
        {
        }

        type("System.Security.Cryptography.X509Certificates.X509Certificate"; "X509Certificate")
        {
        }

        type("System.Security.Cryptography.SHA1Managed"; "SHA1Managed")
        {
        }

        type("System.Security.Cryptography.RijndaelManaged"; "RijndaelManaged")
        {
        }

        type("System.Security.Cryptography.SymmetricAlgorithm"; "SymmetricAlgorithm")
        {
        }

        type("System.Security.Cryptography.Aes"; "Aes")
        {
        }

        type("System.Security.Cryptography.DESCryptoServiceProvider"; "DESCryptoServiceProvider")
        {
        }

        type("System.Security.Cryptography.RC2CryptoServiceProvider"; "RC2CryptoServiceProvider")
        {
        }

        type("System.Security.Cryptography.TripleDESCryptoServiceProvider"; "TripleDESCryptoServiceProvider")
        {
        }

        type("System.Security.Cryptography.ICryptoTransform"; "ICryptoTransform")
        {
        }

        type("System.Collections.Generic.IList`1"; "IList_Of_T")
        {
        }

        type("System.Security.Cryptography.RNGCryptoServiceProvider"; "RNGCryptoServiceProvider")
        {
        }

        type("System.Collections.Generic.KeyValuePair`2"; "KeyValuePair_Of_T_U")
        {
        }

        type("System.Threading.Tasks.Task`1"; "Task_Of_T")
        {
        }

        type("System.Security.Cryptography.X509Certificates.X509ContentType"; "X509ContentType")
        {
        }

        type("System.IO.TextWriter"; "TextWriter")
        {
        }

        type("System.IO.TextReader"; "TextReader")
        {
        }

        type("System.Nullable`1"; "Nullable_Of_T")
        {
        }

        type("System.Collections.Generic.IReadOnlyList`1"; "IReadOnlyList_Of_T")
        {
        }

        type("System.Collections.Generic.Dictionary`2+KeyCollection"; "Dictionary_Of_T_U_KeyCollection")
        {
        }
    }

    assembly("System")
    {
        Version = '4.0.0.0';
        Culture = 'neutral';
        PublicKeyToken = 'b77a5c561934e089';

        type("System.Diagnostics.ProcessStartInfo"; "ProcessStartInfo")
        {
        }

        type("System.Diagnostics.EventLog"; "EventLog")
        {
        }

        type("System.Diagnostics.EventLogEntryType"; "EventLogEntryType")
        {
        }

        type("System.UriTypeConverter"; "UriTypeConverter")
        {
        }

        type("Microsoft.CSharp.CSharpCodeProvider"; "CSharpCodeProvider")
        {
        }

        type("System.Net.HttpRequestHeader"; "HttpRequestHeader")
        {
        }

        type("System.Diagnostics.PerformanceCounter"; "PerformanceCounter")
        {
        }

        type("System.Net.FtpWebRequest"; "FtpWebRequest")
        {
        }

        type("System.Net.FtpWebResponse"; "FtpWebResponse")
        {
        }

        type("System.Net.ServicePoint"; "ServicePoint")
        {
        }

        type("System.Net.NetworkInformation.IPGlobalProperties"; "IPGlobalProperties")
        {
        }

        type("System.Net.WebResponse"; "WebResponse")
        {
        }

        type("System.Diagnostics.ProcessWindowStyle"; "ProcessWindowStyle")
        {
        }

        type("System.ComponentModel.PropertyChangedEventArgs"; "PropertyChangedEventArgs")
        {
        }

        type("System.ComponentModel.PropertyChangingEventArgs"; "PropertyChangingEventArgs")
        {
        }

        type("System.ComponentModel.ListChangedEventArgs"; "ListChangedEventArgs")
        {
        }

        type("System.ComponentModel.AddingNewEventArgs"; "AddingNewEventArgs")
        {
        }

        type("System.Collections.Specialized.NotifyCollectionChangedEventArgs"; "NotifyCollectionChangedEventArgs")
        {
        }
    }

    assembly("DocumentFormat.OpenXml")
    {
        PublicKeyToken = '8fb06cb64d019a17';

        type("DocumentFormat.OpenXml.Drawing.AdjustValueList"; "AdjustValueList")
        {
        }

        type("DocumentFormat.OpenXml.Drawing.Blip"; "Blip")
        {
        }

        type("DocumentFormat.OpenXml.Drawing.BlipCompressionValues"; "BlipCompressionValues")
        {
        }

        type("DocumentFormat.OpenXml.Drawing.BlipExtension"; "BlipExtension")
        {
        }

        type("DocumentFormat.OpenXml.Drawing.BlipExtensionList"; "BlipExtensionList")
        {
        }

        type("DocumentFormat.OpenXml.Drawing.Extents"; "Extents")
        {
        }

        type("DocumentFormat.OpenXml.Drawing.FillRectangle"; "FillRectangle")
        {
        }

        type("DocumentFormat.OpenXml.Drawing.Graphic"; "Graphic")
        {
        }

        type("DocumentFormat.OpenXml.Drawing.GraphicData"; "GraphicData")
        {
        }

        type("DocumentFormat.OpenXml.Drawing.Offset"; "Offset")
        {
        }

        type("DocumentFormat.OpenXml.Drawing.Pictures.BlipFill"; "BlipFill")
        {
        }

        type("DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties"; "NonVisualDrawingProperties")
        {
        }

        type("DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties"; "NonVisualPictureDrawingProperties")
        {
        }

        type("DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties"; "NonVisualPictureProperties")
        {
        }

        type("DocumentFormat.OpenXml.Drawing.Pictures.Picture"; "Picture")
        {
        }

        type("DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties"; "ShapeProperties")
        {
        }

        type("DocumentFormat.OpenXml.Drawing.PresetGeometry"; "PresetGeometry")
        {
        }

        type("DocumentFormat.OpenXml.Drawing.ShapeTypeValues"; "ShapeTypeValues")
        {
        }

        type("DocumentFormat.OpenXml.Drawing.Stretch"; "Stretch")
        {
        }

        type("DocumentFormat.OpenXml.Drawing.Transform2D"; "Transform2D")
        {
        }

        type("DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties"; "DocProperties")
        {
        }

        type("DocumentFormat.OpenXml.Drawing.Wordprocessing.EffectExtent"; "EffectExtent")
        {
        }

        type("DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent"; "Extent")
        {
        }

        type("DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline"; "Inline")
        {
        }

        type("DocumentFormat.OpenXml.Drawing.Wordprocessing.NonVisualGraphicFrameDrawingProperties"; "NonVisualGraphicFrameDrawingProperties")
        {
        }

        type("DocumentFormat.OpenXml.Int64Value"; "Int64Value")
        {
        }

        type("DocumentFormat.OpenXml.Office2013.ExcelAc.AbsolutePath"; "AbsolutePath")
        {
        }

        type("DocumentFormat.OpenXml.OnOffValue"; "OnOffValue")
        {
        }

        type("DocumentFormat.OpenXml.OpenXmlAttribute"; "OpenXmlAttribute")
        {
        }

        type("DocumentFormat.OpenXml.OpenXmlElementList"; "OpenXmlElementList")
        {
        }

        type("DocumentFormat.OpenXml.Packaging.AlternativeFormatImportPart"; "AlternativeFormatImportPart")
        {
        }

        type("DocumentFormat.OpenXml.Packaging.AlternativeFormatImportPartType"; "AlternativeFormatImportPartType")
        {
        }

        type("DocumentFormat.OpenXml.Packaging.DocumentSettingsPart"; "DocumentSettingsPart")
        {
        }

        type("DocumentFormat.OpenXml.Packaging.FooterPart"; "FooterPart")
        {
        }

        type("DocumentFormat.OpenXml.Packaging.HeaderPart"; "HeaderPart")
        {
        }

        type("DocumentFormat.OpenXml.Packaging.ImagePart"; "ImagePart")
        {
        }

        type("DocumentFormat.OpenXml.Packaging.ImagePartType"; "ImagePartType")
        {
        }

        type("DocumentFormat.OpenXml.Packaging.MainDocumentPart"; "MainDocumentPart")
        {
        }

        type("DocumentFormat.OpenXml.Packaging.SpreadsheetDocument"; "SpreadsheetDocument")
        {
        }

        type("DocumentFormat.OpenXml.Packaging.WordprocessingDocument"; "WordprocessingDocument")
        {
        }

        type("DocumentFormat.OpenXml.Packaging.WorkbookStylesPart"; "WorkbookStylesPart")
        {
        }

        type("DocumentFormat.OpenXml.Packaging.WorksheetPart"; "WorksheetPart")
        {
        }

        type("DocumentFormat.OpenXml.SpaceProcessingModeValues"; "SpaceProcessingModeValues")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.Alignment"; "Alignment")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.BookViews"; "BookViews")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.Border"; "Border")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.BottomBorder"; "BottomBorder")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues"; "BorderStyleValues")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.CalculationCell"; "CalculationCell")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.CalculationChain"; "CalculationChain")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.CalculationProperties"; "CalculationProperties")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.Cell"; "Cell")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.CellFormat"; "CellFormat")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.CellValue"; "CellValue")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.CellValues"; "CellValues")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.CellWatches"; "CellWatches")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.Color"; "Color1")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.ColumnBreaks"; "ColumnBreaks")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.Column"; "Column1")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.Columns"; "Columns1")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatting"; "ConditionalFormatting")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.Controls"; "Controls")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.CustomProperties"; "CustomProperties")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.CustomSheetViews"; "CustomSheetViews")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.CustomWorkbookViews"; "CustomWorkbookViews")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.DataConsolidate"; "DataConsolidate")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.DataValidations"; "DataValidations")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.DefinedName"; "DefinedName")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.DefinedNames"; "DefinedNames")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.DiagonalBorder"; "DiagonalBorder")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.Drawing"; "Drawing1")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.DrawingHeaderFooter"; "DrawingHeaderFooter")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.ExternalReferences"; "ExternalReferences")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.FileRecoveryProperties"; "FileRecoveryProperties")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.FileSharing"; "FileSharing")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.FileVersion"; "FileVersion")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.ForegroundColor"; "ForegroundColor")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.FunctionGroups"; "FunctionGroups")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.HeaderFooter"; "HeaderFooter")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues"; "HorizontalAlignmentValues")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.Hyperlinks"; "Hyperlinks")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.IgnoredErrors"; "IgnoredErrors")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.Italic"; "Italic1")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.LeftBorder"; "LeftBorder")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.LegacyDrawingHeaderFooter"; "LegacyDrawingHeaderFooter")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.MergeCell"; "MergeCell")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.MergeCells"; "MergeCells")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.NumberingFormat"; "NumberingFormat")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.NumberingFormats"; "NumberingFormats")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.OddHeader"; "OddHeader")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.OddFooter"; "OddFooter")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.OleObjects"; "OleObjects")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.OleSize"; "OleSize")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.PageMargins"; "PageMargins")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.PageSetup"; "PageSetup")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.PageSetupProperties"; "PageSetupProperties")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.Pane"; "Pane")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.PaneValues"; "PaneValues")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.PaneStateValues"; "PaneStateValues")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.PhoneticProperties"; "PhoneticProperties")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.PrintOptions"; "PrintOptions")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.Picture"; "Picture1")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.PivotCaches"; "PivotCaches")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.ProtectedRanges"; "ProtectedRanges")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.RightBorder"; "RightBorder")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.Row"; "Row")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.RowBreaks"; "RowBreaks")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.Scenarios"; "Scenarios")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.SortState"; "SortState")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.SharedStringItem"; "SharedStringItem")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.SharedStringTable"; "SharedStringTable")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.Sheet"; "Sheet")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.SheetCalculationProperties"; "SheetCalculationProperties")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.SheetData"; "SheetData")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.SheetDimension"; "SheetDimension")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.SheetFormatProperties"; "SheetFormatProperties")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.SheetProperties"; "SheetProperties")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.SheetProtection"; "SheetProtection")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.Sheets"; "Sheets")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.SheetView"; "SheetView")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.SheetViews"; "SheetViews")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.TableParts"; "TableParts")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.TopBorder"; "TopBorder")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.Underline"; "Underline1")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.UnderlineValues"; "UnderlineValues1")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues"; "VerticalAlignmentValues")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.WebPublishing"; "WebPublishing")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.WebPublishItems"; "WebPublishItems")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.WebPublishObjects"; "WebPublishObjects")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.WorkbookExtensionList"; "WorkbookExtensionList")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.WorkbookProperties"; "WorkbookProperties")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.WorkbookProtection"; "WorkbookProtection")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.WorkbookView"; "WorkbookView")
        {
        }

        type("DocumentFormat.OpenXml.Spreadsheet.WorksheetExtensionList"; "WorksheetExtensionList")
        {
        }

        type("DocumentFormat.OpenXml.SpreadsheetDocumentType"; "SpreadsheetDocumentType")
        {
        }

        type("DocumentFormat.OpenXml.Wordprocessing.AltChunk"; "AltChunk")
        {
        }

        type("DocumentFormat.OpenXml.Wordprocessing.AttachedTemplate"; "AttachedTemplate")
        {
        }

        type("DocumentFormat.OpenXml.Wordprocessing.Bold"; "Bold0")
        {
        }

        type("DocumentFormat.OpenXml.Wordprocessing.BookmarkEnd"; "BookmarkEnd")
        {
        }

        type("DocumentFormat.OpenXml.Wordprocessing.BookmarkStart"; "BookmarkStart")
        {
        }

        type("DocumentFormat.OpenXml.Wordprocessing.Break"; "Break")
        {
        }

        type("DocumentFormat.OpenXml.Wordprocessing.BreakValues"; "BreakValues")
        {
        }

        type("DocumentFormat.OpenXml.Wordprocessing.Drawing"; "Drawing")
        {
        }

        type("DocumentFormat.OpenXml.Wordprocessing.Italic"; "Italic")
        {
        }

        type("DocumentFormat.OpenXml.Wordprocessing.PageMargin"; "PageMargin")
        {
        }

        type("DocumentFormat.OpenXml.Wordprocessing.PageSize"; "PageSize")
        {
        }

        type("DocumentFormat.OpenXml.Wordprocessing.Paragraph"; "Paragraph")
        {
        }

        type("DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties"; "ParagraphProperties")
        {
        }

        type("DocumentFormat.OpenXml.Wordprocessing.Run"; "Run0")
        {
        }

        type("DocumentFormat.OpenXml.Wordprocessing.RunProperties"; "RunProperties0")
        {
        }

        type("DocumentFormat.OpenXml.Wordprocessing.SectionProperties"; "SectionProperties")
        {
        }

        type("DocumentFormat.OpenXml.Wordprocessing.TableCell"; "TableCell")
        {
        }

        type("DocumentFormat.OpenXml.Wordprocessing.TableCellProperties"; "TableCellProperties")
        {
        }

        type("DocumentFormat.OpenXml.Wordprocessing.TableCellWidth"; "TableCellWidth")
        {
        }

        type("DocumentFormat.OpenXml.Wordprocessing.Table"; "WordProcessingTable")
        {
        }

        type("DocumentFormat.OpenXml.Wordprocessing.TableRow"; "TableRow")
        {
        }

        type("DocumentFormat.OpenXml.Wordprocessing.Text"; "Text0")
        {
        }

        type("DocumentFormat.OpenXml.Wordprocessing.Underline"; "Underline")
        {
        }

        type("DocumentFormat.OpenXml.Wordprocessing.UnderlineValues"; "UnderlineValues")
        {
        }

        type("DocumentFormat.OpenXml.WordprocessingDocumentType"; "WordprocessingDocumentType")
        {
        }
    }

    assembly("Microsoft.Office.Interop.Word")
    {
        Culture = 'neutral';
        PublicKeyToken = '71e9bce111e9429c';

        type("Microsoft.Office.Interop.Word.Range"; "Range")
        {
        }

        type("Microsoft.Office.Interop.Word.Dialogs"; "Dialogs")
        {
        }

        type("Microsoft.Office.Interop.Word.Dialog"; "Dialog")
        {
        }

        type("Microsoft.Office.Interop.Word.Window"; "Window")
        {
        }

        type("Microsoft.Office.Interop.Word.Selection"; "Selection0")
        {
        }

        type("Microsoft.Office.Interop.Word.XMLNode"; "XMLNode0")
        {
        }

        type("Microsoft.Office.Interop.Word.ProtectedViewWindow"; "ProtectedViewWindow")
        {
        }

        type("Microsoft.Office.Interop.Word.MailMergeField"; "MailMergeField")
        {
        }

        type("Microsoft.Office.Interop.Word.MailMergeDataField"; "MailMergeDataField")
        {
        }

        type("Microsoft.Office.Interop.Word.Documents"; "Documents")
        {
        }

        type("Microsoft.Office.Interop.Word.Rows"; "Rows")
        {
        }

        type("Microsoft.Office.Interop.Word.Borders"; "Borders")
        {
        }

        type("Microsoft.Office.Interop.Word.Row"; "Row0")
        {
        }

        type("Microsoft.Office.Interop.Word.Columns"; "Columns")
        {
        }

        type("Microsoft.Office.Interop.Word.Column"; "Column")
        {
        }

        type("Microsoft.Office.Interop.Word.Bookmarks"; "Bookmarks")
        {
        }
    }

    assembly("System.Xml")
    {
        Version = '4.0.0.0';
        Culture = 'neutral';
        PublicKeyToken = 'b77a5c561934e089';

        type("System.Xml.XmlNodeChangedEventArgs"; "XmlNodeChangedEventArgs")
        {
        }

        type("System.Xml.XmlDocumentFragment"; "XmlDocumentFragment")
        {
        }

        type("System.Xml.Xsl.XslTransform"; "XslTransform")
        {
        }

        type("System.Xml.Serialization.XmlRootAttribute"; "XmlRootAttribute")
        {
        }
    }

    assembly("WindowsBase")
    {
        Version = '4.0.0.0';
        Culture = 'neutral';
        PublicKeyToken = '31bf3856ad364e35';

        type("System.IO.Packaging.Package"; "Package")
        {
        }

        type("System.IO.Packaging.PackagePart"; "PackagePart")
        {
        }
    }

    assembly("System.Drawing")
    {
        Version = '4.0.0.0';
        Culture = 'neutral';
        PublicKeyToken = 'b03f5f7f11d50a3a';

        type("System.Drawing.Size"; "Size")
        {
        }

        type("System.Drawing.Color"; "Color0")
        {
        }
    }

    assembly("System.Data")
    {
        Version = '4.0.0.0';
        Culture = 'neutral';
        PublicKeyToken = 'b77a5c561934e089';

        type("System.Data.DataRowVersion"; "DataRowVersion")
        {
        }

        type("System.Data.SqlClient.SqlConnection"; "SqlConnection")
        {
        }

        type("System.Data.SqlClient.SqlCommand"; "SqlCommand")
        {
        }

        type("System.Data.SqlClient.SqlDataReader"; "SqlDataReader")
        {
        }

        type("System.Data.SqlClient.SqlInfoMessageEventArgs"; "SqlInfoMessageEventArgs")
        {
        }

        type("System.Data.StateChangeEventArgs"; "StateChangeEventArgs")
        {
        }

        type("System.Data.StatementCompletedEventArgs"; "StatementCompletedEventArgs")
        {
        }

    }

    assembly("Microsoft.Dynamics.Nav.Types")
    {
        PublicKeyToken = '31bf3856ad364e35';

        type("Microsoft.Dynamics.Nav.Types.ServerUserSettings"; "ServerUserSettings")
        {
        }

        type("Microsoft.Dynamics.Nav.Types.DatabaseServer"; "DatabaseServer")
        {
        }

        type("Microsoft.Dynamics.Nav.Types.DatabaseInstance"; "DatabaseInstance")
        {
        }

        type("Microsoft.Dynamics.Nav.Types.DatabaseName"; "DatabaseName")
        {
        }

        type("Microsoft.Dynamics.Nav.Types.ALConfigSettings"; "ALConfigSettings")
        {
        }

        type("Microsoft.Dynamics.Nav.Types.SettingsChangedEventArgs"; "SettingsChangedEventArgs")
        {
        }
    }

    assembly("System.ServiceModel")
    {
        Version = '4.0.0.0';
        Culture = 'neutral';
        PublicKeyToken = 'b77a5c561934e089';

        type("System.ServiceModel.BasicHttpBinding"; "BasicHttpBinding")
        {
        }

        type("System.ServiceModel.BasicHttpSecurityMode"; "BasicHttpSecurityMode")
        {
        }

        type("System.ServiceModel.EndpointAddress"; "EndpointAddress")
        {
        }

        type("System.ServiceModel.HttpClientCredentialType"; "HttpClientCredentialType")
        {
        }

        type("System.ServiceModel.WSMessageEncoding"; "WSMessageEncoding")
        {
        }

        type("System.ServiceModel.Description.ClientCredentials"; "ClientCredentials")
        {
        }
    }

    assembly("System.Messaging")
    {
        Version = '4.0.0.0';
        Culture = 'neutral';
        PublicKeyToken = 'b03f5f7f11d50a3a';

        type("System.Messaging.MessageQueue"; "MessageQueue")
        {
        }

        type("System.Messaging.XmlMessageFormatter"; "XmlMessageFormatter")
        {
        }

        type("System.Messaging.Message"; "Message")
        {
        }

        type("System.Messaging.PeekCompletedEventArgs"; "PeekCompletedEventArgs")
        {
        }

        type("System.Messaging.ReceiveCompletedEventArgs"; "ReceiveCompletedEventArgs")
        {
        }
    }

    assembly("System.Security")
    {
        Version = '4.0.0.0';
        Culture = 'neutral';
        PublicKeyToken = 'b03f5f7f11d50a3a';

        type("System.Security.Cryptography.X509Certificates.X509SelectionFlag"; "X509SelectionFlag")
        {
        }
    }

    assembly("System.Net.Http")
    {
        Version = '4.0.0.0';
        Culture = 'neutral';
        PublicKeyToken = 'b03f5f7f11d50a3a';

        type("System.Net.Http.StringContent"; "StringContent")
        {
        }

        type("System.Net.Http.HttpRequestMessage"; "HttpRequestMessage")
        {
        }

        type("System.Net.Http.HttpMethod"; "HttpMethod")
        {
        }

        type("System.Net.Http.Headers.HttpHeaderValueCollection`1"; "HttpHeaderValueCollection_Of_T")
        {
        }

        type("System.Net.Http.MultipartFormDataContent"; "MultipartFormDataContent")
        {
        }

        type("System.Net.Http.ByteArrayContent"; "ByteArrayContent")
        {
        }
    }

    assembly("System.Xml")  //added missing types
    {
        Version = '4.0.0.0';
        Culture = 'neutral';
        PublicKeyToken = 'b77a5c561934e089';

        type("System.Xml.Serialization.XmlSerializer"; "XmlSerializer")
        {
        }
    }

    assembly("Microsoft.SharePoint.Client.Runtime")
    {
        Version = '16.1.0.0';
        Culture = 'neutral';
        PublicKeyToken = '71e9bce111e9429c';

        type("Microsoft.SharePoint.Client.SharePointOnlineCredentials"; "SharePointOnlineCredentials")
        {
        }
        type("Microsoft.SharePoint.Client.ClientRuntimeContext"; "ClientRuntimeContext0")
        {
        }
    }

    assembly("Microsoft.SharePoint.Client")
    {
        Version = '16.1.0.0';
        Culture = 'neutral';
        PublicKeyToken = '71e9bce111e9429c';

        type("Microsoft.SharePoint.Client.ClientContext"; "ClientContext0")
        {
        }

        type("Microsoft.SharePoint.Client.Web"; "Web0")
        {
        }

        type("Microsoft.SharePoint.Client.File"; "File1")
        {
        }

        type("Microsoft.SharePoint.Client.Folder"; "Folder1")
        {
        }

        type("Microsoft.SharePoint.Client.MoveOperations"; MoveOperations)
        {
        }

        type("Microsoft.SharePoint.Client.FileCreationInformation"; FileCreationInformation)
        {
        }
    }

    assembly("Newtonsoft.Json")
    {

        type("Newtonsoft.Json.JsonToken"; "JsonToken")
        {
        }
    }

    //eVerbinding
    assembly("System.Xml.Linq")
    {
        Version = '4.0.0.0';
        Culture = 'neutral';
        PublicKeyToken = 'b77a5c561934e089';

        type("System.Xml.Linq.XElement"; "XElement")
        {
        }
    }
}
