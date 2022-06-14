/*using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
namespace WordFormatter
{
	class Program

	{

		[STAThread] //to communicate with Windows OS and System dialogs
		static void Main(string[] args)
		{

			Application.EnableVisualStyles(); //to enable colour fonts and other visual elements on the Form
			Application.SetCompatibleTextRenderingDefault(false);
			Application.Run(new FileSelectForm()); // opens up the dialog
			foreach (String file in FileSelectForm.fileToOpen) //for every file selected by the user
			{

				WordprocessingDocument doc = WordprocessingDocument.Open(file, true);
				MainDocumentPart mainDoc = doc.MainDocumentPart;
				using (doc)
				{

					Body body = mainDoc.Document.Body;
					Paragraph para1 = body.AppendChild(new Paragraph());
					Run run = para1.AppendChild(new Run());
					run.AppendChild(new Text("Titleee"));


					Table approvalsTable = body.Elements<Table>().First();
					//Adding Test Lead Name
					TableRow testLeadRow = approvalsTable.Elements<TableRow>().ElementAt(3);
					TableCell testLeadNameCell = testLeadRow.Elements<TableCell>().ElementAt(1);
					testLeadNameCell.AppendChild(new Paragraph()).AppendChild(new Run()).AppendChild(new Text("Kim Davis"));
					//Adding Quality Name
					approvalsTable.Elements<TableRow>().ElementAt(4).Elements<TableCell>().ElementAt(1).AppendChild(new Paragraph()).AppendChild(new Run()).AppendChild(new Text("Sarah Spleidt"));

					int totalNumberOfTablesInDoc = mainDoc.Document.Body.Elements<Table>().Count();
					// Handle only the test case tables
					for (int i = 7; i < totalNumberOfTablesInDoc - 1; i++)
					{

						//doc.MainDocumentPart.Document.Body.AppendChild(new Paragraph()).AppendChild(new Run()).AppendChild(new Text("Software Version: -________________"));
						//doc.MainDocumentPart.Document.Body.AppendChild(new Paragraph()).AppendChild(new Run()).AppendChild(new Break());
						//doc.MainDocumentPart.Document.Body.AppendChild(new Paragraph()).AppendChild(new Run()).AppendChild(new Text("Initial / Date: - ________________"));


						Table table1 = mainDoc.Document.Body.Elements<Table>().ElementAt(i);
						IEnumerable<TableRow> rows = table1.Elements<TableRow>();
						int numberOfCols = rows.ElementAt(0).Elements<TableCell>().Count();
						//Change Heading of 4th column
						rows.ElementAt(0).Elements<TableCell>().ElementAt(3).Elements<Paragraph>().First().Elements<Run>().First().Elements<Text>().First().Remove();
						rows.ElementAt(0).Elements<TableCell>().ElementAt(3).Elements<Paragraph>().First().Elements<Run>().First().AppendChild(new Text("Pass/Fail"));

						for (int r = 1; r < rows.Count(); r++)

						{
							rows.ElementAt(r).Elements<TableCell>().ElementAt(3).Elements<Paragraph>().First().AppendChild(new Run()).AppendChild(new Text("P____/F____"));

						}

						// Auto fit

						TableCellProperties tableCellProperties = new TableCellProperties();
						for (int r = 0; r < rows.Count(); r++)
						{
							for (int c = 0; c < numberOfCols; c++)
							{
								TableCellWidth tableCellWidth = new TableCellWidth() { Width = "0", Type = TableWidthUnitValues.Auto };
								rows.ElementAt(r).Elements<TableCell>().ElementAt(c).Elements<TableCellProperties>().First().TableCellWidth = tableCellWidth;
							}
						}
					}


					//Adding Footer
					FooterPart footerPart = mainDoc.AddNewPart<FooterPart>();
					string footerPartId = mainDoc.GetIdOfPart(footerPart);
					GenerateFooterPartContent(footerPart);

					IEnumerable<SectionProperties> sections = mainDoc.Document.Body.Elements<SectionProperties>();

					foreach (var section in sections)
					{
						// Delete existing references footers

						section.RemoveAllChildren<FooterReference>();

						// Create new footer reference node

						section.PrependChild<FooterReference>(new FooterReference() { Id = footerPartId });
					}

				}
			}

			MessageBox.Show("Formatting Done");

		}

		public static void GenerateFooterPartContent(FooterPart part)
		{

			Footer footer = new Footer();

			Paragraph footerPara1 = new Paragraph();
			ParagraphProperties footerPara1Properties = new ParagraphProperties();
			ParagraphStyleId objParagraphStyleId = new ParagraphStyleId() { Val = "Footer" };
			Indentation indentation = new Indentation() { Right = "260" };
			footerPara1Properties.Append(objParagraphStyleId);
			footerPara1Properties.Append(indentation);


			Run run01 = new Run();
			Text text1 = new Text() { Space = SpaceProcessingModeValues.Preserve };
			text1.Text = "Becton Dickinson Proprietary Information  ";
			run01.Append(text1);

			Run run02 = new Run();
			Text text2 = new Text();
			text2.Text = "Enter Feature Name";
			run02.Append(text2);

			Run run03 = new Run();
			Text text3 = new Text() { Space = SpaceProcessingModeValues.Preserve };
			text3.Text = "   Rev 1.0";
			run03.Append(text3);



			footerPara1.Append(footerPara1Properties);
			footerPara1.Append(run01);
			footerPara1.Append(run02);
			footerPara1.Append(run03);


			Paragraph footerPara2 = new Paragraph();
			ParagraphMarkRunProperties paragraphMarkRunProperties = new ParagraphMarkRunProperties();
			RunStyle runStylePara = new RunStyle() { Val = "PageNumber" };
			paragraphMarkRunProperties.Append(runStylePara);
			FrameProperties frameProperties = new FrameProperties() { Wrap = TextWrappingValues.Around, HorizontalPosition = HorizontalAnchorValues.Margin, VerticalPosition = VerticalAnchorValues.Text, XAlign = HorizontalAlignmentValues.Right, Y = "1" };
			ParagraphProperties footerPara2Properties = new ParagraphProperties();
			ParagraphStyleId objParagraphStyleId2 = new ParagraphStyleId() { Val = "Footer" };
			footerPara2Properties.Append(objParagraphStyleId2);
			footerPara2Properties.Append(frameProperties);
			footerPara2Properties.Append(paragraphMarkRunProperties);


			Run run1 = new Run();
			RunProperties runProperties1 = new RunProperties();
			RunStyle runStyle1 = new RunStyle() { Val = "PageNumber" };
			runProperties1.Append(runStyle1);
			FieldChar fieldChar1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };
			run1.Append(runProperties1);
			run1.Append(fieldChar1);

			Run run2 = new Run();
			RunProperties runProperties2 = new RunProperties();
			RunStyle runStyle2 = new RunStyle() { Val = "PageNumber" };
			runProperties2.Append(runStyle2);
			FieldCode fieldCode1 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
			fieldCode1.Text = " PAGE ";
			run2.Append(runProperties2);
			run2.Append(fieldCode1);

			Run run3 = new Run();
			RunProperties runProperties5 = new RunProperties();
			RunStyle runStyle5 = new RunStyle() { Val = "PageNumber" };
			runProperties5.Append(runStyle5);
			FieldChar fieldChar3 = new FieldChar() { FieldCharType = FieldCharValues.End };
			run3.Append(runProperties5);
			run3.Append(fieldChar3);


			footerPara2.Append(footerPara2Properties);
			footerPara2.Append(run1);
			footerPara2.Append(run2);
			footerPara2.Append(run3);


			footer.Append(footerPara2);
			footer.Append(footerPara1);


			part.Footer = footer;

		}
	}
}


*/