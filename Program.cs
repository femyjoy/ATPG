using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Diagnostics;




namespace WordFormatter
{
	class Program
	{

		static int numberOfFiles = 0;
		//start time
		static DateTime StartAt = DateTime.Now;
		static string finalDir = null;
		[STAThread] //to communicate with Windows OS and System dialogs
		static void Main(string[] args)
		{
			//ver 2.0
			if (Process.GetProcessesByName(Process.GetCurrentProcess().ProcessName).Length > 1)
			{
				MessageBox.Show("There is another instance of this application already running. Please use only one instance at a time.");
				return;
			}

			Application.EnableVisualStyles(); //to enable colour fonts and other visual elements on the Form
			Application.SetCompatibleTextRenderingDefault(true);
			Application.Run(new FileSelectForm()); // opens up the dialog

			string intermediatDir = null;

			//ver 2.0
			try
			{
				string inputFileName = null;
				string[] inputFiles = FileSelectForm.fileToOpen;
				List<string> inputFilesList = inputFiles.ToList();
				List<string> interList = new List<string>();
				interList.AddRange(inputFilesList);
				WordprocessingDocument doc = null;
				MainDocumentPart mainDoc = null;
				Body body = null;
				List<string> filesNotProcessed = new List<string>();

				foreach (string initialFile in inputFilesList)
				{
					doc = WordprocessingDocument.Open(initialFile, false);
					mainDoc = doc.MainDocumentPart;
					body = mainDoc.Document.Body;

					//Handle empty input .docx file and already formatted file//ver 2.0
					if (body.GetFirstChild<Paragraph>().Descendants().ToList().Count() == 0 || !body.Descendants().OfType<Text>().ToList()[0].Text.Contains("Test plan"))
					{
						inputFileName = Path.GetFileNameWithoutExtension(initialFile);
						filesNotProcessed.Add(initialFile);
						interList.Remove(initialFile);
					}
				}

				if (filesNotProcessed.Count == 1)
				{
					MessageBox.Show("The file \"" + inputFileName + "\" is not a valid test case file and hence will not be processed.");
					if (inputFiles.Length == 1)
					{
						Environment.Exit(0);
					}

				}
				else if (filesNotProcessed.Count > 1)
				{
					MessageBox.Show("The following files are not valid test case files and hence will not be processed:\n" + filesNotProcessed.GetRange(0, filesNotProcessed.Count - 1));
				}

				inputFilesList = interList;

				//foreach (string file in FileSelectForm.fileToOpen) //for every file selected by the user
				foreach (string file in inputFilesList) //for every file selected by the user
				{

					inputFileName = Path.GetFileNameWithoutExtension(file);


					String currDir = Path.GetDirectoryName(file);
					intermediatDir = Path.Combine(currDir, "Intermediate");
					Directory.CreateDirectory(intermediatDir);
					String inputFullFileName = Path.GetFileName(file);
					String intermediateFile = intermediatDir + "\\" + inputFileName + "_Formatted.docx";
					//string testFile = @"C:\\Data\\Test.docx";
					//there test two files
					//string[] filepaths = new[] { currDir+"\\T.docx", currFile};
					File.Copy(file, intermediateFile, true); //srcFile, destFile, overwrite if exists


					String currFile = intermediateFile;
					numberOfFiles++;


					if (numberOfFiles == 1)
					{
						StartAt = DateTime.Now;
					}

					//ver 2.0
					try
					{
						doc = WordprocessingDocument.Open(intermediateFile, true);
					}
					catch (IOException e)
					{
						MessageBox.Show("Please close all the input files and try again.");
						Environment.Exit(1);
					}
					mainDoc = doc.MainDocumentPart;
					using (doc)
					{

						body = mainDoc.Document.Body;

						List<Paragraph> initialParas = body.Elements<Paragraph>().ToList();


						//remove the first paragraph and the first table which are not needed
						body.Elements<Table>().First().Remove();

						/*for (int i = 0; i <= 2; i++)
						{*/
						initialParas[0].Remove();
						initialParas[2].Remove();
						/*	}	*/

						mainDoc.Document.Save();
						//All parahraphs directly under body
						List<Paragraph> paragraphs = body.Elements<Paragraph>().ToList();

						//remove the lines between the tables from the whole document
						List<Paragraph> pWithPicture = paragraphs.Where<Paragraph>(p => p.Descendants().OfType<Picture>().ToList().Count() != 0).ToList();
						foreach (var p in pWithPicture)
						{
							p.Remove();
						}

						mainDoc.Document.Save();

						/*var runProp = new RunProperties(
							 new RunFonts()
							 {
								 Ascii = "Arial",
								 ComplexScript = "Arial",
								 HighAnsi = "Arial"
							 }
							 ) ;*/

						//var runFont = new RunFonts { Ascii = "Arial" };
						var fontSize = new FontSize { Val = new StringValue("20") };
						var fontSizeCS = new FontSizeComplexScript { Val = new StringValue("20") };

						/*runProp.Append(fontSize);
						runProp.Append(fontSizeCS);*/

						var runFont = new RunFonts();
						runFont.EastAsia = "Arial";
						runFont.Ascii = "Arial";
						runFont.ComplexScript = "Arial";
						runFont.HighAnsi = "Arial";

						List<Paragraph> pWithRun = new List<Paragraph>();
						List<Paragraph> pWithTextsOutsideTable = new List<Paragraph>();
						List<Paragraph> pWithPprRpr = new List<Paragraph>();

						try //ver 2.0
						{
							List<Paragraph> paragraphsUnderBody = body.Elements<Paragraph>().ToList(); //all ps outside table
							List<Paragraph> allParagraphs = body.Descendants<Paragraph>().ToList(); //includes p in table as well
							List<Table> allTables = body.Elements().OfType<Table>().ToList();
							pWithRun = allParagraphs.Where<Paragraph>(p => p.Descendants().OfType<Run>().ToList().Count() != 0).ToList();
							List<Paragraph> pWithRunOutsideTable = paragraphsUnderBody.Where<Paragraph>(p => p.Descendants().OfType<Run>().ToList().Count() != 0).ToList();
							pWithTextsOutsideTable = pWithRunOutsideTable.Where<Paragraph>(p => p.GetFirstChild<Run>().Descendants().OfType<Text>().ToList().Count() != 0).ToList();
							pWithPprRpr = allParagraphs.Where<Paragraph>(p => p.GetFirstChild<ParagraphProperties>().GetFirstChild<ParagraphMarkRunProperties>().ToList().Count() != 0).ToList();
						}
						catch (ArgumentNullException e)
						{
							MessageBox.Show("This document has already been formatted and a TestProtocol generated.");
							Environment.Exit(1);
						}

						/*foreach (var p in pWithPpr) //outside table - s/w version, test case
						{

							p.GetFirstChild<ParagraphProperties>().GetFirstChild<ParagraphMarkRunProperties>().GetFirstChild<RunFonts>().Ascii = "Segoe UI"; //add attributes to exiting node
							p.GetFirstChild<ParagraphProperties>().GetFirstChild<ParagraphMarkRunProperties>().GetFirstChild<RunFonts>().ComplexScript = "Arial";
							p.GetFirstChild<ParagraphProperties>().GetFirstChild<ParagraphMarkRunProperties>().GetFirstChild<RunFonts>().HighAnsi = "Arial";
							p.GetFirstChild<ParagraphProperties>().GetFirstChild<ParagraphMarkRunProperties>().Elements().ToList().Append<FontSize>(fontSize); //add a new child
							p.GetFirstChild<ParagraphProperties>().GetFirstChild<ParagraphMarkRunProperties>().Elements().ToList().Add(fontSizeCS);

							p.GetFirstChild<ParagraphProperties>().GetFirstChild<ParagraphMarkRunProperties>().ChildElements.ToList().Add(fontSize);
							p.GetFirstChild<ParagraphProperties>().GetFirstChild<ParagraphMarkRunProperties>().ChildElements.ToList().Add(fontSizeCS);
							//p.GetFirstChild<ParagraphProperties>().GetFirstChild<ParagraphMarkRunProperties>().InsertAfter(fontSize, p.GetFirstChild<ParagraphProperties>().GetFirstChild<ParagraphMarkRunProperties>().GetFirstChild<RunFonts>());
							//p.GetFirstChild<ParagraphProperties>().GetFirstChild<ParagraphMarkRunProperties>().AppendChild<FontSize>((FontSize)fontSize.CloneNode(true));
							p.GetFirstChild<ParagraphProperties>().GetFirstChild<ParagraphMarkRunProperties>().AppendChild(new FontSize { Val = new StringValue("20") });
							p.GetFirstChild<ParagraphProperties>().GetFirstChild<ParagraphMarkRunProperties>().AppendChild(new FontSizeComplexScript { Val = new StringValue("20") });
							var runs = p.Descendants<Run>().ToList();

							foreach (var r in runs)
							{
								r.GetFirstChild<ParagraphMarkRunProperties>().AppendChild<ParagraphMarkRunProperties>(runProp);
							}

						}
	*/
						foreach (var p in pWithRun) //s/w version, Initial Date and Test Case
						{
							List<Run> runs = p.Descendants<Run>().ToList();

							foreach (var r in runs)
							{
								RunProperties rp = r.GetFirstChild<RunProperties>();
								rp.GetFirstChild<RunFonts>().Ascii = "Calibri";
								rp.GetFirstChild<RunFonts>().ComplexScript = "Calibri";
								rp.GetFirstChild<RunFonts>().HighAnsi = "Calibri";
								rp.AppendChild(new FontSize { Val = new StringValue("22") });
								rp.AppendChild(new FontSizeComplexScript { Val = new StringValue("22") });
							}

						}

						//set font type and size of spacing above software version and initial date and in in table content - heading
						foreach (var p in pWithPprRpr)
						{

							ParagraphMarkRunProperties pMRP = p.GetFirstChild<ParagraphProperties>().GetFirstChild<ParagraphMarkRunProperties>();
							pMRP.GetFirstChild<RunFonts>().Ascii = "Calibri";
							pMRP.GetFirstChild<RunFonts>().ComplexScript = "Calibri";
							pMRP.GetFirstChild<RunFonts>().HighAnsi = "Calibri";
							pMRP.AppendChild(new FontSize { Val = new StringValue("22") });
							pMRP.AppendChild(new FontSizeComplexScript { Val = new StringValue("22") });

						}

						//move TestCase line to above Software Version //ver 2.0
						List<Paragraph> pWithTestCaseID = pWithTextsOutsideTable.Where<Paragraph>(p => p.GetFirstChild<Run>().GetFirstChild<Text>().Text.Contains("Test case")).ToList();

						foreach (Paragraph p in pWithTestCaseID)
						{
							Paragraph newP = (Paragraph)p.CloneNode(true);
							p.PreviousSibling<Paragraph>().PreviousSibling<Paragraph>().PreviousSibling<Paragraph>().InsertAfterSelf<Paragraph>(newP);
							//p.PreviousSibling<Paragraph>().AppendChild(new Paragraph()).AppendChild(new Run()).AppendChild(new Break());
							p.Remove();
						}

						//ver 2.0
						//search and replace some texts
						searchAndReplace(mainDoc);


						foreach (var table in body.Elements().OfType<Table>())
						{

							//all rows in the table
							IEnumerable<TableRow> rows = table.Elements<TableRow>();
							//Defining table properties
							TableProperties tblProperties = new TableProperties();

							// Create Table Borders
							TableBorders tblBorders = new TableBorders();

							InsideHorizontalBorder insideHBorder = new InsideHorizontalBorder();
							insideHBorder.Val = new EnumValue<BorderValues>(BorderValues.Single);
							insideHBorder.Color = "";
							insideHBorder.Size = 6;
							insideHBorder.Space = 0;
							tblBorders.AppendChild(insideHBorder);
							InsideVerticalBorder insideVBorder = new InsideVerticalBorder();
							insideVBorder.Val = new EnumValue<BorderValues>(BorderValues.Single);
							insideVBorder.Color = "808080";
							insideVBorder.Size = 6;
							insideHBorder.Space = 0;
							tblBorders.AppendChild(insideVBorder);

							// Add the table borders to the properties
							table.GetFirstChild<TableProperties>().GetFirstChild<TableBorders>().AppendChild<TableBorders>(tblBorders); //append to existing child 

							// Auto fit at table level and cell level
							TableWidth tblW = new TableWidth();
							tblW.Type = TableWidthUnitValues.Auto;
							tblW.Width = "0";
							table.GetFirstChild<TableProperties>().GetFirstChild<TableWidth>().Remove();
							table.GetFirstChild<TableProperties>().Elements().ToList().Add(tblW);

							var cellsInTable = table.Descendants<TableCell>().ToList();
							foreach (var c in cellsInTable)

							{
								c.GetFirstChild<TableCellProperties>().GetFirstChild<TableCellWidth>().Remove();
								c.GetFirstChild<TableCellProperties>().Elements().ToList().Add(tblW);

							}

							// remove shading in some rows
							var cellsWithShading = cellsInTable.Where<TableCell>(c => c.GetFirstChild<TableCellProperties>().Descendants().OfType<Shading>().ToList().Count() != 0);
							foreach (var c in cellsWithShading)
							{

								c.GetFirstChild<TableCellProperties>().Descendants().OfType<Shading>().ToList()[0].Remove();
							}

							//make column heading bold and alignment correction
							Bold bold = new Bold();
							bold.Val = OnOffValue.FromBoolean(true);
							SpacingBetweenLines spacing = new SpacingBetweenLines();
							spacing.After = "240";
							spacing.Before = "240";

							//to make table header float to every new page											
							rows.ToList()[0].GetFirstChild<TableRowProperties>().AppendChild<TableHeader>(new TableHeader());


							var cellsInHeadingRow = table.GetFirstChild<TableRow>().Descendants().OfType<TableCell>().ToList();
							foreach (var c in cellsInHeadingRow)
							{

								if (c.GetFirstChild<Paragraph>().GetFirstChild<ParagraphProperties>().Descendants().OfType<SpacingBetweenLines>().ToList().Count() != 0)
								{
									c.GetFirstChild<Paragraph>().GetFirstChild<ParagraphProperties>().GetFirstChild<SpacingBetweenLines>().Remove();
								}
								c.GetFirstChild<Paragraph>().GetFirstChild<ParagraphProperties>().Elements().ToList().Add(spacing);
								c.GetFirstChild<Paragraph>().GetFirstChild<ParagraphProperties>().AppendChild(new Justification { Val = JustificationValues.Center });
								c.GetFirstChild<Paragraph>().GetFirstChild<Run>().GetFirstChild<RunProperties>().AppendChild<Bold>(new Bold()); //Bold

							}

							//insert sl no in first cell of every row except heading
							for (int i = 1; i <= rows.ToList().Count() - 1; i++)
							{
								rows.ToList()[i].GetFirstChild<TableCell>().AppendChild(new Paragraph());
								//Setting para properties
								Paragraph newP = rows.ToList()[i].GetFirstChild<TableCell>().Elements<Paragraph>().ElementAt(1);
								newP.AppendChild<ParagraphProperties>(new ParagraphProperties());
								newP.GetFirstChild<ParagraphProperties>().AppendChild(new Justification { Val = JustificationValues.Center });
								//Setting run properties and values
								newP.AppendChild<Run>(new Run());
								Run theRun = newP.GetFirstChild<Run>();
								theRun.AppendChild<RunProperties>(new RunProperties());
								theRun.GetFirstChild<RunProperties>().AppendChild(new RunFonts
								{
									Ascii = "Calibri",
									ComplexScript = "Calibri",
									HighAnsi = "Calibri"
								});
								theRun.GetFirstChild<RunProperties>().AppendChild(new FontSize { Val = new StringValue("22") });
								theRun.GetFirstChild<RunProperties>().AppendChild(new FontSizeComplexScript { Val = new StringValue("22") });
								theRun.AppendChild(new Text(i.ToString()));
							}
						}

						mainDoc.Document.Save();
					} // end for each doc formatting

					//ver 3.0
					mergeFiles(FileSelectForm.templateFileName, currFile, inputFileName);
				} // all documents formatted
			}
			catch (ArgumentNullException e) //ver 2.0
			{
				MessageBox.Show("The tool dialog was closed prematurely, please try again.");
				Environment.Exit(1);
			}


			//log the time taken for formatting
			DateTime EndAt = DateTime.Now;
			double timeTaken = Math.Round((EndAt - StartAt).TotalSeconds, 2);
			//int timeTaken = (EndAt - StartAt).Seconds;
			// This will give us the full name path of the executable file including the .exe file:
			string strExeFilePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
			//This will strip just the working path name:
			string strWorkPath = System.IO.Path.GetDirectoryName(strExeFilePath);
			string format = "dd-MMM-yyyy HH:mm";
			string inLog = DateTime.Now.ToString(format) + ":Time taken to format the testcase documents: " + timeTaken + " seconds.\n";
			File.AppendAllText(strWorkPath + "\\" + "log.txt", inLog);
			if (numberOfFiles == 1)
			{

				//MessageBox.Show("It took " + timeTaken + " seconds to format the file.\nThe formatted file is available in the location: "+ intermediatDir); //ver 2.0
				MessageBox.Show("It took " + timeTaken + " seconds to format the file.\nThe Test Protocol file is available in the location: " + finalDir); //ver 3.0

			}
			else
			{
				//MessageBox.Show("It took " + timeTaken + " seconds to format the files.\nThe formatted files are available in the location: " + intermediatDir); //ver 2.0
				MessageBox.Show("It took " + timeTaken + " seconds to format the files.\nThe corresponding Test Protocol files are available in the location: " + finalDir); //ver 3.0
			}
		}





		private static void searchAndReplace(MainDocumentPart mainDoc)
		{
			Body body = mainDoc.Document.Body;
			int n = 1;
			foreach (var text in body.Descendants<Text>())
			{
				if (text.Text.Contains("Test case"))
				{
					text.Text = text.Text.Replace("Test case", "1." + n + " ID");
					n++;
				}

				if (text.Text.Contains("Test Instructions"))
				{
					text.Text = text.Text.Replace("Test Instructions", "Action");
				}
			}
			mainDoc.Document.Save();
		}

		// ver 3.0
		private static void mergeFiles(String templateFile, String currFile, String inputFileName)
		{
			String currDir = Path.GetDirectoryName(currFile);
			//String fileName = Path.GetFileNameWithoutExtension(currFile);
			finalDir = Path.Combine(Directory.GetParent(currDir).FullName, "Final");
			Directory.CreateDirectory(finalDir);
			String destFilePath = finalDir + "\\" + inputFileName + "_TP.docx";
			//string testFile = @"C:\\Data\\Test.docx";
			//there test two files
			//string[] filepaths = new[] { currDir+"\\T.docx", currFile};
			string[] filepaths = new[] { templateFile, currFile };
			File.Copy(@filepaths[0], destFilePath, true); //srcFile. DestFile
														  //for (int i = 1; i < filepaths.Length; i++)
														  //using (WordprocessingDocument myDoc = WordprocessingDocument.Open(@filepaths[0], true))
			using (WordprocessingDocument myDoc = WordprocessingDocument.Open(destFilePath, true))
			{
				MainDocumentPart mainPart = myDoc.MainDocumentPart;
				Body body = mainPart.Document.Body;
				List<Paragraph> pWithRunOutsideTable = body.Elements<Paragraph>().ToList().Where<Paragraph>(p => p.Descendants().OfType<Run>().ToList().Count() != 0).ToList();
				List<Paragraph> pWithTextsOutsideTable = pWithRunOutsideTable.Where<Paragraph>(p => p.GetFirstChild<Run>().Descendants().OfType<Text>().ToList().Count() != 0).ToList();
				Paragraph pWithIntegrationTest = pWithTextsOutsideTable.Where<Paragraph>(p => p.GetFirstChild<Run>().GetFirstChild<Text>().Text.Contains("Integration Test")).ToList()[0];
				string altChunkId = "AltChunkId" + 1;
				AlternativeFormatImportPart chunk = mainPart.AddAlternativeFormatImportPart(
					AlternativeFormatImportPartType.WordprocessingML, altChunkId);
				using (FileStream fileStream = File.Open(@filepaths[1], FileMode.Open))
				{
					chunk.FeedData(fileStream);
				}
				AltChunk altChunk = new AltChunk();
				altChunk.Id = altChunkId;
				//mainPart.Document.Body.InsertAfter(altChunk, mainPart.Document.Body.Elements<Paragraph>().Last());
				mainPart.Document.Body.InsertAfter(altChunk, pWithIntegrationTest);
				mainPart.Document.Save();
				myDoc.Close();
			}
		}

	} // end of class
} //end of namespace

