using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Drawing.Image;
using static System.Net.Mime.MediaTypeNames;


namespace WordFormatter
{
	class FileSelectForm : Form
	{
		private TextBox textbox;
		private Button browse;
		private TextBox textboxtemplate; //ver 3.0
		private Button browseTemplate;
		private Button format;
		public static String[] fileToOpen;
		public static String templateFileName; //ver 3.0
		public FileSelectForm()
		{
			InitializeComponent();
			this.FormBorderStyle = FormBorderStyle.Fixed3D;
			//this.Icon = new Icon("C:\\ver3.0\\Data\\Picture1.ico");
			//Text = "Test Case Document Formatter-V1.0.0";
			//ForeColor = Color.FromArgb(0, 70, 170);
			Text = "Automatic Test Protocol Generator"; //ver 3.0

			
			Label label = new Label
			{
				MaximumSize = new Size(200, 0),
				AutoSize = true,
				ForeColor = Color.FromArgb(0,70, 170),
				Text = "Select the input .docx file(s) with the testcases to be formatted:",
				Location = new Point(15, 30)

			};


			textbox = new TextBox
			{
				Size = new Size(300, 100), //Width and Height
				Location = new Point(210, 30),
				AutoSize = true,
				BorderStyle = BorderStyle.FixedSingle,
				//WordWrap = true, //ver 2.0
				Multiline = true,
				//MaximumSize = new Size(200, 50),
				ScrollBars = ScrollBars.Both,
				ReadOnly = true,
				BackColor = System.Drawing.SystemColors.Window
			};


			browse = new Button
			{
				Text = "Browse",
				TextAlign = ContentAlignment.MiddleCenter,
				Size = new Size(80, 24), //Width and Height
				Location = new Point(520, 30),
				FlatStyle = FlatStyle.Flat
			};

			browse.Click += new EventHandler(browse_Click);


			//ver 3.0
			Label labelForTemplate = new Label
			{
				MaximumSize = new Size(200, 0),
				AutoSize = true,
				//ForeColor = System.Drawing.SystemColors.HotTrack,
				ForeColor = Color.FromArgb(0, 70, 170),
				//Font = new Font(Label.DefaultFont, FontStyle.Bold),
				Text = "Select the Test Protocol template file:",
				Location = new Point(15, 150)

			};


			//ver 3.0
			textboxtemplate = new TextBox
			{
				Size = new Size(280, 100), //Width and Height
				Location = new Point(210, 150),
				AutoSize = true,
				BorderStyle = BorderStyle.FixedSingle,
				WordWrap = true,
				ReadOnly = true,
				BackColor = System.Drawing.SystemColors.Window
			};

			//textboxtemplate.TextChanged += new EventHandler(enable_FormatButton); // no longer needed

			//ver 3.0
			browseTemplate = new Button
			{
				ForeColor = Color.FromArgb(0, 70, 170),
				Text = "Browse",
				TextAlign = ContentAlignment.MiddleCenter,
				Size = new Size(80, 24), //Width and Height
				Location = new Point(520, 150),
				Enabled = false
			};

			browseTemplate.Click += new EventHandler(browseTemplate_Click);

			format = new Button
			{
				Text = "Format",
				TextAlign = ContentAlignment.MiddleCenter,
				Size = new Size(80, 24),
				Location = new Point(210, 200),
				//Enabled = true // ver 2.0

			};
			format.Click += new EventHandler(format_Click);



			ClientSize = new Size(620, 250);
			//this.Controls.Add(pb);
			this.Controls.Add(label);
			this.Controls.Add(textbox);
			this.Controls.Add(browse);
			this.Controls.Add(labelForTemplate); //ver 3.0
			this.Controls.Add(textboxtemplate); //ver 3.0
			this.Controls.Add(browseTemplate); //ver 3.0
			this.Controls.Add(format);



		}

		private void browse_Click(object sender, EventArgs e)
		{
			var FD = new OpenFileDialog();
			FD.Multiselect = true;
			if (FD.ShowDialog() == DialogResult.OK)
			{
				fileToOpen = FD.FileNames;
				int numberOfFiles = fileToOpen.Length;
				textbox.Clear();

				for (int i = 0; i < numberOfFiles; i++)
				{

					textbox.Text += fileToOpen[i];
					if (i != numberOfFiles - 1)
					{
						textbox.Text += "," + Environment.NewLine;
					}
				}

				bool isDocx = true;
				foreach (String file in fileToOpen)
				{
					isDocx = isDocx && Path.GetExtension(file).Equals(".docx");
				}
				if (isDocx == true)
				{
					browseTemplate.Enabled = true;
					browseTemplate.FlatStyle = FlatStyle.Flat;
				}
				else
				{
					MessageBox.Show("One or many of the input files have an unsupported file format.\nPlease use only input files with \".docx\" extension and try again.");
				}
			}
		}

		//ver 3.0 // no longer needed
		/*private void enable_FormatButton(object sender, EventArgs e)
		{
			format.Enabled = true;
			//textbox.MinimumSize = new Size(200, textbox.TextLength *100);
		}*/

		//ver 3.0
		private void browseTemplate_Click(object sender, EventArgs e)
		{
			var FD = new OpenFileDialog();
			FD.Multiselect = false;
			if (FD.ShowDialog() == DialogResult.OK)
			{
				templateFileName = FD.FileNames[0];
				textboxtemplate.Text += templateFileName;
			}
			if (Path.GetExtension(templateFileName).Equals(".docx"))
			{
				format.Enabled = true;
				format.FlatStyle = FlatStyle.Flat;
			}
			else
			{
				format.Enabled = false;
				MessageBox.Show("The template file must have \".docx\" extension. Please save the file with \".docx\" extension and try again.");
			}
		}


		private void format_Click(object sender, EventArgs e)
		{
			this.Dispose();
		}

		
		private void button1_Paint(object sender, PaintEventArgs e)
		{
				ControlPaint.DrawBorder(e.Graphics, browse.ClientRectangle,
				SystemColors.ControlLightLight, 5, ButtonBorderStyle.Outset,
				SystemColors.ControlLightLight, 5, ButtonBorderStyle.Outset,
				SystemColors.ControlLightLight, 5, ButtonBorderStyle.Outset,
				SystemColors.ControlLightLight, 5, ButtonBorderStyle.Outset);
		}

		private void InitializeComponent()
		{
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FileSelectForm));
			this.SuspendLayout();
			// 
			// FileSelectForm
			// 
			this.AllowDrop = true;
			this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
			this.ClientSize = new System.Drawing.Size(292, 212);
			this.Font = new System.Drawing.Font("Segoe UI", 9F);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.Name = "FileSelectForm";
			this.UseWaitCursor = true;
			this.Load += new System.EventHandler(this.FileSelectForm_Load);
			this.ResumeLayout(false);

		}

		private void FileSelectForm_Load(object sender, EventArgs e)
		{

		}
	}
}
