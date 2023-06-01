using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Aspose.Cells;
using eBAControls.eBABaseForm;
using eBADB;
using eBAPI.Connection;

namespace AsposeCellTest
{
	class Program
	{
		static void Main(string[] args)
		{
			if (Debugger.IsAttached)
				CultureInfo.DefaultThreadCurrentUICulture = CultureInfo.GetCultureInfo("en-US");
			/*
			string filePath = @"C:\Users\akina\Downloads\" + "denemeAsposeOut" + ".xlsx";
			Workbook wb = new Workbook(@"C:\Users\akina\Downloads\asposeDeneme.xlsx");
			int sheetNumber = 0;
			Worksheet ws = wb.Worksheets[sheetNumber];
			//ws.Cells.Merge(6, 1, 2, 2);
			CreateExcel(ws); 
			AutoFitterOptions options = new AutoFitterOptions();

			options.AutoFitMergedCells = true;


			ws.AutoFitRows(options);
			ws.Cells.SetRowHeight(0, 12.75);
			ws.ViewType = ViewType.PageBreakPreview;
			using (Stream respStream = new MemoryStream())
			{
			
				wb.Save(respStream, Aspose.Cells.SaveFormat.Xlsx);
				respStream.Seek(0, SeekOrigin.Begin);
				WriteToResponse(respStream, "denemeAsposeOut" + ".xlsx");
				//string filePath = @"C:\Users\akina\Downloads\" + "denemeAsposeOut" + ".xlsx";
			
	
				wb.Save(respStream, Aspose.Cells.SaveFormat.Pdf);
				respStream.Seek(0, SeekOrigin.Begin);
				WriteToResponse(respStream, "denemeAsposeOut" + ".pdf");
	
				Process.Start(filePath);
			}*/
			Thread staThread = new Thread(new ThreadStart(StartUI));
			staThread.SetApartmentState(ApartmentState.STA);
			staThread.Start();
			staThread.Join();

		}

		private static void StartUI()
		{

			TextBox textbox = new TextBox();
			textbox.Text = @"10001004 10001010 10001012 10001015 10001017 10001018 10001020 10001021 10001022 10001023 10001026 10001029 10001031 10001032 10001035 10001036 10001037 10001038 10001041 10001044 10001045 10001046 10001051 10001052 10001053 10001055 10001059 10001062 10001065 10001066 10001069 10001070 10001074 10001018 11";
			CreateDocument(textbox, Aspose.Cells.SaveFormat.Xlsx);
		}

		public static void CreateDocument(TextBox textbox, Aspose.Cells.SaveFormat saveFormat)
		{
			List<byte[]> byteArrays = new List<byte[]>();
			string query = string.Format(@"SELECT DISTINCT FRM.txtRawMaterialId,txtRawMaterial,txtPtsDocumentNo
												FROM[dbo].[E_Pts015HammaddeKarti_Form] FRM
												INNER JOIN[dbo].[E_Pts015HammaddeKarti_MdlSpectInfo] SI WITH(NOLOCK) ON FRM.txtSpecInfoModalFormId = SI.ID
												INNER JOIN[dbo].[E_Pts015HammaddeKarti_MdlSpectManagement] SM WITH(NOLOCK) ON SI.txtSMModalFormId = SM.ID
												WHERE FRM.txtRawMaterialId IS NOT NULL AND SM.cbStatus = 1
												ORDER BY  FRM.txtRawMaterialId ASC");
			eBADBProvider db = CreateDatabaseProvider();
			SqlConnection sqlcon = (SqlConnection)db.Connection;
			SqlCommand cmd = new SqlCommand(query, sqlcon); //Spekt Karti Aktif olan Hammaddelerin, Id ve Valuelarini cekmek icin
			SqlDataAdapter adp = new SqlDataAdapter(cmd);
			DataTable dtActiveRawMaterials = new DataTable();
			adp.Fill(dtActiveRawMaterials);
			//Dictionary<string,Tuple<string, string,string>> materialInfos = new Dictionary<string, Tuple<string, string,string>>();//
			//List<Tuple<string, string, string>> existsMaterialInfos = new List<Tuple<string, string, string>>();
			Dictionary<string, (string materialId, string material, string documentNo)> materialInfos = new Dictionary<string, (string materialId, string material, string documentNo)>();//
			List<(string materialId,string material,string documentNo)> existsMaterialInfos = new List<(string materialId, string material, string documentNo)>();
			foreach (DataRow dr in dtActiveRawMaterials.Rows)
			{
				materialInfos.Add(dr["txtRawMaterialId"].ToString(), ( materialId: dr["txtRawMaterialId"].ToString(),material: dr["txtRawMaterial"].ToString(),documentNo: dr["txtPtsDocumentNo"].ToString()));
			}

			var materialCodes = textbox.Text.Trim().Split(' ');
			foreach (string materialCode in materialCodes)
			{
				if (materialInfos.ContainsKey(materialCode))
				{
					existsMaterialInfos.Add(materialInfos[materialCode]);
				}
			}

			foreach (var materialCode in existsMaterialInfos)
			{
				Workbook wb = new Workbook(@"C:\Users\akina\Downloads\asposeDeneme.xlsx");
				int sheetNumber = 0;
				Worksheet ws = wb.Worksheets[sheetNumber];
				//ws.Cells.Merge(6, 1, 2, 2);
				int rowCount = CreateExcel(ws,materialCode.materialId);
				AutoFitterOptions options = new AutoFitterOptions();

				options.AutoFitMergedCells = true;

				int pageCount = (int)Math.Ceiling(rowCount / 84.0);
				ws.AutoFitRows(options);
				ws.Cells.SetRowHeight(0, 12.75);
				ws.Cells[0, 16].PutValue(materialCode.Item3);//Dokuman No
				ws.Cells[1, 16].PutValue(DateTime.Now);//Yayin Tarihi
				ws.Cells[4, 16].PutValue(pageCount);//Sayfa No
				ws.ViewType = ViewType.PageBreakPreview;
				using (MemoryStream respStream = new MemoryStream())
				{

					wb.Save(respStream, saveFormat);
					respStream.Seek(0, SeekOrigin.Begin);
					byteArrays.Add(respStream.ToArray());
				}

			}
			MakeZip(byteArrays, "HammaddeSpektleri-"+DateTime.Now.ToString("dd-MM-yyyy"), existsMaterialInfos, saveFormat);

		}
		public static void MakeZip(List<byte[]> byteArrays, string fileName, List<(string materialId, string material, string documentNo)> materialCodes, Aspose.Cells.SaveFormat saveFormat)
		{
			using (var compressedFileStream = new MemoryStream())
			{
				//Create an archive and store the stream in memory.
				using (var zipArchive = new ZipArchive(compressedFileStream, ZipArchiveMode.Create, false))
				{
					int i = 0;
					foreach (var byteFile in byteArrays)
					{
						//Create a zip entry for each attachment
						//Bu kısımda doğru hatırlamamakla birlikte path verirsek o pathı kullanarak folderları oluşturuyordu sanırım method
						var zipEntry = zipArchive.CreateEntry(materialCodes[i].documentNo+" - "+ materialCodes[i].materialId + " - " + materialCodes[i].material + (saveFormat== Aspose.Cells.SaveFormat.Pdf ? ".pdf" : ".xlsx"));
						//Get the stream of the attachment
						using (var originalFileStream = new MemoryStream(byteFile))
						using (var zipEntryStream = zipEntry.Open())
						{
							//Copy the attachment stream to the zip entry stream
							originalFileStream.CopyTo(zipEntryStream);
						}
						i++;
					}
				}

				WriteToResponse(compressedFileStream.ToArray(), fileName + ".zip");
			}
		}
		public static  int CreateExcel(Worksheet ws,string materialCode)
		{
			eBADBProvider db = CreateDatabaseProvider();
			SqlConnection sqlcon = (SqlConnection)db.Connection;
			int row = 5;
			//GENEL OZELLIKLER
			row = GeneralSpecifications(ws, sqlcon, row, materialCode);
			Cell cell;
			Aspose.Cells.Style style;
			#region sectionRow

			ws.Cells.Merge(row, 0, 1, 18);
			GiveBorder(ws.Cells[row, 0].GetMergedRange());
			cell = ws.Cells[row, 0];

			// Get the style of the cell
			style = cell.GetStyle();

			// Change BackGroundColor
			style.ForegroundColor = System.Drawing.Color.FromArgb(160,0,98,137);
			style.Pattern = BackgroundType.Solid;

			// Apply the updated style to the cell
			cell.SetStyle(style);
			row++;
			#endregion
			#region SpectLabels
			//SPECT LABELS
			ws.Cells.Merge(row, 0, 1, 8);
			ws.Cells.Merge(row, 8, 1, 4);
			ws.Cells.Merge(row, 12, 1, 4);
			ws.Cells.Merge(row, 16, 1, 2);
			ws.Cells[row, 0].PutValue("KALİTE KRİTERLERİ");
			ws.Cells[row, 8].PutValue("KABUL LİMİTLERİ");
			ws.Cells[row, 12].PutValue("REFERANS/TEBLİĞ /STANDART");
			ws.Cells[row, 16].PutValue("ANALİZ YÖNTEMİ");
			//Give them borders
			GiveBorder(ws.Cells[row, 0].GetMergedRange());
			GiveBorder(ws.Cells[row, 8].GetMergedRange());
			GiveBorder(ws.Cells[row, 12].GetMergedRange());
			GiveBorder(ws.Cells[row, 16].GetMergedRange());
			//make them bold
			MakeBold(ws.Cells[row, 0]);
			MakeBold(ws.Cells[row, 8]);
			MakeBold(ws.Cells[row, 12]);
			MakeBold(ws.Cells[row, 16]);
			//Alignment Of Titles
			changeAlignment(ws.Cells[row, 0], TextAlignmentType.Center, TextAlignmentType.Center);
			changeAlignment(ws.Cells[row, 8], TextAlignmentType.Center, TextAlignmentType.Center);
			changeAlignment(ws.Cells[row, 12], TextAlignmentType.Center, TextAlignmentType.Center);
			changeAlignment(ws.Cells[row, 16], TextAlignmentType.Center, TextAlignmentType.Center);
			row++;
			#endregion
			row = PutSpectTitles( ws,  sqlcon,  row, materialCode);
			#region GKK Ozel Alan
			//GKK Özel Alanlar
			string query = string.Format(@"SELECT Vls.VALUE, txtAcceptableCriteria
												FROM [dbo].[TbPts015SpectDetails] 
												INNER JOIN TbPts000LookupValues VLS ON VLS.ID = mtlQualityCriteria
												WHERE mtlSpectTitle = '9653' AND txtRawMaterialId = '{0}' AND mtlQualityCriteria != 9659", materialCode);
			SqlCommand cmd = new SqlCommand(query, sqlcon);
			SqlDataAdapter adp = new SqlDataAdapter(cmd);
			DataTable dt = new DataTable();
			adp.Fill(dt);
			if (dt.Rows.Count == 0)
			{
				goto noGKK;
			}
			foreach (DataRow dr in dt.Rows)
			{

				ws.Cells.Merge(row, 0, 1, 3);
				ws.Cells.Merge(row, 3, 1, 15);
				GiveBorder(ws.Cells[row, 0].GetMergedRange());
				GiveBorder(ws.Cells[row, 3].GetMergedRange());
				ws.Cells[row, 0].PutValue(dr[0]);
				ws.Cells[row, 3].PutValue(dr[1]);
				//ws.Cells[row, 3].SetStyle(GiveBorder(ws.Cells[row, 3].GetStyle()));
				MakeBold(ws.Cells[row, 0]);
				MakeTextWrap(ws.Cells[row, 3]);
				changeAlignment(ws.Cells[row, 0], TextAlignmentType.Center, TextAlignmentType.Center);
				row++;

			}
			ws.Cells.Merge(row, 0, 1, 18);
			GiveBorder(ws.Cells[row, 0].GetMergedRange());
			cell = ws.Cells[row, 0];
			// Get the style of the cell
			style = cell.GetStyle();
			// Set the font to bold
			style.ForegroundColor = System.Drawing.Color.FromArgb(160, 0, 98, 137);
			style.Pattern = BackgroundType.Solid;
			// Apply the updated style to the cell
			cell.SetStyle(style);
			row++;
			noGKK:
			#endregion
			row = CreateAlerjenTitles(ws,row);
			#region Alerjenler
			query = string.Format(@"SELECT VALUE AS 'ALERJENLER'
												 ,CASE WHEN cbInsideRawMaterial = 1 THEN 'X' ELSE '' END AS 'Hammaddenin İçinde'
												 ,CASE WHEN cbSameProductionLine = 1 THEN 'X' ELSE '' END AS 'Aynı Üretim Hattında Çapraz Bulaşma'
												 ,CASE WHEN cbSameFactory = 1 THEN 'X' ELSE '' END AS 'Aynı Fabrikada Çapraz Bulaşma'
												 ,'' AS Aciklama
												 FROM[dbo].[TbPts015AllergenInfo] FRM
												 INNER JOIN[dbo].[TbPts000LookupValues] LV WITH(NOLOCK) ON LV.ID = FRM.mtlAllergenInfo
												 WHERE txtRawMaterialId = '{0}'
												 ORDER BY mtlAllergenInfo ASC", materialCode);
			cmd = new SqlCommand(query, sqlcon);
			adp = new SqlDataAdapter(cmd);
			DataTable dtAlerjen = new DataTable();
			adp.Fill(dtAlerjen);
			foreach (DataRow dr in dtAlerjen.Rows)
			{

				ws.Cells.Merge(row, 0 , 1, 3);
				ws.Cells.Merge(row, 3 , 1, 5);
				ws.Cells.Merge(row, 8 , 1, 5);
				ws.Cells.Merge(row, 13, 1, 5);
				//ws.Cells.Merge(row, 14, 1, 4);
				GiveBorder(ws.Cells[row, 0 ].GetMergedRange());
				GiveBorder(ws.Cells[row, 3 ].GetMergedRange());
				GiveBorder(ws.Cells[row, 8 ].GetMergedRange());
				GiveBorder(ws.Cells[row, 13].GetMergedRange());
				//GiveBorder(ws.Cells[row, 14].GetMergedRange());
				ws.Cells[row, 0 ].PutValue(dr[0]);
				ws.Cells[row, 3 ].PutValue(dr[1]);
				ws.Cells[row, 8 ].PutValue(dr[2]);
				ws.Cells[row, 13].PutValue(dr[3]);
				//ws.Cells[row, 14].PutValue(dr[4]);
				//ws.Cells[row, 3].SetStyle(GiveBorder(ws.Cells[row, 3].GetStyle()));
				MakeBold(ws.Cells[row, 0]);
				//MakeTextWrap(ws.Cells[row, 3]);
				changeAlignment(ws.Cells[row, 0 ], TextAlignmentType.Center, TextAlignmentType.Center);
				changeAlignment(ws.Cells[row, 3 ], TextAlignmentType.Center, TextAlignmentType.Center);
				changeAlignment(ws.Cells[row, 8 ], TextAlignmentType.Center, TextAlignmentType.Center);
				changeAlignment(ws.Cells[row, 13], TextAlignmentType.Center, TextAlignmentType.Center);
				//changeAlignment(ws.Cells[row, 14], TextAlignmentType.Center, TextAlignmentType.Center);
				row++;

			}
			ws.Cells.Merge(row, 0, 1, 18);
			GiveBorder(ws.Cells[row, 0].GetMergedRange());
			ws.Cells[row, 0].PutValue("Bu doküman PATI Society platformunda elektronik olarak hazırlanmıştır ve onaylanmıştır");
			MakeBold(ws.Cells[row, 0]);
			#endregion
			return row;
		}

		public static int GeneralSpecifications(Worksheet ws, SqlConnection sqlcon, int row, string materialCode)
		{
			string query = string.Format(@"SELECT VALUE , txtAcceptableCriteria
									FROM[dbo].[TbPts015SpectDetails] FRM
									INNER JOIN[dbo].[TbPts000LookupValues] LV WITH(NOLOCK) ON LV.ID = FRM.mtlQualityCriteria
									WHERE txtRawMaterialId = '{0}' AND mtlSpectTitle = '{1}' AND mtlQualityCriteria <> 9659
									ORDER BY LV.ID ", materialCode, "8811");
			SqlCommand cmd = new SqlCommand(query, sqlcon);
			SqlDataAdapter adp = new SqlDataAdapter(cmd);
			DataTable dtGeneralAndGKKSpecialFields = new DataTable();
			adp.Fill(dtGeneralAndGKKSpecialFields);
			foreach (DataRow dr in dtGeneralAndGKKSpecialFields.Rows)
			{
				//Merge Cells
				ws.Cells.Merge(row, 0, 1, 3);
				ws.Cells.Merge(row, 3, 1, 15);
				//Create Borders for merged cells
				GiveBorder(ws.Cells[row, 0].GetMergedRange());
				GiveBorder(ws.Cells[row, 3].GetMergedRange());
				//Put Values
				ws.Cells[row, 0].PutValue(dr[0]);
				ws.Cells[row, 3].PutValue(dr[1]);
				//Make Title Bold
				MakeBold(ws.Cells[row, 0]);
				row++;
			}
			return row;
		}

		public static int PutSpectTitles(Worksheet ws, SqlConnection sqlcon, int row, string materialCode)
		{
			Cell cell;
			Aspose.Cells.Style style;
			List<KeyValuePair<int, string>> spectTitleIdDefault = new List<KeyValuePair<int, string>>() { //Siralama bu sekilde olmali
				// new KeyValuePair<int, string>(8811, "Genel Özellikler")
				new KeyValuePair<int, string>(8800, "Duyusal/Fiziksel Özellikler")
				,new KeyValuePair<int, string>(8801, "Fiziksel Özellikler")
				,new KeyValuePair<int, string>(8805, "Analitik Değerler")
				,new KeyValuePair<int, string>(8802, "Kimyasal Özellikler")
				,new KeyValuePair<int, string>(8810, "Aminoasit İçeriği (g/100g)")
				,new KeyValuePair<int, string>(8803, "X-Ray Kontrolleri")
				,new KeyValuePair<int, string>(8804, "Metal Dedektör (Doypack Ambalaj İçin)")
				,new KeyValuePair<int, string>(8807, "Bulaşanlar")
				,new KeyValuePair<int, string>(8806, "Mikrobiyolojik Özellikler")
				,new KeyValuePair<int, string>(8808, "Besin Enerji")
			//	,new KeyValuePair<int, string>(9653, "GKK Özel Alanlar")
			//	,new KeyValuePair<int, string>(8809, "Alerjen Bilgileri") 
			};
			foreach (var titles in spectTitleIdDefault)
			{
				string query = string.Format(@"SELECT Vls.VALUE, txtAcceptableCriteria,txtReferanceNotificationStandard,txtAnalysisMethod
												FROM [dbo].[TbPts015SpectDetails] 
												INNER JOIN TbPts000LookupValues VLS ON VLS.ID = mtlQualityCriteria
												WHERE mtlSpectTitle = '{0}' AND txtRawMaterialId = '{1}'", titles.Key, materialCode);
				SqlCommand cmd = new SqlCommand(query, sqlcon);
				SqlDataAdapter adp = new SqlDataAdapter(cmd);
				DataTable data = new DataTable();
				adp.Fill(data);
				int dataCount = data.Rows.Count;
				if (dataCount == 0) continue;
				//En Solda Spect Titlein Mergelenmis bir sekilde yazmasi icin
				ws.Cells.Merge(row, 0, dataCount, 3);
				GiveBorder(ws.Cells[row, 0].GetMergedRange());
				ws.Cells[row, 0].PutValue(titles.Value);
				changeAlignment(ws.Cells[row, 0], TextAlignmentType.Center, TextAlignmentType.Center);
				MakeBold(ws.Cells[row, 0]);

				foreach (DataRow dr in data.Rows)
				{
					//Merge Cells
					ws.Cells.Merge(row, 3, 1, 5);
					ws.Cells.Merge(row, 8, 1, 4);
					ws.Cells.Merge(row, 12, 1, 4);
					ws.Cells.Merge(row, 16, 1, 2);
					//Create Borders
					GiveBorder(ws.Cells[row, 3].GetMergedRange());
					GiveBorder(ws.Cells[row, 8].GetMergedRange());
					GiveBorder(ws.Cells[row, 12].GetMergedRange());
					GiveBorder(ws.Cells[row, 16].GetMergedRange());
					//Put Values
					ws.Cells[row, 3].PutValue(dr[0]);
					ws.Cells[row, 8].PutValue(dr[1]);
					ws.Cells[row, 12].PutValue(dr[2]);
					ws.Cells[row, 16].PutValue(dr[3]);
					//Make Title Bold
					MakeBold(ws.Cells[row, 3]);
					//Wrap the text if it exceed the width
					MakeTextWrap(ws.Cells[row, 3]);
					MakeTextWrap(ws.Cells[row, 8]);
					MakeTextWrap(ws.Cells[row, 12]);
					MakeTextWrap(ws.Cells[row, 16]);
					changeAlignment(ws.Cells[row, 3], TextAlignmentType.Center, TextAlignmentType.Left);
					row++;

				}
				ws.Cells.Merge(row, 0, 1, 18);
				GiveBorder(ws.Cells[row, 0].GetMergedRange());
				cell = ws.Cells[row, 0];
				// Get the style of the cell
				style = cell.GetStyle();
				// Set the font to bold
				style.ForegroundColor = System.Drawing.Color.FromArgb(160, 0, 98, 137);
				style.Pattern = BackgroundType.Solid;
				// Apply the updated style to the cell
				cell.SetStyle(style);
				row++;

			}
			return row;
		}

		public static int CreateAlerjenTitles(Worksheet ws,int row)
		{
			
			ws.Cells.Merge(row, 0, 1, 18);
			ws.Cells[row, 0].PutValue("ALERJEN BİLGİLERİ");
			GiveBorder(ws.Cells[row, 0].GetMergedRange());
			MakeBold(ws.Cells[row, 0]);
			row++;
			//Merge Cells
			ws.Cells.Merge(row, 0 , 1, 3);
			ws.Cells.Merge(row, 3 , 1, 5);
			ws.Cells.Merge(row, 8 , 1, 5);
			ws.Cells.Merge(row, 13, 1, 5);
			//ws.Cells.Merge(row, 14, 1, 4);
			//Put Values
			ws.Cells[row, 0 ].PutValue("ALERJENLER");
			ws.Cells[row, 3 ].PutValue("VAR");
			ws.Cells[row, 8 ].PutValue("YOK");
			ws.Cells[row, 13].PutValue("ESER MİKTARDA");
			//ws.Cells[row, 14].PutValue("AÇIKLAMA");
			//Give them borders
			GiveBorder(ws.Cells[row, 0 ].GetMergedRange());
			GiveBorder(ws.Cells[row, 3 ].GetMergedRange());
			GiveBorder(ws.Cells[row, 8 ].GetMergedRange());
			GiveBorder(ws.Cells[row, 13].GetMergedRange());
			//GiveBorder(ws.Cells[row, 14].GetMergedRange());
			//make them bold
			MakeBold(ws.Cells[row, 0 ] );
			MakeBold(ws.Cells[row, 3 ]);
			MakeBold(ws.Cells[row, 8 ]);
			MakeBold(ws.Cells[row, 13]);
			//MakeBold(ws.Cells[row, 14]);
			//Alignment Of Titles
			changeAlignment(ws.Cells[row, 0 ], TextAlignmentType.Center, TextAlignmentType.Center);
			changeAlignment(ws.Cells[row, 3 ], TextAlignmentType.Center, TextAlignmentType.Center);
			changeAlignment(ws.Cells[row, 8 ], TextAlignmentType.Center, TextAlignmentType.Center);
			changeAlignment(ws.Cells[row, 13], TextAlignmentType.Center, TextAlignmentType.Center);
			//changeAlignment(ws.Cells[row, 14], TextAlignmentType.Center, TextAlignmentType.Center);
			row++;
			return row;
		}

		public static Aspose.Cells.Style GiveBorder(Aspose.Cells.Style Labelstyle)
		{
			//Set border for the cell
			Labelstyle.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, System.Drawing.Color.Gray);
			Labelstyle.SetBorder(BorderType.TopBorder, CellBorderType.Thin, System.Drawing.Color.Gray);
			Labelstyle.SetBorder(BorderType.RightBorder, CellBorderType.Thin, System.Drawing.Color.Gray);
			Labelstyle.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, System.Drawing.Color.Gray);
			return Labelstyle;
		}

		public static void MakeBold(Cell cell)
		{
			Aspose.Cells.Style style = cell.GetStyle();
			style.Font.IsBold = true;
			cell.SetStyle(GiveBorder(style));
		}

		public static void GiveBorder(Range range)
		{
			//Set border for the merged cells
			range.SetOutlineBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Gray);
			range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Gray);
			range.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Gray);
			range.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Gray);
		}

		public static void MakeTextWrap(Cell cell)
		{
			Aspose.Cells.Style style = cell.GetStyle();
			style.IsTextWrapped = true;
			cell.SetStyle(style);
		}

		public static void changeAlignment(Cell cell,TextAlignmentType vAlign, TextAlignmentType hAlign)
		{
			Aspose.Cells.Style stl = cell.GetStyle();
			stl.VerticalAlignment = vAlign;
			stl.HorizontalAlignment = hAlign;
			cell.SetStyle(GiveBorder(stl));
		}





		public static void ShowMessageBox(string stringg, eBAMessageBoxType type)
		{
			MessageBox.Show(stringg);
		}
		public static void WriteToResponse(Stream stream, string name)
		{
			var fileStream = File.Create(@"C:\Users\akina\Downloads\" + name);
			stream.CopyTo(fileStream);
		}
		public static void WriteToResponse(Byte[] byteArray, string name)
		{
			System.IO.MemoryStream stream = new MemoryStream(byteArray);
			var fileStream = File.Create(@"C:\Users\akina\Downloads\" + name);
			stream.CopyTo(fileStream);
		}
		public static void ShowMessageBar(string s, int i, ShowInfoBarType t)
		{
			MessageBox.Show("Bar" + s);
		}

		public static eBADBProvider CreateDatabaseProvider()
		{
			return new eBADBProvider(ServerType.SqlServer, "System.Data.SqlClient", @"Data Source = *********; " +
				"Initial Catalog=EBA;" +
				"User id=*******;" +
				@"Password=********;");
		}


		public static void ShowMessageBox(string obj)
		{
			MessageBox.Show(obj);
		}

	}


	public class Excel
	{
		/// <summary>
		/// rgb renklerden kirmizi icin in deger 0-255 orasinda olmali
		/// </summary>
		public int Red { get; set; }
		/// <summary>
		/// rgb renklerden Yesil icin in deger 0-255 orasinda olmali
		/// </summary>
		public int Green { get; set; }
		/// <summary>
		/// rgb renklerden Mavi icin in deger 0-255 orasinda olmali
		/// </summary>
		public int Blue { get; set; }
		/// <summary>
		/// Exceldeki Basliklar isin style 
		/// </summary>
		public Aspose.Cells.Style Labelstyle { get; set; }
		/// <summary>
		/// Exceldeki celler icin style
		/// </summary>
		public Aspose.Cells.Style CellStyle { get; set; }
		/// <summary>
		/// Default Constructor
		/// </summary>
		public Excel()
		{
			Red = 100;
			Green = 180;
			Blue = 250;
			Labelstyle = new CellsFactory().CreateStyle();
			Labelstyle.ForegroundColor = System.Drawing.Color.FromArgb(Red, Green, Blue);
			Labelstyle.Pattern = BackgroundType.Solid;
			Labelstyle.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, System.Drawing.Color.Black);
			Labelstyle.SetBorder(BorderType.TopBorder, CellBorderType.Thin, System.Drawing.Color.Black);
			Labelstyle.SetBorder(BorderType.RightBorder, CellBorderType.Thin, System.Drawing.Color.Black);
			Labelstyle.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, System.Drawing.Color.Black);
			CellStyle = new CellsFactory().CreateStyle();
		}
		/// <summary>
		/// Basliklarin renklerini degistirmek icin kullanilacak constructor
		/// </summary>
		/// <param name="red">rgb renklerden kirmizi icin in deger 0-255 orasinda olmali</param>
		/// <param name="green">rgb renklerden Yesil icin in deger 0-255 orasinda olmali</param>
		/// <param name="blue">rgb renklerden Mavi icin in deger 0-255 orasinda olmali</param>
		public Excel(int red, int green, int blue)
		{
			this.Red = red;
			this.Green = green;
			this.Blue = blue;
		}
		/// <summary>
		/// Basliklarin renkleri ve stilini degistirmek icin kullanilmasi gereken constructor
		/// </summary>
		/// <param name="red">rgb renklerden kirmizi icin in deger 0-255 orasinda olmali</param>
		/// <param name="green">rgb renklerden Yesil icin in deger 0-255 orasinda olmali</param>
		/// <param name="blue">rgb renklerden Mavi icin in deger 0-255 orasinda olmali</param>
		/// <param name="labelstyle">Aspose cell Stil elementi</param>
		public Excel(int red, int green, int blue, Aspose.Cells.Style labelstyle)
		{
			this.Red = red;
			this.Green = green;
			this.Blue = blue;
			this.Labelstyle = labelstyle;
		}
		/// <summary>
		/// Basliklarin renkleri ve stilini degistirmek icin kullanilmasi gereken constructor
		/// </summary>
		/// <param name="red">rgb renklerden kirmizi icin in deger 0-255 orasinda olmali</param>
		/// <param name="green">rgb renklerden Yesil icin in deger 0-255 orasinda olmali</param>
		/// <param name="blue">rgb renklerden Mavi icin in deger 0-255 orasinda olmali</param>
		/// <param name="labelstyle">Aspose cell Stil elementi</param>
		/// <param name="cellStyle">Aspose cell Stil elementi</param>
		public Excel(int red, int green, int blue, Aspose.Cells.Style labelstyle, Aspose.Cells.Style cellStyle)
		{
			this.Red = red;
			this.Green = green;
			this.Blue = blue;
			this.Labelstyle = labelstyle;
			this.CellStyle = cellStyle;
		}
		/// <summary>
		/// Excel Indirme methodu
		/// </summary>
		/// <param name="parameters">Her bir tuple 1 sheeti simgeler, 1. eleman sheet adi iken 2. eleman verileri tutan datatable idir(DataTablein kolon isimleri excelde gosterilir)</param>
		/// <param name="outputName">Indirilecek dosyanin uzanti dahil edilmemis haliyle adi</param>
		/// <param name="WriteToResponse">Ebada kullanilan dosya indirme methodu</param>
		public void Download(List<Tuple<string, DataTable>> parameters, string outputName, Action<Stream, string> WriteToResponse)
		{
			Workbook wb = new Workbook();
			int sheetNumber = 0;
			foreach (Tuple<string, DataTable> parameter in parameters)
			{
				string sheetName = parameter.Item1;
				DataTable data = parameter.Item2;
				if (sheetNumber > 0)
				{
					wb.Worksheets.Add();
				}
				Worksheet ws = wb.Worksheets[sheetNumber];
				ws.Name = sheetName;
				for (int col = 0; col < data.Columns.Count; col++)
				{
					ws.Cells[0, col].PutValue(data.Columns[col].ColumnName);
					ws.Cells[0, col].SetStyle(Labelstyle);
				}
				for (int row = 0; row < data.Rows.Count; row++)
				{
					for (int col = 0; col < data.Columns.Count; col++)
					{
						ws.Cells[row + 1, col].PutValue(data.Rows[row][col]);
						if (data.Columns[col].DataType == typeof(DateTime) || data.Columns[col].DataType == typeof(DateTime))
						{
							CellStyle.Number = 14;
						}
						else
						{
							CellStyle.Number = 0;
						}
						ws.Cells[row + 1, col].SetStyle(CellStyle);

					}
				}
				ws.AutoFitColumns();
				sheetNumber++;
			}
			using (Stream respStream = new MemoryStream())
			{
				wb.Save(respStream, Aspose.Cells.SaveFormat.Xlsx);
				respStream.Seek(0, SeekOrigin.Begin);
				WriteToResponse(respStream, outputName + ".xlsx");
			}
		}
	}
}
